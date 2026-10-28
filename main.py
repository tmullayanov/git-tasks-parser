"""Export commits from Git repositories to one Excel workbook."""

import argparse
import os
import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime
from collections.abc import Iterator
from pathlib import Path
from typing import Any, BinaryIO
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell

HEADERS = ("Commit hash", "Issue key", "author email", "Commit date", "repository name")
MAX_EXCEL_ROWS = 1_048_576
DATE_FORMAT = "yyyy-mm-dd hh:mm"
ISSUE_KEY = re.compile(r"(?<![\w-])[A-Z][A-Z0-9_]+-[1-9][0-9]*(?![\w-])")
HASH = re.compile(rb"[0-9a-f]{40}(?:[0-9a-f]{24})?")


class ExportError(Exception):
    """An expected error that can be reported without a traceback."""


@dataclass
class Repository:
    path: Path
    name: str | None = None


class RepoAction(argparse.Action):
    def __call__(self, parser: argparse.ArgumentParser, namespace: argparse.Namespace, values: str, option_string: str | None = None) -> None:
        repositories = getattr(namespace, self.dest, None)
        if repositories is None:
            repositories = []
            setattr(namespace, self.dest, repositories)
        repositories.append(Repository(Path(values)))
        namespace._renameable = True


class RepositoryNameAction(argparse.Action):
    def __call__(self, parser: argparse.ArgumentParser, namespace: argparse.Namespace, values: str, option_string: str | None = None) -> None:
        if not getattr(namespace, "_renameable", False):
            parser.error("--repository-name must immediately follow its --repo")
        if not values.strip():
            parser.error("--repository-name cannot be empty")
        namespace.repositories[-1].name = values
        namespace._renameable = False


class OtherOptionAction(argparse.Action):
    def __call__(self, parser: argparse.ArgumentParser, namespace: argparse.Namespace, values: str, option_string: str | None = None) -> None:
        setattr(namespace, self.dest, values)
        namespace._renameable = False


def git(repo: Path, *args: str) -> subprocess.CompletedProcess[bytes]:
    try:
        result = subprocess.run(
            ["git", "-C", str(repo), *args], capture_output=True, check=False
        )
    except OSError as exc:
        raise ExportError(f"Cannot run git: {exc}") from exc
    if result.returncode:
        detail = result.stderr.decode("utf-8", errors="replace").strip()
        raise ExportError(f"Git error ({' '.join(args)}): {detail or result.returncode}")
    return result


def resolve_branch(repo: Path, name: str) -> str:
    if not repo.is_dir():
        raise ExportError(f"Repository directory does not exist: {repo}")
    git(repo, "rev-parse", "--git-dir")

    ref = name if name.startswith("refs/") else f"refs/heads/{name}"
    if not (ref.startswith("refs/heads/") or ref.startswith("refs/remotes/")):
        raise ExportError("Branch must be a local or remote-tracking branch")
    try:
        git(repo, "check-ref-format", ref)
    except ExportError as exc:
        raise ExportError(f"Invalid branch name: {name}") from exc
    try:
        commit = git(repo, "show-ref", "--verify", "--hash", ref).stdout.strip().decode("ascii")
    except ExportError as exc:
        raise ExportError(f"Branch does not exist: {name}") from exc
    kind = git(repo, "cat-file", "-t", commit).stdout.strip()
    if kind != b"commit":
        raise ExportError(f"Branch does not point to a commit: {name}")
    return commit


def split_nul_fields(stream: BinaryIO) -> Iterator[bytes]:
    pending = b""
    while chunk := stream.read(64 * 1024):
        parts = (pending + chunk).split(b"\0")
        pending = parts.pop()
        yield from parts
    if pending not in (b"", b"\n"):
        raise ExportError("Incomplete Git log output")


def commits(repo: Path, commit: str) -> Iterator[tuple[str, str, datetime, str]]:
    command = [
        "git", "-C", str(repo), "log", "--topo-order", "--no-use-mailmap",
        "--no-patch", "--encoding=UTF-8", "--format=%H%x00%ae%x00%cI%x00%B%x00", commit,
    ]
    try:
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except OSError as exc:
        raise ExportError(f"Cannot run git: {exc}") from exc

    try:
        assert process.stdout is not None
        fields = []
        for field in split_nul_fields(process.stdout):
            fields.append(field)
            if len(fields) != 4:
                continue
            raw_hash, raw_email, raw_date, raw_message = fields
            fields = []
            raw_hash = raw_hash.lstrip(b"\r\n")
            if not HASH.fullmatch(raw_hash):
                raise ExportError("Unexpected Git log output: invalid commit hash")
            try:
                commit_date = datetime.fromisoformat(raw_date.decode("ascii"))
            except (UnicodeError, ValueError) as exc:
                raise ExportError("Unexpected Git log output: invalid commit date") from exc
            if commit_date.tzinfo is None:
                raise ExportError("Unexpected Git log output: commit date has no time zone")
            yield (
                raw_hash.decode("ascii"),
                raw_email.decode("utf-8", errors="replace"),
                commit_date,
                raw_message.decode("utf-8", errors="replace"),
            )
        if fields:
            raise ExportError("Incomplete Git log output")
        assert process.stderr is not None
        detail = process.stderr.read().decode("utf-8", errors="replace").strip()
        if process.wait():
            raise ExportError(f"Git log failed: {detail or process.returncode}")
    finally:
        if process.poll() is None:
            process.terminate()
        process.communicate()


def text_cell(sheet: Any, value: str) -> Any:
    cell = WriteOnlyCell(sheet, value=value)
    cell.data_type = "s"
    return cell


def export(repositories: list[Repository], branch: str, timezone: str, output: Path) -> int:
    if not repositories:
        raise ExportError("At least one repository is required")
    try:
        zone = ZoneInfo(timezone)
    except (ZoneInfoNotFoundError, ValueError) as exc:
        raise ExportError(f"Unknown IANA time zone: {timezone}") from exc
    if output.suffix.lower() != ".xlsx":
        raise ExportError("Output file must have the .xlsx extension")
    if output.exists():
        raise ExportError(f"Output file already exists: {output}")

    selected = []
    names = set()
    for repository in repositories:
        path = repository.path.resolve()
        name = repository.name if repository.name is not None else path.name
        if not name.strip():
            raise ExportError(f"Repository name cannot be empty: {path}")
        if name.casefold() in names:
            raise ExportError(f"Duplicate repository name: {name}")
        names.add(name.casefold())
        try:
            commit = resolve_branch(path, branch)
        except ExportError as exc:
            raise ExportError(f"{path}: {exc}") from exc
        selected.append((path, name, commit))

    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("Commits")
    sheet.append([text_cell(sheet, header) for header in HEADERS])
    row_count = 1
    temp_path = None
    try:
        for path, name, commit in selected:
            try:
                for commit_hash, email, commit_date, message in commits(path, commit):
                    local_date = commit_date.astimezone(zone).replace(tzinfo=None)
                    keys = dict.fromkeys(ISSUE_KEY.findall(message))
                    for key in keys or {"": None}:
                        if row_count >= MAX_EXCEL_ROWS:
                            raise ExportError(f"Excel sheet limit exceeded ({MAX_EXCEL_ROWS} rows)")
                        date_cell = WriteOnlyCell(sheet, value=local_date)
                        date_cell.number_format = DATE_FORMAT
                        sheet.append([
                            text_cell(sheet, commit_hash), text_cell(sheet, key),
                            text_cell(sheet, email), date_cell, text_cell(sheet, name),
                        ])
                        row_count += 1
            except ExportError as exc:
                raise ExportError(f"{path}: {exc}") from exc

        fd, filename = tempfile.mkstemp(prefix=".git-tasks-parser-", suffix=".xlsx", dir=output.parent)
        os.close(fd)
        temp_path = Path(filename)
        workbook.save(temp_path)
        try:
            os.link(temp_path, output)
        except FileExistsError as exc:
            raise ExportError(f"Output file already exists: {output}") from exc
        return row_count - 1
    except ExportError:
        raise
    except Exception as exc:
        raise ExportError(f"Cannot write Excel file: {exc}") from exc
    finally:
        workbook.close()
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", dest="repositories", action=RepoAction, required=True, metavar="PATH", help="Local Git repository; repeat for several")
    parser.add_argument("--repository-name", action=RepositoryNameAction, metavar="NAME", help="Name for the immediately preceding --repo")
    parser.add_argument("--main-branch", action=OtherOptionAction, required=True, metavar="REF", help="Local branch name or full remote-tracking ref")
    parser.add_argument("--timezone", action=OtherOptionAction, required=True, metavar="ZONE", help="IANA time zone for commit dates")
    parser.add_argument("--output", action=OtherOptionAction, type=Path, required=True, help="New .xlsx file")
    args = parser.parse_args()
    try:
        count = export(args.repositories, args.main_branch, args.timezone, args.output)
    except ExportError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    print(f"Exported {count} rows to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
