"""Export commits reachable from one Git branch to an Excel workbook."""

import argparse
import os
import re
import subprocess
import sys
import tempfile
from collections.abc import Iterator
from pathlib import Path
from typing import Any, BinaryIO

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell

HEADERS = ("Commit hash", "Issue key", "author email")
MAX_EXCEL_ROWS = 1_048_576
ISSUE_KEY = re.compile(r"(?<![\w-])[A-Z][A-Z0-9_]+-[1-9][0-9]*(?![\w-])")
HASH = re.compile(rb"[0-9a-f]{40}(?:[0-9a-f]{24})?")


class ExportError(Exception):
    """An expected error that can be reported without a traceback."""


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


def commits(repo: Path, commit: str) -> Iterator[tuple[str, str, str]]:
    command = [
        "git", "-C", str(repo), "log", "--topo-order", "--no-use-mailmap",
        "--no-patch", "--format=%H%x00%ae%x00%B%x00", commit,
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
            if len(fields) != 3:
                continue
            raw_hash, raw_email, raw_message = fields
            fields = []
            raw_hash = raw_hash.lstrip(b"\r\n")
            if not HASH.fullmatch(raw_hash):
                raise ExportError("Unexpected Git log output: invalid commit hash")
            yield (
                raw_hash.decode("ascii"),
                raw_email.decode("utf-8", errors="replace"),
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


def export(repo: Path, branch: str, output: Path) -> int:
    commit = resolve_branch(repo, branch)
    if output.suffix.lower() != ".xlsx":
        raise ExportError("Output file must have the .xlsx extension")
    if output.exists():
        raise ExportError(f"Output file already exists: {output}")

    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("Commits")
    sheet.append([text_cell(sheet, header) for header in HEADERS])
    row_count = 1
    temp_path = None
    try:
        for commit_hash, email, message in commits(repo, commit):
            keys = dict.fromkeys(ISSUE_KEY.findall(message))
            for key in keys or {"": None}:
                if row_count >= MAX_EXCEL_ROWS:
                    raise ExportError(f"Excel sheet limit exceeded ({MAX_EXCEL_ROWS} rows)")
                sheet.append([text_cell(sheet, value) for value in (commit_hash, key, email)])
                row_count += 1

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
    parser.add_argument("--repo", type=Path, required=True, help="Local Git repository")
    parser.add_argument("--main-branch", required=True, metavar="REF", help="Local branch name or full remote-tracking ref")
    parser.add_argument("--output", type=Path, required=True, help="New .xlsx file")
    args = parser.parse_args()
    try:
        count = export(args.repo, args.main_branch, args.output)
    except ExportError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    print(f"Exported {count} rows to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
