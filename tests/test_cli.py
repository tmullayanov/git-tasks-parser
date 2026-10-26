import os
import subprocess
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from main import ExportError, export


class ExportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parents[1])
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.repo = self.root / "repo"
        self.repo.mkdir()
        self.env = os.environ.copy()
        self.env.update({
            "GIT_AUTHOR_NAME": "Author",
            "GIT_AUTHOR_EMAIL": "=1+1",
            "GIT_COMMITTER_NAME": "Committer",
            "GIT_COMMITTER_EMAIL": "committer@example.com",
        })
        self.git("init", "-b", "main")

    def git(self, *args):
        result = subprocess.run(
            ["git", "-C", str(self.repo), *args],
            env=self.env, capture_output=True, text=True, check=True,
        )
        return result.stdout.strip()

    def commit(self, message, body=None):
        args = ["commit", "--allow-empty", "-m", message]
        if body is not None:
            args.extend(["-m", body])
        return self.git(*args)

    def test_merge_keys_author_and_text_cells(self):
        self.commit("No issue here")
        root_hash = self.git("rev-parse", "HEAD")
        self.git("switch", "-c", "feature")
        self.commit("Feature", "PROJ-123 OPS-7 PROJ-123 proj-2 xPROJ-6 PROJ-0")
        feature_hash = self.git("rev-parse", "HEAD")
        self.git("switch", "main")
        self.commit("MAIN_2-4")
        main_hash = self.git("rev-parse", "HEAD")
        self.git("merge", "--no-ff", "feature", "-m", "Merge feature")
        merge_hash = self.git("rev-parse", "HEAD")
        self.git("switch", "-c", "unmerged")
        self.commit("HIDDEN-1")
        self.git("switch", "main")

        output = self.root / "report.xlsx"
        self.assertEqual(export(self.repo, "main", output), 5)
        workbook = load_workbook(output)
        self.assertEqual(workbook.sheetnames, ["Commits"])
        rows = list(workbook.active.values)
        self.assertEqual(rows[0], ("Commit hash", "Issue key", "author email"))
        self.assertEqual(rows[1], (merge_hash, None, "=1+1"))
        self.assertEqual(
            {(hash_, key) for hash_, key, _ in rows[1:]},
            {
                (merge_hash, None), (main_hash, "MAIN_2-4"),
                (feature_hash, "PROJ-123"), (feature_hash, "OPS-7"),
                (root_hash, None),
            },
        )
        self.assertEqual(rows[-1][0], root_hash)
        self.assertTrue(all(email == "=1+1" for _, _, email in rows[1:]))
        self.assertEqual(workbook.active["C2"].data_type, "s")

        with self.assertRaisesRegex(ExportError, "already exists"):
            export(self.repo, "main", output)
        self.assertEqual(list(load_workbook(output).active.values), rows)

    def test_invalid_branch_and_detached_head(self):
        self.commit("ROOT-1")
        self.git("update-ref", "refs/remotes/origin/main", "HEAD")
        self.git("checkout", "--detach")
        output = self.root / "report.xlsx"
        self.assertEqual(export(self.repo, "main", output), 1)
        self.assertEqual(
            export(self.repo, "refs/remotes/origin/main", self.root / "remote.xlsx"), 1
        )
        with self.assertRaises(ExportError):
            export(self.repo, "HEAD", self.root / "head.xlsx")
        with self.assertRaises(ExportError):
            export(self.repo, "refs/tags/v1", self.root / "tag.xlsx")
        self.assertFalse((self.root / "head.xlsx").exists())


if __name__ == "__main__":
    unittest.main()
