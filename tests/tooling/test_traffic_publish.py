# SPDX-License-Identifier: MIT
"""Offline publication and workflow tests; no GitHub credentials or network."""
import importlib.util
import os
import re
import shlex
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location("traffic_publish", ROOT / "tools/traffic_publish.py")
assert SPEC is not None and SPEC.loader is not None
publisher = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(publisher)
WORKFLOW = ROOT / ".github/workflows/traffic-export.yml"
PARENT = "a" * 40


class Writer:
    def __init__(self, head=PARENT, race=False):
        self.head, self.race, self.calls = head, race, []

    def request(self, method, path, data=None):
        self.calls.append((method, path, data))
        if method == "GET" and "/git/ref/" in path:
            return {"object": {"sha": self.head}}
        if method == "GET" and "/git/commits/" in path:
            return {"tree": {"sha": "b" * 40}}
        if method == "POST":
            return {"sha": "c" * 40}
        if method == "PATCH":
            if self.race:
                raise RuntimeError("non-fast-forward")
            return {}
        raise AssertionError((method, path))


def git(directory, *args):
    subprocess.run(["git", "-C", str(directory), *args], check=True, capture_output=True,
                   env={**os.environ, "GIT_AUTHOR_NAME": "t", "GIT_AUTHOR_EMAIL": "t@example.com",
                        "GIT_COMMITTER_NAME": "t", "GIT_COMMITTER_EMAIL": "t@example.com"})


def history(directory):
    """A traffic-history checkout with one committed history file."""
    git(directory, "init", "-q")
    (directory / "data").mkdir()
    (directory / "data/daily.csv").write_text("date\n", encoding="utf-8")
    (directory / "README.md").write_text("data\n", encoding="utf-8")
    git(directory, "add", ".")
    git(directory, "commit", "-q", "-m", "seed")


def steps(workflow):
    """The step blocks of the export job, by name."""
    export = workflow.split("\n  alert:\n", 1)[0]
    blocks = re.split(r"^      - name: ", export, flags=re.MULTILINE)
    return blocks[0], {block.splitlines()[0]: block for block in blocks[1:]}


class WorkflowTests(unittest.TestCase):
    def test_app_credentials_stay_in_their_steps(self):
        workflow = WORKFLOW.read_text()
        head, named = steps(workflow)
        self.assertNotIn("TRAFFIC_APP_PRIVATE_KEY", head)
        self.assertNotIn("steps.app-token.outputs.token", head)
        key_blocks = [name for name, block in named.items() if "secrets.TRAFFIC_APP_PRIVATE_KEY" in block]
        self.assertEqual(key_blocks, ["Check the traffic App configuration", "Create the traffic App token"])
        check = named["Check the traffic App configuration"]
        self.assertEqual(check.count("secrets.TRAFFIC_APP_PRIVATE_KEY"), 1)
        self.assertIn("${{ secrets.TRAFFIC_APP_PRIVATE_KEY != '' }}", check)
        token = named["Create the traffic App token"]
        self.assertIn("private-key: ${{ secrets.TRAFFIC_APP_PRIVATE_KEY }}", token)
        self.assertIn("permission-administration: read", token)
        self.assertIn("actions/create-github-app-token@64ab729a6294000f5aa4a4bc0b73d1f32bc45510", token)
        fetches = [name for name, block in named.items() if "steps.app-token.outputs.token" in block]
        self.assertEqual(fetches, ["Fetch traffic from the GitHub API"])
        run = named["Fetch traffic from the GitHub API"].split("        run: |\n", 1)[1]
        run = "\n".join(line for line in run.splitlines() if not line.lstrip().startswith("#"))
        commands = run.replace("\\\n", " ").strip().splitlines()
        self.assertEqual(len(commands), 5)
        outputs = []
        for command in commands:
            words = shlex.split(command)
            self.assertEqual(words[0], "curl")
            self.assertFalse(set(words) & {";", "&&", "||", "--location", "-L"})
            self.assertNotIn("$(", command)
            self.assertNotIn("`", command)
            outputs.append(words[words.index("-o") + 1].rsplit("/", 1)[1])
        self.assertEqual(outputs, ["views.json", "clones.json", "referrers.json",
                                   "paths.json", "repository.json"])
        self.assertNotIn("secrets.TRAFFIC_TOKEN", workflow)

    def test_merge_and_publication_never_see_the_app_token(self):
        _, named = steps(WORKFLOW.read_text())
        merge = named["Merge the collection into the history"]
        publish = named["Atomically publish the merged history"]
        for block in (merge, publish):
            self.assertNotIn("app-token", block)
            self.assertNotIn("TRAFFIC_APP", block)
        self.assertNotIn("GITHUB_TOKEN", merge)
        self.assertIn("GITHUB_TOKEN: ${{ github.token }}", publish)
        checkout = named["Check out the traffic history without credentials"]
        self.assertIn("persist-credentials: false", checkout)

    def test_alert_job_has_no_environment_or_app_credentials(self):
        alert = WORKFLOW.read_text().split("\n  alert:\n", 1)[1]
        self.assertNotRegex(alert, r"(?m)^    environment:")
        self.assertNotIn("TRAFFIC_APP", alert)
        self.assertNotIn("app-token", alert)
        self.assertIn("needs: export", alert)
        self.assertIn("issues: write", alert)


class PublishTests(unittest.TestCase):
    def test_changed_files_lists_new_and_modified_history_files_only(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            history(root)
            self.assertEqual(publisher.changed_files(root), {})
            (root / "data/daily.csv").write_text("date\n2026-10-17\n", encoding="utf-8")
            (root / "data/badges").mkdir()
            (root / "data/badges/clones.json").write_text("{}\n", encoding="utf-8")
            (root / "README.md").write_text("changed\n", encoding="utf-8")
            self.assertEqual(publisher.changed_files(root), {
                "data/badges/clones.json": "{}\n", "data/daily.csv": "date\n2026-10-17\n"})
            (root / "data/daily.csv").unlink()
            with self.assertRaisesRegex(ValueError, "deleted"):
                publisher.changed_files(root)

    def test_atomic_commit_on_the_expected_parent(self):
        writer = Writer()
        files = {"data/daily.csv": "date\n", "data/badges/visits.json": "{}\n"}
        self.assertEqual(publisher.publish(writer, "owner/repo", PARENT, files, "m"), "c" * 40)
        writes = [(method, path, data) for method, path, data in writer.calls if method != "GET"]
        self.assertEqual([m for m, _, _ in writes], ["POST", "POST", "PATCH"])
        self.assertEqual(writes[0][2]["base_tree"], "b" * 40)
        self.assertEqual({x["path"] for x in writes[0][2]["tree"]}, set(files))
        self.assertEqual(writes[1][2]["parents"], [PARENT])
        self.assertEqual(writes[-1][2], {"sha": "c" * 40, "force": False})
        self.assertTrue(writes[-1][1].endswith("/refs/heads/traffic-history"))

    def test_nothing_changed_writes_nothing(self):
        writer = Writer()
        self.assertIsNone(publisher.publish(writer, "owner/repo", PARENT, {}, "m"))
        self.assertEqual(writer.calls, [])

    def test_moved_branch_fails_before_any_write(self):
        writer = Writer(head="d" * 40)
        with self.assertRaisesRegex(RuntimeError, "moved"):
            publisher.publish(writer, "owner/repo", PARENT, {"data/daily.csv": "x"}, "m")
        self.assertTrue(all(method == "GET" for method, _, _ in writer.calls))

    def test_concurrent_history_update_is_not_forced_or_retried(self):
        writer = Writer(race=True)
        with self.assertRaisesRegex(RuntimeError, "non-fast-forward"):
            publisher.publish(writer, "owner/repo", PARENT, {"data/daily.csv": "x"}, "m")
        self.assertEqual(sum(m == "PATCH" for m, _, _ in writer.calls), 1)

    def test_invalid_identity_and_parent_fail_before_api_access(self):
        writer = Writer()
        with self.assertRaisesRegex(ValueError, "repository"):
            publisher.publish(writer, "owner/../repo", PARENT, {"data/x": "1"}, "m")
        with self.assertRaisesRegex(ValueError, "parent"):
            publisher.publish(writer, "owner/repo", "abc", {"data/x": "1"}, "m")
        self.assertEqual(writer.calls, [])

    def test_non_main_and_missing_secret_fail_before_api_access(self):
        arguments = ["--history", ".", "--collected-at", "2026-10-18T06:00:00Z"]
        with patch.dict(publisher.os.environ, {"GITHUB_REF": "refs/pull/1/merge"}, clear=True):
            with self.assertRaisesRegex(ValueError, "restricted to main"):
                publisher.main(arguments)
        environment = {"GITHUB_REF": "refs/heads/main", "GITHUB_REPOSITORY": "owner/repo"}
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            history(root)
            (root / "data/daily.csv").write_text("date\n2026-10-17\n", encoding="utf-8")
            with patch.dict(publisher.os.environ, environment, clear=True):
                with self.assertRaisesRegex(ValueError, "credential is missing"):
                    publisher.main(["--history", str(root), "--collected-at", "2026-10-18T06:00:00Z"])
                with self.assertRaisesRegex(ValueError, "collection time"):
                    publisher.main(["--history", str(root), "--collected-at", "yesterday"])


if __name__ == "__main__":
    unittest.main()
