# SPDX-License-Identifier: MIT
"""Exercise the trusted workflow implementation offline; no live mutations."""

import copy
import sys
import unittest
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools"))
import check_repository as checker


class PrMilestoneTests(unittest.TestCase):
    workflow: Any
    reconcile: "staticmethod[..., Any]"  # quoted: staticmethod is not subscriptable at runtime
    summary: "staticmethod[..., Any]"

    @classmethod
    def setUpClass(cls):
        path = Path(__file__).resolve().parents[2] / ".github/workflows/pr-milestones.yml"
        cls.workflow = checker.load_yaml(path.read_text())
        shell = cls.workflow["jobs"]["reconcile"]["steps"][0]["run"]
        assert shell.startswith("python3 - <<'PY'\n") and shell.endswith("\nPY\n")
        namespace: dict[str, Any] = {"__name__": "fixture"}
        exec(compile(shell.split("\n", 1)[1].rsplit("\nPY", 1)[0], str(path), "exec"), namespace)
        cls.reconcile = staticmethod(namespace["reconcile"])
        cls.summary = staticmethod(namespace["summary"])

    def test_full_scan_includes_closed_and_merged_without_shifting_pages(self):
        calls = []
        pending = {2, 101}

        def request(method, path, data=None):
            calls.append((method, path, data))
            if "/pulls?" in path:
                self.assertIn("state=all&sort=created&direction=asc", path)
                numbers = range(1, 101) if path.endswith("page=1") else [101]
                return [{"number": number, "milestone": {} if number in pending else None} for number in numbers]
            number = int(path.rsplit("/", 1)[1])
            if method == "PATCH":
                self.assertEqual(data, {"milestone": None})
                pending.remove(number)
                return {"milestone": None}
            return {"pull_request": {}, "milestone": {"title": "v0.2.0"} if number in pending else None}

        self.assertEqual(self.reconcile(request, "owner/repo", {}), [(2, "v0.2.0"), (101, "v0.2.0")])
        self.assertEqual(pending, set())
        # Each PATCH is followed by an independent read of the new state
        methods = [(method, int(path.rsplit("/", 1)[1])) for method, path, _ in calls if "/issues/" in path]
        self.assertEqual(methods, [("GET", 2), ("PATCH", 2), ("GET", 2), ("GET", 101), ("PATCH", 101), ("GET", 101)])
        self.assertEqual([int(path.rsplit("/", 1)[1]) for method, path, _ in calls if method == "PATCH"], [2, 101])
        calls.clear()
        self.reconcile(request, "owner/repo", {})
        self.assertFalse(any(method == "PATCH" for method, _, _ in calls))

    def test_event_reads_live_state_and_refuses_actual_issues(self):
        calls = []

        def request(method, path, data=None):
            calls.append(method)
            return {"pull_request": {}, "milestone": None}

        self.reconcile(request, "owner/repo", {"number": 7, "pull_request": {"milestone": {}}})
        self.assertEqual(calls, ["GET"])
        with self.assertRaisesRegex(ValueError, "Refusing to change an issue"):
            self.reconcile(lambda *args: {"milestone": {}}, "owner/repo", {"number": 7, "pull_request": {}})

    def test_failures_are_not_reported_as_success(self):
        def failing(*args):
            raise OSError("API unavailable")

        with self.assertRaisesRegex(OSError, "API unavailable"):
            self.reconcile(failing, "owner/repo", {})
        with self.assertRaisesRegex(ValueError, "Malformed PR collection"):
            self.reconcile(lambda *args: {}, "owner/repo", {})
        with self.assertRaisesRegex(ValueError, "pagination exceeded"):
            self.reconcile(lambda *args: [{"milestone": None}] * 100, "owner/repo", {})
        with self.assertRaisesRegex(ValueError, "still has a milestone"):
            self.reconcile(lambda *args: {"pull_request": {}, "milestone": {}}, "owner/repo", {"number": 7, "pull_request": {}})

    def test_a_milestone_still_read_back_fails(self):
        # The PATCH response alone does not prove the new state (#330)
        reads = []

        def request(method, path, data=None):
            if method == "PATCH":
                return {"milestone": None}
            reads.append(path)
            return {"pull_request": {}, "milestone": {"title": "kept"}}

        with self.assertRaisesRegex(ValueError, "PR #7 still has a milestone"):
            self.reconcile(request, "owner/repo", {"number": 7, "pull_request": {}})
        self.assertEqual(len(reads), 2)

    def test_summary_lists_removed_milestones_as_data(self):
        self.assertIn("No PR had a milestone.", self.summary([]))
        text = self.summary([(5, "v0.2.0 | a\\b\nnext")])
        self.assertIn("| PR | Removed milestone |", text)
        self.assertIn("| #5 | v0.2.0 \\| a\\\\b next |", text)
        self.assertEqual(text.count("\n| #"), 1)

    def test_control_cannot_gain_checkout_interpolation_or_extra_permissions(self):
        checker.validate_milestone_workflow(self.workflow)
        for mutation in ["checkout", "interpolation", "permissions", "issue_write", "coverage", "condition", "any_ref"]:
            with self.subTest(mutation=mutation):
                document = copy.deepcopy(self.workflow)
                job = document["jobs"]["reconcile"]
                if mutation == "checkout":
                    job["steps"].append({"uses": "actions/checkout@" + "a" * 40})
                elif mutation == "interpolation":
                    job["steps"][0]["run"] += "${{ github.event.pull_request.title }}"
                elif mutation == "permissions":
                    job["permissions"]["contents"] = "write"
                elif mutation == "issue_write":
                    job["permissions"] = {"issues": "write", "pull-requests": "read"}
                elif mutation == "condition":
                    del job["if"]
                elif mutation == "any_ref":
                    job["if"] = "github.event_name != 'pull_request'"
                else:
                    document["on"]["pull_request_target"]["branches"] = ["main"]
                with self.assertRaises(ValueError):
                    checker.validate_milestone_workflow(document)
