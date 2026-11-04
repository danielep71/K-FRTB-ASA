# SPDX-License-Identifier: MIT
"""Synthetic failures for the universal maintenance gate, without network."""

import copy
import json
from pathlib import Path
import sys
import subprocess
import tempfile
import unittest
import unittest.mock

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools"))
import check_repository as checker


CONTROL_FILES = [".github/workflows/integrity.yml", ".github/workflows/pr-milestones.yml", ".github/dependabot.yml",
                 ".github/requirements.txt", "tests/requirements.txt"]


class RepositoryIntegrityTests(unittest.TestCase):
    def check_fixture(self, files):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name, content in files.items():
                path = root / name
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(content)
            return checker.validate_files(root, list(files))

    def test_valid_links_and_markdown_hard_breaks(self):
        self.assertEqual(self.check_fixture({"guide.md": "# Guide\n\n[Child](docs/child.md#scope)  \n[Site](https://example.com)\n", "docs/child.md": "# Scope\n"}), [])

    def test_invalid_and_duplicate_configurations(self):
        for name, text in [("a.json", '{"a":1,"a":2}\n'), ("a.json", "\0\n"), ("a.yml", "a: 1\na: 2\n"), ("a.yml", "a: [\n")]:
            with self.subTest(name=name, text=text):
                self.assertTrue(self.check_fixture({name: text}))
        self.assertIn("on", checker.load_yaml("on:\n  push: null\n"))
        for constant in ["NaN", "Infinity", "-Infinity"]:
            with self.subTest(constant=constant):
                errors = self.check_fixture({"a.json": '{"value":' + constant + '}\n'})
                self.assertTrue(any("non-standard JSON constant" in error for error in errors))

    def test_binary_formats_follow_git_attributes(self):
        source = Path(__file__).resolve().parents[2]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            (root / ".gitattributes").write_text((source / ".gitattributes").read_text() + "\n*.fixture opaque\n")
            names = ["preview.png", "screenshot.jpeg", "bundle.zip", "K-FRTB-SA.xlsm", "template.xlsx", "custom.fixture"]
            for name in names + ["bad.txt"]:
                (root / name).write_bytes(b"\x00\xff\x10")
            binary = checker.binary_paths(root, names + ["bad.txt"])
            self.assertEqual(binary, set(names))
            self.assertEqual(checker.validate_files(root, names, binary), [])
            self.assertTrue(checker.validate_files(root, ["bad.txt"], binary))

    def test_broken_links_and_tracked_whitespace(self):
        for text in ["[Missing](missing.md)\n", "[Outside](../outside.md)\n", "bad \n", "no newline", "extra\n\n"]:
            with self.subTest(text=text):
                self.assertTrue(self.check_fixture({"guide.md": text}))
        self.assertEqual(self.check_fixture({"guide.md": "```md\n[Example](absent.md)\n```\n"}), [])

    def test_markdown_fragments_must_name_an_anchor(self):
        child = "# 🧭 Scope (MAR21.5)\n\n## Notes\n\n## Notes\n\n<a id=\"explicit\"></a>\n"
        valid = ("# Guide\n\n[Top](#guide) [Child](docs/child.md#-scope-mar215) [Repeat](docs/child.md#notes-1)\n"
                 "[Explicit](docs/child.md#explicit) [Case](docs/child.md#Notes)\n")
        self.assertEqual(self.check_fixture({"guide.md": valid, "docs/child.md": child}), [])
        for text, fragment in [("[Gone](#nowhere)\n", "nowhere"), ("[Gone](docs/child.md#absent)\n", "absent"),
                               ("[Third](docs/child.md#notes-2)\n", "notes-2")]:
            with self.subTest(fragment=fragment):
                errors = self.check_fixture({"guide.md": "# Guide\n\n" + text, "docs/child.md": child})
                self.assertTrue(any("link fragment not found" in error and fragment in error for error in errors), errors)
        errors = self.check_fixture({"guide.md": '<a id="twice"></a>\n<a name="twice"></a>\n'})
        self.assertTrue(any("duplicate explicit anchor: twice" in error for error in errors), errors)
        # a heading inside a fenced block defines no anchor, and a link there is not checked
        fenced = "```md\n# Fake\n[Fake](#fake)\n```\n"
        self.assertEqual(self.check_fixture({"guide.md": fenced}), [])
        self.assertTrue(self.check_fixture({"guide.md": fenced + "[Fake](#fake)\n"}))
        # an anchor shown as inline code is an example, not an anchor
        example = "Write `<a id=\"sample\"></a>`, as in `<a id=\"sample\"></a>`.\n"
        self.assertEqual(self.check_fixture({"guide.md": example}), [])
        self.assertTrue(self.check_fixture({"guide.md": example + "[Sample](#sample)\n"}))

    def test_setext_headings_define_anchors(self):
        anchors, _ = checker.markdown_anchors("Scope\n=====\n\nDetails here\n---\n\n- item\n---\n\n| a |\n| --- |\n")
        self.assertEqual(anchors, {"scope", "details-here"})

    def test_repeated_heading_slugs_skip_taken_suffixes(self):
        # GitHub: Scope, Scope, Scope-1 give scope, scope-1, scope-1-1
        anchors, _ = checker.markdown_anchors("# Scope\n\n# Scope\n\n# Scope-1\n\n# Scope\n")
        self.assertEqual(anchors, {"scope", "scope-1", "scope-1-1", "scope-2"})

    def test_heading_slugs_follow_github(self):
        for heading, slug in [("🧭 Stages", "-stages"), ("Curvature input contract (MAR21.5, MAR21.98)", "curvature-input-contract-mar215-mar2198"),
                              ("Link [text](https://x.y) and `code`", "link-text-and-code"), ("snake_case — dash", "snake_case--dash"),
                              ("__Scope__ and *notes*", "scope-and-notes"), ("Terms &amp; conditions", "terms--conditions"),
                              ("**Bold** _it_ snake_case_name", "bold-it-snake_case_name"),
                              ("The `<input>` field", "the-input-field"), ("`__init__` and `&amp;`", "__init__-and-amp"),
                              ("__Bold `x` text__", "bold-x-text"), ("``a `b` c``", "a-b-c")]:
            with self.subTest(heading=heading):
                self.assertEqual(checker.heading_slug(heading), slug)

    @staticmethod
    def workflow_fixture():
        return {
            "name": "Fixture",
            "on": {"push": None},
            "permissions": {"contents": "read"},
            "jobs": {"test": {
                "timeout-minutes": 5,
                "steps": [{"uses": "actions/checkout@" + "a" * 40, "with": {"persist-credentials": False}}],
            }},
        }

    def test_unpinned_actions_are_rejected(self):
        good = self.workflow_fixture()
        checker.validate_workflow(good)
        workflow = copy.deepcopy(good)
        workflow["jobs"]["test"]["steps"][0]["uses"] = "actions/checkout@main"
        with self.assertRaisesRegex(ValueError, "full commit SHA"):
            checker.validate_workflow(workflow)

    def test_workflow_controls_are_enforced(self):
        def mutated(change):
            document = self.workflow_fixture()
            change(document, document["jobs"]["test"])
            return document

        cases = {
            "persist-credentials": lambda d, j: j["steps"][0].pop("with"),
            "persist-credentials: false": lambda d, j: j["steps"][0]["with"].update({"persist-credentials": True}),
            "top-level permissions": lambda d, j: d.pop("permissions"),
            "timeout-minutes": lambda d, j: j.pop("timeout-minutes"),
            "write permission: contents": lambda d, j: j.update({"permissions": {"contents": "write"}}),
            "write permission: \\*": lambda d, j: d.update({"permissions": "write-all"}),
            "write permission: issues": lambda d, j: d.update({"permissions": {"issues": "write"}}),
        }
        for message, change in cases.items():
            with self.subTest(rule=message):
                with self.assertRaisesRegex(ValueError, message):
                    checker.validate_workflow(mutated(change))

    def test_write_permissions_are_limited_to_the_listed_jobs(self):
        document = self.workflow_fixture()
        document["jobs"] = {"export": document["jobs"]["test"]}
        document["jobs"]["export"]["permissions"] = {"contents": "write"}
        checker.validate_workflow(document, "traffic-export.yml")
        with self.assertRaisesRegex(ValueError, "write permission: contents"):
            checker.validate_workflow(document, "integrity.yml")
        document["jobs"]["export"]["permissions"] = {"contents": "write", "issues": "write"}
        with self.assertRaisesRegex(ValueError, "write permission: issues"):
            checker.validate_workflow(document, "traffic-export.yml")
        document["jobs"] = {"alert": document["jobs"]["export"]}
        document["jobs"]["alert"]["permissions"] = {"contents": "read", "issues": "write"}
        checker.validate_workflow(document, "traffic-export.yml")
        document["jobs"]["alert"]["permissions"] = {"contents": "write", "issues": "write"}
        with self.assertRaisesRegex(ValueError, "write permission: contents"):
            checker.validate_workflow(document, "traffic-export.yml")

    def test_label_reconcile_must_be_limited_to_the_default_branch(self):
        root = Path(__file__).resolve().parents[2]
        document = checker.load_yaml((root / ".github/workflows/labels-sync.yml").read_text())
        checker.validate_workflow(document, "labels-sync.yml")
        for condition in ["github.event_name != 'pull_request'", None]:
            with self.subTest(condition=condition):
                broken = copy.deepcopy(document)
                if condition is None:
                    del broken["jobs"]["reconcile"]["if"]
                else:
                    broken["jobs"]["reconcile"]["if"] = condition
                with self.assertRaisesRegex(ValueError, "default branch"):
                    checker.validate_workflow(broken, "labels-sync.yml")

    def test_path_filters_and_wrong_dependabot_target_fail(self):
        source = Path(__file__).resolve().parents[2]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name in CONTROL_FILES:
                path = root / name
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text((source / name).read_text())
            checker.validate_controls(root)
            workflow_path = root / ".github/workflows/integrity.yml"
            workflow = checker.load_yaml(workflow_path.read_text())
            workflow["on"]["pull_request"]["paths"] = ["tools/**"]
            workflow_path.write_text(yaml.safe_dump(workflow))
            with self.assertRaisesRegex(ValueError, "without path filters"):
                checker.validate_controls(root)
            workflow_path.write_text((source / ".github/workflows/integrity.yml").read_text())
            dependabot_path = root / ".github/dependabot.yml"
            dependabot_path.write_text(dependabot_path.read_text().replace("target-branch: release/v0.2.0", "target-branch: main", 1))
            with self.assertRaisesRegex(ValueError, "Dependabot must cover"):
                checker.validate_controls(root)

    def test_reproducibility_controls(self):
        source = Path(__file__).resolve().parents[2]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name in CONTROL_FILES:
                (root / name).parent.mkdir(parents=True, exist_ok=True)
                (root / name).write_text((source / name).read_text())
            checker.validate_controls(root)
            integrity = root / ".github/workflows/integrity.yml"
            original = integrity.read_text()

            def broken(change, message):
                workflow = checker.load_yaml(original)
                change(workflow, workflow["jobs"]["integrity"]["steps"])
                integrity.write_text(yaml.safe_dump(workflow))
                with self.subTest(message=message), self.assertRaisesRegex(ValueError, message):
                    checker.validate_controls(root)

            def uses(steps, action):
                return next(step for step in steps if step.get("uses", "").startswith(action + "@"))

            broken(lambda w, s: w["concurrency"].update({"cancel-in-progress": True}), "never push runs")
            broken(lambda w, s: uses(s, "actions/setup-python")["with"].update({"python-version": "3.13"}),
                   "exact python-version")
            broken(lambda w, s: uses(s, "actions/setup-node")["with"].update({"node-version": "lts/*"}),
                   "exact node-version")
            broken(lambda w, s: [step.update({"run": step["run"].replace("--require-hashes ", "")})
                                 for step in s if "pip install" in step.get("run", "")], "by hash")
            broken(lambda w, s: uses(s, "actions/upload-artifact").pop("if"), "always upload")
            broken(lambda w, s: uses(s, "actions/upload-artifact")["with"].update({"if-no-files-found": "warn"}),
                   "always upload")
            integrity.write_text(original)
            milestones = root / ".github/workflows/pr-milestones.yml"
            milestones.write_text(milestones.read_text().replace(
                "github.event_name == 'pull_request_target' && format('pr-milestones-pr-{0}', "
                "github.event.pull_request.number) || ", "").replace("${{ 'pr-milestones-scan' }}", "pr-milestones"))
            with self.assertRaisesRegex(ValueError, "one concurrency group per PR"):
                checker.validate_controls(root)
            milestones.write_text((source / ".github/workflows/pr-milestones.yml").read_text())
            requirements = root / "tests/requirements.txt"
            requirements.write_text("numpy==2.5.4\n")
            with self.assertRaisesRegex(ValueError, "exact pin and a hash"):
                checker.validate_controls(root)

    def test_every_workflow_installs_by_hash(self):
        workflow = self.workflow_fixture()
        steps = workflow["jobs"]["test"]["steps"]
        steps.append({"run": "python -m pip install --require-hashes --only-binary=:all: -r requirements.txt"})
        checker.validate_workflow(workflow)
        steps[-1]["run"] = "python -m pip install -r requirements.txt"
        with self.assertRaisesRegex(ValueError, "--require-hashes"):
            checker.validate_workflow(workflow)

    def test_report_is_bound_to_commit_run_and_attempt(self):
        checks = [{"name": "one", "status": "PASS", "seconds": 0.1}, {"name": "two", "status": "FAIL", "detail": "exit code 1"}]
        environment = {"GITHUB_SHA": "a" * 40, "GITHUB_RUN_ID": "7", "GITHUB_RUN_ATTEMPT": "2", "REPORT_HEAD_SHA": "b" * 40}
        with tempfile.TemporaryDirectory() as directory, unittest.mock.patch.dict(checker.os.environ, environment):
            path = Path(directory) / "reports" / "integrity.json"
            checker.write_report(path, checks, 1, {"python": "3.13.18"})
            report = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual((report["commit"], report["run_id"], report["run_attempt"], report["result"]),
                         ("a" * 40, "7", "2", "FAIL"))
        self.assertEqual(report["head_commit"], "b" * 40)
        self.assertEqual(report["checks"], checks)
        self.assertEqual(report["runtimes"], {"python": "3.13.18"})

    def test_dependabot_retirement_and_coverage_regressions(self):
        source = Path(__file__).resolve().parents[2]
        baseline = checker.load_yaml((source / ".github/dependabot.yml").read_text())
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workflow_path = root / ".github/workflows/integrity.yml"
            workflow_path.parent.mkdir(parents=True)
            workflow_path.write_text((source / ".github/workflows/integrity.yml").read_text())
            milestone_path = root / ".github/workflows/pr-milestones.yml"
            milestone_path.write_text((source / ".github/workflows/pr-milestones.yml").read_text())
            for name in (".github/requirements.txt", "tests/requirements.txt"):
                (root / name).parent.mkdir(parents=True, exist_ok=True)
                (root / name).write_text((source / name).read_text())
            dependabot_path = root / ".github/dependabot.yml"

            def validate(document):
                dependabot_path.write_text(yaml.safe_dump(document))
                checker.validate_controls(root)

            validate(baseline)
            for index, entry in enumerate(baseline["updates"]):
                with self.subTest(ecosystem=entry["package-ecosystem"], directory=entry["directory"]):
                    missing = copy.deepcopy(baseline)
                    del missing["updates"][index]
                    duplicate = copy.deepcopy(baseline)
                    duplicate["updates"].append(copy.deepcopy(entry))
                    retired = copy.deepcopy(baseline)
                    retired["updates"][index]["target-branch"] = "release/v0.1.0"
                    for broken in [missing, duplicate, retired]:
                        with self.assertRaisesRegex(ValueError, "Dependabot must cover"):
                            validate(broken)
                    wrong_milestone = copy.deepcopy(baseline)
                    wrong_milestone["updates"][index]["milestone"] = 1
                    with self.assertRaisesRegex(ValueError, "no milestone setting"):
                        validate(wrong_milestone)


if __name__ == "__main__":
    unittest.main()
