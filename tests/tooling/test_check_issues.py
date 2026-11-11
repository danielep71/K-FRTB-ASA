# SPDX-License-Identifier: MPL-2.0
# This Source Code Form is subject to the terms of the Mozilla Public License,
# v. 2.0. If a copy of the MPL was not distributed with this file, You can
# obtain one at https://mozilla.org/MPL/2.0/.
# See docs/licenses/THIRD_PARTY_NOTICES.txt.
"""Tests for the issue triage rules (tools/check_issues.py)."""

from __future__ import annotations

import contextlib
import io
import json
import shutil
import sys
import tempfile
import unittest
import unittest.mock
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools"))
import check_issues  # noqa: E402
from check_issues import Issue  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
REPOSITORY = "example-org/sample-project"
BLOCKED_LABEL = check_issues.BLOCKED
READY = "## Dependencies and blocking status\n\n**Ready to start**: none.\n"
PLANNED = "- **Target completion date:** 2026-11-13\n\n"


def issue(number: int, *labels: str, state: str = "open", body: str = READY,
          assignees: tuple[str, ...] = ("triage-person",), milestone: str | None = "v0.1.0",
          planned: bool = True, milestone_state: str | None = "open",
          milestone_due: str | None = "2026-11-22") -> Issue:
    return Issue(number, state, labels or ("P3",), assignees, milestone,
                 (PLANNED if planned else "") + body, milestone_state, milestone_due)


def blocked_by(*numbers: int, heading: str = "## Dependencies and blocking status") -> str:
    lines = "\n".join(f"- #{n} — prerequisite" for n in numbers)
    return f"## Objective\n\nx\n\n{heading}\n\n**blocked**:\n\n{lines}\n\n## Tracking\n\n#99\n"


def rules(issues: list[Issue], states: dict[int, str] | None = None,
          pulls: frozenset[int] = frozenset()) -> list[tuple[int, str]]:
    known = states if states is not None else {i.number: i.state for i in issues}
    return [(f.number, f.rule)
            for f in check_issues.audit(issues, known, REPOSITORY, pulls=pulls, assignee="triage-person")]


class MetadataTests(unittest.TestCase):
    def test_compliant_issue_has_no_findings(self) -> None:
        self.assertEqual(rules([issue(1)]), [])

    def test_assignee_priority_and_milestone_are_required(self) -> None:
        self.assertEqual(rules([issue(1, assignees=())]), [(1, "assignee")])
        self.assertEqual(rules([issue(1, "documentation")]), [(1, "priority")])
        self.assertEqual(rules([issue(1, "P2", "P3")]), [(1, "priority")])
        self.assertEqual(rules([issue(1, milestone=None)]), [(1, "milestone")])

    def test_p1_is_reserved_for_bugs(self) -> None:
        self.assertEqual(rules([issue(1, "P1", "documentation")]), [(1, "p1-reserved")])
        self.assertEqual(rules([issue(1, "P1", "bug")]), [])

    def test_closed_issues_are_not_audited(self) -> None:
        self.assertEqual(rules([issue(1, "documentation", state="closed", assignees=())]), [])

    def test_unknown_labels_are_reported_against_the_catalogue(self) -> None:
        findings = check_issues.audit([issue(1, "P3", "ready")], {1: "open"}, REPOSITORY,
                                      catalogue={"P3"}, assignee="triage-person")
        self.assertEqual([(f.number, f.rule) for f in findings], [(1, "catalogue")])


class PlanningTests(unittest.TestCase):
    def test_the_milestone_must_be_open(self) -> None:
        self.assertEqual(rules([issue(1, milestone_state="closed")]), [(1, "milestone")])

    def test_target_date_is_required_and_real(self) -> None:
        self.assertEqual(rules([issue(1, planned=False)]), [(1, "target-date")])
        for value in ("2026-02-30", "2026-13-01", "20261113", "soon", "2026-11-1"):
            with self.subTest(value=value):
                body = f"- **Target completion date:** {value}\n\n" + READY
                self.assertEqual(rules([issue(1, planned=False, body=body)]), [(1, "target-date")])

    def test_target_date_must_not_pass_the_due_date(self) -> None:
        self.assertEqual(rules([issue(1, milestone_due="2026-11-12")]), [(1, "target-date")])
        self.assertEqual(rules([issue(1, milestone_due="2026-11-13")]), [])
        self.assertEqual(rules([issue(1, milestone_due=None)]), [])

    def test_form_field_and_planning_line_precedence(self) -> None:
        form = "### 📅 Target end date\n\n2026-11-20\n\n" + READY
        self.assertEqual(rules([issue(1, planned=False, body=form)]), [])
        self.assertEqual(rules([issue(1, planned=False, body="### Target end date\n\n_No response_\n\n" + READY)]),
                         [(1, "target-date")])
        self.assertEqual(rules([issue(1, planned=False, body="### Target end date\n\n" + READY)]),
                         [(1, "target-date")])
        # the planning line wins over the form field
        both = "- **Target completion date:** 2026-11-30\n\n" + form
        self.assertEqual(rules([issue(1, planned=False, body=both)]), [(1, "target-date")])
        self.assertEqual(check_issues.target_date(both), "2026-11-30")

    def test_milestone_fields_come_from_the_api(self) -> None:
        item = {"number": 1, "state": "open", "labels": [], "assignees": [], "body": "",
                "milestone": {"title": "v0.2.0", "state": "closed", "due_on": "2026-11-22T08:00:00Z"}}
        parsed = check_issues.issue_from_api(item)
        self.assertEqual((parsed.milestone_state, parsed.milestone_due), ("closed", "2026-11-22"))
        item["milestone"] = {"title": "v0.3.0", "state": "open", "due_on": None}
        parsed = check_issues.issue_from_api(item)
        self.assertEqual((parsed.milestone_state, parsed.milestone_due), ("open", None))


class DependencyTests(unittest.TestCase):
    def test_blocked_by_an_open_prerequisite(self) -> None:
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(1)), issue(1)]), [])

    def test_blocked_with_only_closed_prerequisites_must_be_unblocked(self) -> None:
        closed = issue(1, state="closed")
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(1)), closed]),
                         [(2, "blocked")])

    def test_open_prerequisite_requires_blocked(self) -> None:
        self.assertEqual(rules([issue(2, "P3", body=blocked_by(1)), issue(1)]), [(2, "blocked")])

    def test_external_condition_counts_as_a_blocker(self) -> None:
        body = "### 🔗 Dependencies and blocking status\n\nExternal condition: Excel host access\n"
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=body)]), [])
        self.assertEqual(rules([issue(2, "P3", body=body)]), [(2, "blocked")])

    def test_references_outside_the_section_are_ignored(self) -> None:
        # blocked_by() also mentions #99 under Tracking; it must not count.
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(1)), issue(1)],
                               {1: "open", 2: "open"}), [])

    def test_issue_urls_and_form_headings_are_recognised(self) -> None:
        body = ("### 🔗 Dependencies and blocking status\n\n"
                f"- https://github.com/{REPOSITORY}/issues/1\n")
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=body), issue(1)]), [])

    def test_missing_section_and_unknown_reference(self) -> None:
        self.assertEqual(rules([issue(1, body="## Objective\n\nx\n")]), [(1, "dependencies")])
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(7))]),
                         [(2, "dependencies"), (2, "blocked")])

    def test_self_reference_is_ignored(self) -> None:
        self.assertEqual(rules([issue(2, "P3", body=blocked_by(2))]), [])

    def test_named_pull_request_is_a_reference_not_a_prerequisite(self) -> None:
        body = "## Dependencies and blocking status\n\nNone. #1 was delivered in #22.\n"
        closed = issue(1, state="closed")
        self.assertEqual(rules([issue(2, "P3", body=body), closed], pulls=frozenset({22})), [])
        # A pull request never justifies blocked, and unknown numbers are still reported.
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(22, 7))],
                               pulls=frozenset({22})),
                         [(2, "dependencies"), (2, "blocked")])



class ConfigTests(unittest.TestCase):
    def test_repository_configuration_passes(self) -> None:
        self.assertEqual(check_issues.config_findings(ROOT), [])

    def test_forms_must_require_a_target_end_date(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for part in (".github", "docs"):
                shutil.copytree(ROOT / part, root / part)
            form = root / ".github/ISSUE_TEMPLATE/02-feature.yml"
            text = form.read_text(encoding="utf-8")
            start = text.index("    id: target_end_date")
            optional = text[:start] + text[start:].replace("required: true", "required: false", 1)
            form.write_text(optional, encoding="utf-8")
            self.assertIn("02-feature.yml must require target_end_date", check_issues.config_findings(root))
            form.write_text(text.replace("    id: target_end_date", "    id: planned_date"), encoding="utf-8")
            self.assertIn("02-feature.yml has no target_end_date field", check_issues.config_findings(root))

    def test_unsupported_overlays_are_rejected_instead_of_ignored(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shutil.copytree(ROOT / ".github/scripts", root / ".github/scripts")
            manifest = json.loads((ROOT / ".github/labels.json").read_text())
            manifest["overlays"] = {"selected": [{"name": "extra"}]}
            (root / ".github/labels.json").write_text(json.dumps(manifest))
            with self.assertRaisesRegex(ValueError, "root must contain exactly"):
                check_issues.catalogue_labels(root)

    def test_assignee_comes_from_configuration_not_repository_identity(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            forms = root / ".github/ISSUE_TEMPLATE"
            forms.mkdir(parents=True)
            (forms / "01-bug.yml").write_text('assignees: [triage-person]\n')
            self.assertEqual(check_issues.configured_assignee(root), "triage-person")
            (forms / "01-bug.yml").write_text('assignees:\n- triage-person\n')
            self.assertEqual(check_issues.configured_assignee(root), "triage-person")
            (forms / "01-bug.yml").write_text('assignees: []\n')
            with self.assertRaises(ValueError):
                check_issues.configured_assignee(root)

    @staticmethod
    def copy_configuration(root: Path) -> None:
        """The files config_findings reads, copied from the repository."""
        for folder in (".github/scripts", ".github/ISSUE_TEMPLATE"):
            shutil.copytree(ROOT / folder, root / folder)
        for name in (".github/labels.json", ".github/dependabot.yml", "docs/LABELS.md"):
            (root / name).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy(ROOT / name, root / name)

    def test_label_guide_explains_every_catalogue_label_once(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self.copy_configuration(root)
            guide = root / "docs/LABELS.md"
            text = guide.read_text(encoding="utf-8")
            row = next(line for line in text.splitlines() if line.startswith("| `wontfix` |"))
            guide.write_text(text.replace(row + "\n", ""), encoding="utf-8")
            self.assertIn("docs/LABELS.md must explain label 'wontfix' exactly once in the label tables "
                          "of Metadata rules, found 0", check_issues.config_findings(root))
            guide.write_text(text.replace(row, row + "\n" + row), encoding="utf-8")
            self.assertIn("docs/LABELS.md must explain label 'wontfix' exactly once in the label tables "
                          "of Metadata rules, found 2", check_issues.config_findings(root))
            guide.write_text(text.replace(row, row + "\n| `stale` | Not a label. |"), encoding="utf-8")
            self.assertIn("docs/LABELS.md explains label 'stale', which is not in the catalogue",
                          check_issues.config_findings(root))

    def test_documented_labels_reads_only_the_label_tables_of_the_metadata_rules(self) -> None:
        text = ("## Branches\n\n| Label | Use |\n| --- | --- |\n| `main` | x |\n\n"
                "## Metadata rules\n\n| Priority | Use |\n| --- | --- |\n| `P1` | x |\n\nText `bug`.\n\n"
                "| Other | Use |\n| --- | --- |\n| `ci` | x |\n\n| Label | Use |\n| --- | --- |\n| `bug` | x |\n\n"
                "## Releases\n\n| Label | Use |\n| --- | --- |\n| `release` | x |\n")
        self.assertEqual(check_issues.documented_labels(text), ["P1", "bug"])

    def test_dependabot_labels_and_assignees_are_checked(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self.copy_configuration(root)
            self.assertEqual(check_issues.config_findings(root), [])
            config = root / ".github/dependabot.yml"
            text = config.read_text(encoding="utf-8")
            config.write_text(text.replace("labels: [repository, ci, P3]", "labels: [repository, no-such-label, P3]", 1),
                              encoding="utf-8")
            self.assertIn(".github/dependabot.yml sets unknown label 'no-such-label'",
                          check_issues.config_findings(root))
            config.write_text(text.replace("assignees: [danielep71]", "assignees: [someone-else]", 1), encoding="utf-8")
            self.assertIn(".github/dependabot.yml assigns 'someone-else', not 'danielep71'",
                          check_issues.config_findings(root))
            config.write_text(text.replace("    labels: [repository, ci, P3]\n",
                                           "    labels:\n      - no-such-label # custom labels\n", 1), encoding="utf-8")
            self.assertIn(".github/dependabot.yml sets unknown label 'no-such-label'",
                          check_issues.config_findings(root))

    def test_block_lists(self) -> None:
        text = ("updates:\n  - x: 1\n    labels:\n      - ci   # a comment\n\n      - \"P3\"\n"
                "    assignees: [triage-person]\n    other: y\n  - x: 2\n    labels:\n    - tests\n")
        self.assertEqual(check_issues.block_lists(text, "labels"), [["ci", "P3"], ["tests"]])
        self.assertEqual(check_issues.block_lists(text, "assignees"), [["triage-person"]])
        self.assertEqual(check_issues.block_lists(text, "missing"), [])
        commented = "labels: [no-such-label, P3]  # custom labels\nassignees: [someone-else] # owner\n"
        self.assertEqual(check_issues.block_lists(commented, "labels"), [["no-such-label", "P3"]])
        self.assertEqual(check_issues.block_lists(commented, "assignees"), [["someone-else"]])

    def test_form_list_reads_inline_and_block_lists(self) -> None:
        block = "name: x\nlabels:\n- bug\n- P2\nassignees:\n- triage-person\nbody:\n- type: textarea\n"
        self.assertEqual(check_issues.form_list(block, "labels"), ["bug", "P2"])
        self.assertEqual(check_issues.form_list(block, "assignees"), ["triage-person"])
        self.assertEqual(check_issues.form_list('labels: ["bug", "P2"]\n', "labels"), ["bug", "P2"])
        self.assertEqual(check_issues.form_list("assignees: [triage-person] # owner\n", "assignees"),
                         ["triage-person"])
        self.assertEqual(check_issues.form_list("assignees: # owner\n  - triage-person\n", "assignees"),
                         ["triage-person"])
        nested = "body:\n  - labels:\n      - nested\nlabels:\n- bug\n"
        self.assertEqual(check_issues.form_list(nested, "labels"), ["bug"])
        self.assertEqual(check_issues.form_list(block, "title"), [])

    def test_flow_list(self) -> None:
        text = 'name: x\nlabels: ["bug", "P2"]\nassignees: [triage-person]\n'
        self.assertEqual(check_issues.flow_list(text, "labels"), ["bug", "P2"])
        self.assertEqual(check_issues.flow_list(text, "assignees"), ["triage-person"])
        self.assertEqual(check_issues.flow_list(text, "title"), [])


class MainTests(unittest.TestCase):
    def test_input_mode_reports_and_exits_nonzero_on_findings(self) -> None:
        items = [
            {"number": 1, "state": "open", "labels": [{"name": "P3"}],
             "assignees": [{"login": "triage-person"}], "milestone": None,
             "body": PLANNED + READY + "Delivered in #2, a pull request.\n"},
            {"number": 2, "state": "open", "labels": [], "assignees": [],
             "milestone": None, "body": READY, "pull_request": {}},
        ]
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "issues.json"
            source.write_text(json.dumps(items), encoding="utf-8")
            output = Path(directory) / "report.json"
            with contextlib.redirect_stdout(io.StringIO()) as stdout, \
                    unittest.mock.patch.object(check_issues, "configured_assignee",
                                               return_value="triage-person"):
                code = check_issues.main(["--root", str(ROOT), "--input", str(source),
                                          "--output", str(output), "--repo-name", REPOSITORY])
            self.assertEqual(code, 1)
            report = json.loads(output.read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "fail")
            self.assertEqual([f["rule"] for f in report["findings"]], ["milestone"])
            self.assertIn("open issues audited: 1", stdout.getvalue())

    def test_offline_mode_requires_explicit_repository(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as error:
            check_issues.parse_args(["--input", "issues.json"])
        self.assertEqual(error.exception.code, 2)

    def test_live_mode_requires_a_token(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), \
                unittest.mock.patch.dict("os.environ", {"GITHUB_TOKEN": ""}):
            self.assertEqual(check_issues.main(["--repository", REPOSITORY]), 2)


if __name__ == "__main__":
    unittest.main()
