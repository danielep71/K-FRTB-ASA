"""Validate an Excel evidence record of K-FRTB-ASA, offline (#82, #28).

A record, host.json, describes one manual run of the workbook in desktop
Excel for Windows, bound to one exact commit. It sits in a bundle folder
with the logs it cites. The format, the procedure and the rules are in
docs/EXCEL_EVIDENCE.md; what a record must contain is in
.github/excel-evidence-policy.json at the candidate commit.

The validator never runs Excel. It checks that a record is complete and
consistent, that its sources are the ones stored in Git at the candidate
and that every log it cites says what the record says. It cannot prove what
a person did. A synthetic bundle is never Excel evidence.

    python3 tools/check_excel_evidence.py --candidate-sha <commit> --inventory
    python3 tools/check_excel_evidence.py --check-policy
    python3 tools/check_excel_evidence.py --candidate-sha <commit> --evidence <bundle>/host.json
        [--output report.json] [--summary report.md]

Exit codes: 0 a consistent passing record (or a successful --inventory or
--check-policy), 1 a non-passing or invalid record, 2 a command error.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[1]
POLICY = ".github/excel-evidence-policy.json"
SOURCE_FOLDERS = ("src/modules", "tests/modules")
TEMPLATE = "src/workbook/FRTB_Template.xlsx"
SETUP_MODULE = "src/modules/M_Setup.bas"
SCHEMA_VERSION = 1
STATUSES = ("PASS", "FAIL", "NOT_RUN", "TIMEOUT")
SHA = re.compile(r"^[0-9a-f]{40}$")
DIGEST = re.compile(r"^[0-9a-f]{64}$")
CANDIDATE_LINE = re.compile(r"^CANDIDATE=(.*)$")
SAVED_LINE = re.compile(r"^Saved \S")
REGRESSION_LINE = re.compile(
    r"^Regression test: Result: (?:(PASS) \((\d+) items\)|(FAIL) \((\d+) of (\d+) items\))\s*$")
COMPARE_CANDIDATE = re.compile(r"^Candidate: (.*)$")
COMPARE_MODE = re.compile(r"^Mode:\s+(\w+) \(")
COMPARE_TEMPLATE = re.compile(r"^Template:\s+(\S+) sha256 ([0-9a-f]{64})$")
COMPARE_WORKBOOK = re.compile(r"^Workbook:\s+.+ sha256 ([0-9a-f]{64})$")
COMPARE_RESULT = re.compile(r"^RESULT=(PASS|FAIL); modules=\d+; differences=\d+; binding=(\w+)$")
GENERATED_ROW = re.compile(r"^\s*PutRow ws, r \+ \d+, Array\(")
ENVIRONMENT_TEXT = ("excel_version", "excel_build", "os", "os_architecture", "decimal_separator",
                    "list_separator", "macro_policy", "vba_project_access")
FAILED_OUTCOME = {"build": "BUILD_FAILED", "manual": "COMPILE_FAILED",
                  "regression": "TEST_FAILED", "compare": "ROUND_TRIP_FAILED"}


class CommandError(Exception):
    """A problem with the command, not with the record: exit code 2."""


def git(root: Path, *args: str) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(["git", "-C", str(root), *args], capture_output=True, check=False)


def resolve_commit(root: Path, candidate: str) -> str:
    completed = git(root, "rev-parse", "--verify", f"{candidate}^{{commit}}")
    if completed.returncode != 0:
        raise CommandError(f"unknown commit: {candidate}")
    return completed.stdout.decode().strip()


def committed(root: Path, sha: str, path: str) -> bytes:
    completed = git(root, "show", f"{sha}:{path}")
    if completed.returncode != 0:
        raise CommandError(f"{path} not found at {sha}")
    return completed.stdout


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def inventory(root: Path, sha: str) -> list[dict[str, str]]:
    """Every VBA module and the template at ``sha``, with the SHA-256 of the
    bytes stored in Git (LF line endings, not a Windows checkout's CRLF)."""
    listing = git(root, "ls-tree", "-r", "--name-only", sha, "--", *SOURCE_FOLDERS)
    if listing.returncode != 0:
        raise CommandError(f"cannot list the sources at {sha}")
    paths = sorted(path for path in listing.stdout.decode("utf-8").splitlines() if path.endswith(".bas"))
    if not paths:
        raise CommandError(f"no modules found at {sha}")
    paths.append(TEMPLATE)
    return [{"path": path, "sha256": sha256(committed(root, sha, path))} for path in paths]


def load_policy(text: str, label: str) -> dict[str, Any]:
    try:
        policy: dict[str, Any] = json.loads(text)
        names = [stage["name"] for stage in policy["stages"]]
        if policy["schema_version"] != 1 or len(set(names)) != len(names):
            raise ValueError("unsupported schema or a stage listed twice")
        for stage in policy["stages"]:
            if stage["kind"] not in FAILED_OUTCOME or not set(stage["depends_on"]) <= set(names):
                raise ValueError(f"stage {stage['name']}: unknown kind or dependency")
            if stage["kind"] == "compare" and stage.get("mode") not in ("production", "development"):
                raise ValueError(f"stage {stage['name']}: compare stages need a mode")
        if not isinstance(policy["regression_items"], int) or policy["regression_items"] <= 0:
            raise ValueError("regression_items is not a positive integer")
        former = policy.get("former_repositories", [])
        if not isinstance(former, list) or not all(isinstance(name, str) for name in former):
            raise ValueError("former_repositories is not a list of names")
        return policy
    except (ValueError, KeyError, TypeError) as error:
        raise CommandError(f"{label}: invalid policy: {error}") from error


def generated_items(setup_source: str) -> int:
    """The expected values M_Setup embeds between its generated markers."""
    block = setup_source.split("' BEGIN GENERATED EXPECTED VALUES", 1)[-1].split("' END GENERATED EXPECTED VALUES", 1)[0]
    return sum(bool(GENERATED_ROW.match(line)) for line in block.splitlines())


def check_policy(root: Path) -> list[str]:
    """The disagreements between the working-tree policy and the source."""
    policy = load_policy((root / POLICY).read_text(encoding="utf-8"), POLICY)
    items = generated_items((root / SETUP_MODULE).read_text(encoding="cp1252"))
    findings = []
    if policy["regression_items"] != items:
        findings.append(f"{POLICY}: regression_items is {policy['regression_items']}, "
                        f"{SETUP_MODULE} embeds {items} expected values")
    if len(set(policy["required_references"])) != len(policy["required_references"]):
        findings.append(f"{POLICY}: a reference is listed twice")
    return findings


def decode_log(data: bytes) -> str:
    """A log as Windows tools write it: UTF-16 with a byte-order mark (as
    PowerShell 5.1 redirection does), UTF-8 with or without one, or cp1252."""
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16")
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1252")


def timestamp(value: Any) -> datetime | None:
    if not isinstance(value, str):
        return None
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError:
        return None
    return parsed if parsed.tzinfo is not None else None


@dataclass
class Validation:
    bundle: Path
    candidate: str
    policy: dict[str, Any]
    inventory: list[dict[str, str]]
    errors: list[str] = field(default_factory=list)
    outcomes: list[str] = field(default_factory=list)
    statuses: dict[str, str] = field(default_factory=dict)

    def error(self, message: str) -> None:
        self.errors.append(message)

    def log(self, stage: str, reference: Any) -> str | None:
        """The text of a cited log, after its path and digest are checked."""
        if not isinstance(reference, dict) or set(reference) != {"path", "sha256"}:
            self.error(f"{stage}: log must be an object with path and sha256")
            return None
        relative = reference["path"]
        if not isinstance(relative, str) or not relative or Path(relative).is_absolute() or ".." in Path(relative).parts:
            self.error(f"{stage}: log path {relative!r} must be relative and stay inside the bundle")
            return None
        path = self.bundle / relative
        current = self.bundle
        for part in Path(relative).parts:
            current = current / part
            if current.is_symlink():
                self.error(f"{stage}: log path {relative} passes through a symbolic link")
                return None
        if not path.is_file():
            self.error(f"{stage}: log {relative} not found in the bundle")
            return None
        data = path.read_bytes()
        if sha256(data) != reference["sha256"]:
            self.error(f"{stage}: log {relative} changed: its SHA-256 is not the one recorded")
            return None
        try:
            return decode_log(data)
        except UnicodeDecodeError as error:
            self.error(f"{stage}: log {relative} cannot be decoded: {error}")
            return None

    def candidate_line(self, stage: str, text: str) -> None:
        values = [match.group(1).strip() for match in map(CANDIDATE_LINE.match, text.splitlines()) if match]
        if len(values) != 1:
            self.error(f"{stage}: the log must hold exactly one CANDIDATE= line, found {len(values)}")
        elif values[0] != self.candidate:
            self.error(f"{stage}: the log names CANDIDATE={values[0]}, not the candidate {self.candidate}")


def check_environment(check: Validation, environment: Any) -> None:
    if not isinstance(environment, dict):
        check.error("environment must be an object")
        return
    for key in ENVIRONMENT_TEXT:
        if not isinstance(environment.get(key), str) or not environment[key].strip():
            check.error(f"environment.{key} must be a nonempty string")
    if environment.get("office_bitness") not in ("32-bit", "64-bit"):
        check.error("environment.office_bitness must be 32-bit or 64-bit")
    if not str(environment.get("os", "")).startswith("Windows"):
        check.error("environment.os must identify a Windows host")
    references = environment.get("references")
    if (not isinstance(references, list) or not all(isinstance(item, str) for item in references)
            or sorted(references) != sorted(check.policy["required_references"])):
        check.error("environment.references must be exactly the policy's required references")
    # The build script needs "Trust access to the VBA project object model".
    # Turning it on for the run is allowed, but the record must say so.
    changes = environment.get("trust_changes")
    if not isinstance(changes, bool):
        check.error("environment.trust_changes must be true or false")
    elif changes and (not isinstance(environment.get("trust_changes_detail"), str)
                      or not environment["trust_changes_detail"].strip()):
        check.error("environment.trust_changes_detail must say which Trust Center setting was changed, and back")


def check_stage(check: Validation, spec: dict[str, Any], record: Any, logs: dict[str, str]) -> None:
    name, kind = spec["name"], spec["kind"]
    if not isinstance(record, dict) or record.get("status") not in STATUSES:
        check.error(f"{name}: status must be one of {', '.join(STATUSES)}")
        return
    status = record["status"]
    check.statuses[name] = status
    if not isinstance(record.get("detail"), str) or not record["detail"].strip():
        check.error(f"{name}: detail must say what was done or observed")
    if status == "NOT_RUN":
        if record.get("log") is not None:
            check.error(f"{name}: a stage that did not run cites no log")
        if kind == "regression" and record.get("items") is not None:
            check.error(f"{name}: a regression that did not run claims no items")
        return
    text = check.log(name, record.get("log"))
    if text is None:
        return
    if kind in ("build", "compare"):
        normalized = sha256("\n".join(line.rstrip() for line in text.strip().splitlines()).encode("utf-8"))
        if normalized in logs:
            check.error(f"{name}: cites the same run as {logs[normalized]}; every build and comparison has its own log")
        logs[normalized] = name
    lines = text.splitlines()
    if kind in ("build", "manual", "regression"):
        check.candidate_line(name, text)
    if kind == "build":
        saved = any(SAVED_LINE.match(line) for line in lines)
        if status == "PASS" and not saved:
            check.error(f"{name}: a passing build log must show the saved workbook ('Saved ...')")
        if status != "PASS" and saved:
            check.error(f"{name}: the log shows a saved workbook, so the build is not {status}")
    elif kind == "regression":
        check_regression(check, name, status, record.get("items"), lines)
    elif kind == "compare":
        check_compare(check, spec, status, record, lines)


def check_regression(check: Validation, name: str, status: str, items: Any, lines: list[str]) -> None:
    results = [match for match in map(REGRESSION_LINE.match, lines) if match]
    expected = check.policy["regression_items"]
    if status == "TIMEOUT":
        if results:
            check.error(f"{name}: the log shows a regression result, so the run finished")
        if items is not None:
            check.error(f"{name}: a timed-out regression claims no items")
        return
    if len(results) != 1:
        check.error(f"{name}: the log must hold exactly one 'Regression test: Result:' line, found {len(results)}")
        return
    match = results[0]
    if match.group(1):
        observed = {"total": int(match.group(2)), "failures": 0}
    else:
        observed = {"total": int(match.group(5)), "failures": int(match.group(4))}
    if items != observed:
        check.error(f"{name}: items {items} do not match the log ({observed})")
    if (status == "PASS") != (observed["failures"] == 0):
        check.error(f"{name}: status {status} contradicts the log ({observed['failures']} failures)")
    if observed["total"] != expected:
        check.error(f"{name}: the log reports {observed['total']} items; the policy at the candidate expects {expected}")


def check_compare(check: Validation, spec: dict[str, Any], status: str, record: dict[str, Any],
                  lines: list[str]) -> None:
    name = spec["name"]

    def single(pattern: re.Pattern[str], label: str) -> re.Match[str] | None:
        # A comparison abandoned as hung may have stopped before printing
        # its header lines, so they are required only of a finished run.
        found = [match for match in map(pattern.match, lines) if match]
        if len(found) == 1:
            return found[0]
        if found or status != "TIMEOUT":
            check.error(f"{name}: the comparison log must hold exactly one {label} line, found {len(found)}")
        return None

    candidate = single(COMPARE_CANDIDATE, "Candidate:")
    if candidate and candidate.group(1).strip() != check.candidate:
        check.error(f"{name}: the comparison was run against {candidate.group(1).strip()}, not the candidate")
    mode = single(COMPARE_MODE, "Mode:")
    if mode and mode.group(1) != spec["mode"]:
        check.error(f"{name}: the comparison ran in {mode.group(1)} mode, the stage needs {spec['mode']}")
    template = single(COMPARE_TEMPLATE, "Template:")
    expected_template = next(item["sha256"] for item in check.inventory if item["path"] == TEMPLATE)
    if template is not None and (template.group(1) != TEMPLATE or template.group(2) != expected_template):
        check.error(f"{name}: the comparison names another template than the candidate's")
    results = [match for match in map(COMPARE_RESULT.match, lines) if match]
    if status == "TIMEOUT":
        if results:
            check.error(f"{name}: the log shows a RESULT line, so the comparison finished")
        return
    if len(results) != 1:
        check.error(f"{name}: the comparison log must hold exactly one RESULT= line, found {len(results)}")
        return
    result, binding = results[0].group(1), results[0].group(2)
    if status == "PASS" and (result != "PASS" or binding != "VERIFIED"):
        check.error(f"{name}: a passing round trip needs RESULT=PASS and binding=VERIFIED, "
                    f"the log shows {result} and {binding}")
    if status == "FAIL" and result != "FAIL":
        check.error(f"{name}: status FAIL contradicts RESULT={result} in the log")
    workbook = [match.group(1) for match in map(COMPARE_WORKBOOK.match, lines) if match]
    if status == "PASS" and (len(workbook) != 1 or record.get("workbook_sha256") != workbook[0]):
        check.error(f"{name}: workbook_sha256 must equal the workbook digest in the comparison log")


def validate(root: Path, candidate: str, evidence: Path) -> Validation:
    sha = resolve_commit(root, candidate)
    policy = load_policy(committed(root, sha, POLICY).decode("utf-8"), f"{POLICY} at {sha}")
    check = Validation(evidence.parent, sha, policy, inventory(root, sha))
    try:
        record = json.loads(evidence.read_bytes().decode("utf-8-sig"))
    except (OSError, UnicodeDecodeError, ValueError) as error:
        check.error(f"{evidence.name}: not a readable JSON record: {error}")
        check.outcomes = ["EVIDENCE_INVALID"]
        return check
    if not isinstance(record, dict):
        check.error(f"{evidence.name}: the record must be a JSON object")
        check.outcomes = ["EVIDENCE_INVALID"]
        return check
    try:
        check_record(check, record)
    except (TypeError, AttributeError, KeyError, ValueError, OSError) as error:
        # A malformed record is invalid evidence, never a traceback.
        check.error(f"{evidence.name}: malformed record: {error!r}")
    if check.errors:
        check.outcomes = ["EVIDENCE_INVALID"]
    return check


def check_record(  # noqa: C901 -- one branch per rule of docs/EXCEL_EVIDENCE.md, in its order
    check: Validation, record: dict[str, Any],
) -> None:
    policy, sha = check.policy, check.candidate
    if record.get("schema_version") != SCHEMA_VERSION:
        check.error(f"schema_version must be {SCHEMA_VERSION}")
    # A record made before a rename names the repository as it was (#261)
    if record.get("repository") not in (policy["repository"], *policy.get("former_repositories", [])):
        check.error(f"repository must be {policy['repository']}")
    if not isinstance(record.get("candidate_sha"), str) or not SHA.match(record["candidate_sha"]):
        check.error("candidate_sha must be a full 40-character commit SHA")
    elif record["candidate_sha"] != sha:
        check.error(f"candidate_sha {record['candidate_sha']} is not the candidate {sha}")
    started, finished = timestamp(record.get("started_at")), timestamp(record.get("finished_at"))
    if started is None or finished is None:
        check.error("started_at and finished_at must be ISO 8601 times with a UTC offset")
    elif finished < started:
        check.error("finished_at precedes started_at")

    execution = record.get("execution")
    if execution == "unavailable":
        if not isinstance(record.get("availability_reason"), str) or not record["availability_reason"].strip():
            check.error("an unavailable record needs an availability_reason")
        for key, empty in (("operator", None), ("environment", None), ("sources", []), ("stages", {})):
            if record.get(key) != empty:
                check.error(f"an unavailable record has {key} set to {json.dumps(empty)}")
        check.outcomes = ["EVIDENCE_INVALID"] if check.errors else ["UNAVAILABLE"]
        return
    if execution != "manual":
        check.error("execution must be manual or unavailable")
    if record.get("availability_reason") is not None:
        check.error("availability_reason must be null for a manual run")
    if not isinstance(record.get("operator"), str) or not record["operator"].strip():
        check.error("operator must name who ran Excel, and where")
    check_environment(check, record.get("environment"))
    sources = record.get("sources")
    if not isinstance(sources, list) or sorted(map(json.dumps, sources)) != sorted(map(json.dumps, check.inventory)):
        check.error("sources must equal the --inventory of the candidate: the record is for other source")

    stages = record.get("stages")
    names = [stage["name"] for stage in policy["stages"]]
    if not isinstance(stages, dict) or set(stages) != set(names):
        check.error(f"stages must be exactly {', '.join(names)}")
        stages = stages if isinstance(stages, dict) else {}
    logs: dict[str, str] = {}
    for spec in policy["stages"]:
        if spec["name"] in stages:
            check_stage(check, spec, stages[spec["name"]], logs)
    for spec in policy["stages"]:
        status = check.statuses.get(spec["name"])
        if status not in (None, "NOT_RUN"):
            blocked = [dep for dep in spec["depends_on"] if check.statuses.get(dep) != "PASS"]
            if blocked:
                check.error(f"{spec['name']} ran although {', '.join(blocked)} did not pass")

    outcomes: list[str] = []
    for spec in policy["stages"]:
        status = check.statuses.get(spec["name"])
        if status == "FAIL" and FAILED_OUTCOME[spec["kind"]] not in outcomes:
            outcomes.append(FAILED_OUTCOME[spec["kind"]])
    if "TIMEOUT" in check.statuses.values():
        outcomes.append("EXECUTION_TIMEOUT")
    if "NOT_RUN" in check.statuses.values():
        outcomes.append("INCOMPLETE")
    check.outcomes = outcomes or ["PASS"]


def write_reports(check: Validation, evidence: Path, output: Path | None, summary: Path | None) -> None:
    bundle = evidence.parent.resolve()
    for path in (output, summary):
        if path is not None and bundle in (path.resolve(), *path.resolve().parents):
            raise CommandError(f"{path}: reports must be written outside the bundle")
    report = {"candidate_sha": check.candidate, "outcomes": check.outcomes,
              "stages": check.statuses, "errors": check.errors}
    if output is not None:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    if summary is not None:
        rows = "\n".join(f"| `{name}` | {status} |" for name, status in check.statuses.items())
        errors = "\n".join(f"- {error}" for error in check.errors) or "None."
        summary.parent.mkdir(parents=True, exist_ok=True)
        summary.write_text(
            f"# Excel evidence\n\nCandidate: `{check.candidate}`\n\n"
            f"Outcome: **{', '.join(check.outcomes)}**\n\n"
            f"| Stage | Status |\n| --- | --- |\n{rows}\n\n## Errors\n\n{errors}\n\n"
            "A validated record checks identity and consistency; it does not prove what was done.\n",
            encoding="utf-8")


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=REPO_ROOT)
    parser.add_argument("--candidate-sha", help="the commit the workbook was built from")
    action = parser.add_mutually_exclusive_group(required=True)
    action.add_argument("--inventory", action="store_true", help="print the source digests of the candidate")
    action.add_argument("--check-policy", action="store_true", help="check the policy against the source")
    action.add_argument("--evidence", type=Path, help="host.json in the evidence bundle")
    parser.add_argument("--output", type=Path, help="JSON report, outside the bundle")
    parser.add_argument("--summary", type=Path, help="Markdown summary, outside the bundle")
    options = parser.parse_args(arguments)
    try:
        if options.check_policy:
            findings = check_policy(options.root)
            for finding in findings:
                print(f"FAIL {finding}")
            print(f"RESULT={'FAIL' if findings else 'PASS'}")
            return 1 if findings else 0
        if not options.candidate_sha:
            raise CommandError("--candidate-sha is required")
        if options.inventory:
            print(json.dumps(inventory(options.root, resolve_commit(options.root, options.candidate_sha)), indent=2))
            return 0
        if not options.evidence.is_file():
            raise CommandError(f"evidence not found: {options.evidence}")
        check = validate(options.root, options.candidate_sha, options.evidence)
        write_reports(check, options.evidence, options.output, options.summary)
    except CommandError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    print(f"Candidate: {check.candidate}")
    for name, status in check.statuses.items():
        print(f"  {name:<24} {status}")
    for message in check.errors:
        print(f"  INVALID {message}")
    print(f"OUTCOME={', '.join(check.outcomes)}")
    return 0 if check.outcomes == ["PASS"] else 1


if __name__ == "__main__":
    sys.exit(main())
