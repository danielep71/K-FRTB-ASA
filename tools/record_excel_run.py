"""Write host.json for an Excel run bundle and validate it.

tools/Record-FRTB-ExcelRun.ps1 runs the stages of the evidence policy in
desktop Excel and leaves, in one bundle folder, the build, compile and
comparison logs and session.log: the candidate, the times, the operator and
the environment read from Excel and the registry, one KEY=value per line.
This script reads that bundle, decides the status of every stage from the
logs only, writes host.json in the format of docs/EXCEL_EVIDENCE.md, and runs
tools/check_excel_evidence.py on it. Nothing is typed by hand, so no field
holds a value that the run did not produce.

A stage passes only when its own log shows it: a build its "Saved" line and
exit code 0 (or 1 after a failed regression, which the build saves), a compile its PASS line in session.log, the regression its
result line, a comparison RESULT=PASS with binding=VERIFIED. A stage whose
prerequisite did not pass is NOT_RUN, as the policy requires.

    python tools/record_excel_run.py BUNDLE [--candidate SHA]

Exit code 0 means the validator reports PASS, 1 any other outcome, 2 a
command error. It never runs Excel.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

import check_excel_evidence as validator

ROOT = Path(__file__).resolve().parents[1]
SESSION = "session.log"
SESSION_LINE = re.compile(r"^([A-Z][A-Z0-9_]*)=(.*)$")
ENVIRONMENT = {
    "excel_version": "EXCEL_VERSION", "excel_build": "EXCEL_BUILD", "office_bitness": "OFFICE_BITNESS",
    "os": "OS", "os_architecture": "OS_ARCHITECTURE", "decimal_separator": "DECIMAL_SEPARATOR",
    "list_separator": "LIST_SEPARATOR", "macro_policy": "MACRO_POLICY",
    "vba_project_access": "VBA_PROJECT_ACCESS",
}
# The stage of the policy, its log and the session.log keys that report it.
LOGS = {
    "build_development": "build-development.log",
    "compile_development": SESSION,
    "regression": "build-development.log",
    "round_trip_development": "round-trip-development.log",
    "build_production": "build-production.log",
    "compile_production": SESSION,
    "round_trip_production": "round-trip-production.log",
}
BUILD_EXIT = {"build_development": "BUILD_DEVELOPMENT_EXIT", "build_production": "BUILD_PRODUCTION_EXIT"}
COMPILE = {"compile_development": "COMPILE_DEVELOPMENT", "compile_production": "COMPILE_PRODUCTION"}


class RecordError(Exception):
    """The bundle cannot be read: exit code 2."""


def read_session(bundle: Path) -> dict[str, str]:
    path = bundle / SESSION
    if not path.is_file():
        raise RecordError(f"{path} not found: run tools/Record-FRTB-ExcelRun.ps1 first")
    values: dict[str, str] = {}
    for line in validator.decode_log(path.read_bytes()).splitlines():
        match = SESSION_LINE.match(line.strip())
        if match:
            # A later line replaces an earlier one, for example a retried compile
            values[match.group(1)] = match.group(2).strip()
    return values


def log_text(bundle: Path, name: str) -> str | None:
    path = bundle / name
    return validator.decode_log(path.read_bytes()) if path.is_file() else None


def reference(bundle: Path, name: str) -> dict[str, str]:
    return {"path": name, "sha256": hashlib.sha256((bundle / name).read_bytes()).hexdigest()}


def not_run(why: str, regression: bool = False) -> dict[str, Any]:
    stage: dict[str, Any] = {"status": "NOT_RUN", "detail": why, "log": None}
    if regression:
        stage["items"] = None
    return stage


def build_stage(bundle: Path, name: str, session: dict[str, str]) -> dict[str, Any]:
    text = log_text(bundle, LOGS[name])
    if text is None:
        return not_run(f"{LOGS[name]} not found: the build did not run")
    lines = text.splitlines()
    saved = any(validator.SAVED_LINE.match(line) for line in lines)
    code = session.get(BUILD_EXIT[name])
    # Build-Workbook.ps1 saves the workbook and exits 1 when the regression
    # fails: the build passed, and the regression stage records the failure.
    failed_tests = [match for match in map(validator.REGRESSION_LINE.match, lines) if match and match.group(3)]
    passed = saved and (code == "0" or (code == "1" and len(failed_tests) == 1))
    detail = f"Build-Workbook.ps1, exit code {code or 'not recorded'}"
    return {"status": "PASS" if passed else "FAIL", "detail": detail, "log": reference(bundle, LOGS[name])}


def compile_stage(bundle: Path, name: str, session: dict[str, str], build: dict[str, Any]) -> dict[str, Any]:
    if build["status"] != "PASS":
        return not_run("the build did not pass, so the workbook was not compiled")
    value = session.get(COMPILE[name])
    if value not in ("PASS", "FAIL"):
        return not_run(f"{COMPILE[name]} not recorded in session.log")
    detail = session.get(COMPILE[name] + "_DETAIL") or "Debug > Compile VBAProject"
    return {"status": value, "detail": detail, "log": reference(bundle, SESSION)}


def regression_stage(bundle: Path, compiled: dict[str, Any]) -> dict[str, Any]:
    if compiled["status"] != "PASS":
        return not_run("the development compile did not pass", regression=True)
    text = log_text(bundle, LOGS["regression"]) or ""
    results = [match for match in map(validator.REGRESSION_LINE.match, text.splitlines()) if match]
    if len(results) != 1:
        return not_run("the build log shows no single regression result line", regression=True)
    match = results[0]
    if match.group(1):
        status, items = "PASS", {"total": int(match.group(2)), "failures": 0}
    else:
        status, items = "FAIL", {"total": int(match.group(5)), "failures": int(match.group(4))}
    return {"status": status, "detail": "Test_RunRegression_Silent through -RunTests",
            "log": reference(bundle, LOGS["regression"]), "items": items}


def compare_stage(bundle: Path, name: str, build: dict[str, Any]) -> dict[str, Any]:
    if build["status"] != "PASS":
        return not_run("the build did not pass, so the workbook was not exported")
    text = log_text(bundle, LOGS[name])
    if text is None:
        return not_run(f"{LOGS[name]} not found: the comparison did not run")
    lines = text.splitlines()
    results = [match for match in map(validator.COMPARE_RESULT.match, lines) if match]
    passed = len(results) == 1 and results[0].group(1) == "PASS" and results[0].group(2) == "VERIFIED"
    if len(results) == 1:
        outcome = results[0].group(0)
    else:
        outcome = f"{len(results)} RESULT lines; see {LOGS[name]}"
    stage: dict[str, Any] = {"status": "PASS" if passed else "FAIL",
                             "detail": f"Export-FRTB-VBA.ps1 and compare_vba_export.py: {outcome}",
                             "log": reference(bundle, LOGS[name])}
    if passed:
        workbook = [match.group(1) for match in map(validator.COMPARE_WORKBOOK.match, lines) if match]
        stage["workbook_sha256"] = workbook[0] if len(workbook) == 1 else None
    return stage


def build_record(root: Path, bundle: Path, candidate: str) -> dict[str, Any]:
    session = read_session(bundle)
    named = session.get("CANDIDATE")
    if named != candidate:
        raise RecordError(f"session.log names CANDIDATE={named}, not {candidate}")
    environment: dict[str, Any] = {field: session.get(key, "") for field, key in ENVIRONMENT.items()}
    environment["references"] = [item for item in session.get("REFERENCES", "").split("|") if item]
    environment["trust_changes"] = bool(session.get("TRUST_CHANGES"))
    if environment["trust_changes"]:
        environment["trust_changes_detail"] = session["TRUST_CHANGES"]

    stages: dict[str, Any] = {}
    stages["build_development"] = build_stage(bundle, "build_development", session)
    stages["compile_development"] = compile_stage(bundle, "compile_development", session,
                                                  stages["build_development"])
    stages["regression"] = regression_stage(bundle, stages["compile_development"])
    stages["round_trip_development"] = compare_stage(bundle, "round_trip_development",
                                                     stages["build_development"])
    stages["build_production"] = build_stage(bundle, "build_production", session)
    stages["compile_production"] = compile_stage(bundle, "compile_production", session,
                                                 stages["build_production"])
    stages["round_trip_production"] = compare_stage(bundle, "round_trip_production",
                                                    stages["build_production"])
    policy = validator.load_policy(validator.committed(root, candidate, validator.POLICY).decode("utf-8"),
                                   validator.POLICY)
    return {
        "schema_version": validator.SCHEMA_VERSION,
        "repository": policy["repository"],
        "candidate_sha": candidate,
        "execution": "manual",
        "availability_reason": None,
        "started_at": session.get("STARTED", ""),
        "finished_at": session.get("FINISHED", ""),
        "operator": session.get("OPERATOR", ""),
        "environment": environment,
        "sources": validator.inventory(root, candidate),
        "stages": stages,
    }


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("bundle", type=Path, help="the bundle folder the recorder wrote")
    parser.add_argument("--candidate", help="the full commit SHA; default: the candidate in session.log")
    parser.add_argument("--root", type=Path, default=ROOT, help=argparse.SUPPRESS)
    options = parser.parse_args(arguments)
    bundle = options.bundle.resolve()
    try:
        candidate = options.candidate or read_session(bundle).get("CANDIDATE", "")
        candidate = validator.resolve_commit(options.root, candidate)
        record = build_record(options.root, bundle, candidate)
    except (RecordError, validator.CommandError, subprocess.CalledProcessError, OSError, ValueError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    host = bundle / "host.json"
    host.write_text(json.dumps(record, indent=2) + "\n", encoding="utf-8")
    reports = bundle.parent / f"{bundle.name}-validation"
    reports.mkdir(exist_ok=True)
    check = validator.validate(options.root, candidate, host)
    validator.write_reports(check, host, reports / "excel-evidence.json", reports / "excel-evidence.md")
    print(f"host.json written: {host}")
    for name, stage in record["stages"].items():
        print(f"  {name:24} {stage['status']:8} {stage['detail']}")
    print(f"Validator outcome: {', '.join(check.outcomes)}")
    for error in check.errors:
        print(f"  ERROR {error}")
    print(f"Reports: {reports}")
    return 0 if check.outcomes == ["PASS"] else 1


if __name__ == "__main__":
    sys.exit(main())
