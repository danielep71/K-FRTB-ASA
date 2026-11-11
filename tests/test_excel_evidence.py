"""Offline tests of tools/check_excel_evidence.py with synthetic bundles.

A temporary Git repository holds a copy of the policy, the modules and the
template, so the tests do not depend on the state of the working tree. The
bundles here are synthetic: they test the validator, and are never Excel
evidence.
"""

import copy
import hashlib
import io
import json
import shutil
import subprocess
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

import pytest

import check_excel_evidence as tool

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_DEV = "a" * 64
WORKBOOK_PROD = "b" * 64


def git(repo, *args):
    return subprocess.run(["git", "-C", str(repo), "-c", "user.name=Test", "-c", "user.email=test@example.invalid",
                           *args], capture_output=True, check=True, text=True).stdout.strip()


@pytest.fixture(scope="module")
def repo(tmp_path_factory):
    repo = tmp_path_factory.mktemp("repo")
    for folder, pattern in (("src/modules", "*.bas"), ("tests/modules", "*.bas")):
        (repo / folder).mkdir(parents=True)
        for path in (ROOT / folder).glob(pattern):
            shutil.copyfile(path, repo / folder / path.name)
    shutil.copyfile(ROOT / ".gitattributes", repo / ".gitattributes")  # the same LF storage as the repository
    (repo / ".github").mkdir()
    shutil.copyfile(ROOT / tool.POLICY, repo / tool.POLICY)
    (repo / "src/workbook").mkdir()
    shutil.copyfile(ROOT / tool.TEMPLATE, repo / tool.TEMPLATE)
    git(repo, "init", "-q")
    git(repo, "add", "-A")
    git(repo, "commit", "-q", "-m", "fixture")
    first = git(repo, "rev-parse", "HEAD")
    (repo / "src/modules/CORE_SBM.bas").write_bytes((repo / "src/modules/CORE_SBM.bas").read_bytes() + b"' later\n")
    git(repo, "commit", "-q", "-am", "later")
    return {"path": repo, "sha": first, "later": git(repo, "rev-parse", "HEAD")}


def digest(data):
    return hashlib.sha256(data).hexdigest()


def compare_log(sha, mode, template, workbook, result="PASS", binding="VERIFIED"):
    return (f"Candidate: {sha}\nMode:      {mode} (11 modules expected, 12 components exported)\n"
            f"Template:  {tool.TEMPLATE} sha256 {template}\nWorkbook:  K-FRTB-ASA.xlsm sha256 {workbook}\n"
            "Re-cased:  0 line(s) differ only in the letter case of code (ignored)\n"
            f"Binding:   {binding} - export-manifest.json records this workbook and every exported file\n"
            f"RESULT={result}; modules=11; differences={0 if result == 'PASS' else 1}; binding={binding}\n")


class Bundle:
    """A passing synthetic bundle, changed by each test before it is validated."""

    def __init__(self, folder, repo):
        self.folder = folder
        self.repo = repo
        sha = repo["sha"]
        self.inventory = tool.inventory(repo["path"], sha)
        template = next(item["sha256"] for item in self.inventory if item["path"] == tool.TEMPLATE)
        self.logs = {
            "session.log": f"CANDIDATE={sha}\nCompile development: no error\nCompile production: no error\n",
            "build-development.log": (f"CANDIDATE={sha}\nOpened a copy of src\\workbook\\FRTB_Template.xlsx\n"
                                      "Imported CORE_Config.bas\nRegression test: Result: PASS (52 items)\n"
                                      "Saved C:\\FRTB\\dist\\K-FRTB-ASA.xlsm\n"),
            "build-production.log": f"CANDIDATE={sha}\nImported CORE_Config.bas\nSaved C:\\FRTB\\dist\\prod.xlsm\n",
            "round-trip-development.log": compare_log(sha, "development", template, WORKBOOK_DEV),
            "round-trip-production.log": compare_log(sha, "production", template, WORKBOOK_PROD),
        }
        self.record = {
            "schema_version": 1, "repository": "danielep71/K-FRTB-ASA", "candidate_sha": sha,
            "execution": "manual", "availability_reason": None,
            "started_at": "2026-10-20T09:00:00+02:00", "finished_at": "2026-10-20T10:00:00+02:00",
            "operator": "Owner, workstation",
            "environment": {
                "excel_version": "16.0", "excel_build": "Version 2609 (Build 16.0.20000.20000)",
                "office_bitness": "64-bit", "os": "Windows 11 Pro 24H2", "os_architecture": "x64",
                "decimal_separator": ",", "list_separator": ";",
                "references": ["Visual Basic For Applications", "Microsoft Excel 16.0 Object Library",
                               "OLE Automation", "Microsoft Office 16.0 Object Library"],
                "macro_policy": "Disable VBA macros with notification",
                "vba_project_access": "On, as found", "trust_changes": False,
            },
            "sources": copy.deepcopy(self.inventory),
            "stages": {
                "build_development": self.stage("build-development.log"),
                "compile_development": self.stage("session.log"),
                "regression": {**self.stage("build-development.log"), "items": {"total": 52, "failures": 0}},
                "round_trip_development": {**self.stage("round-trip-development.log"), "workbook_sha256": WORKBOOK_DEV},
                "build_production": self.stage("build-production.log"),
                "compile_production": self.stage("session.log"),
                "round_trip_production": {**self.stage("round-trip-production.log"), "workbook_sha256": WORKBOOK_PROD},
            },
        }

    def stage(self, log, status="PASS"):
        return {"status": status, "detail": "observed", "log": {"path": log, "sha256": None}}

    def write(self):
        for name, text in self.logs.items():
            (self.folder / name).write_text(text, encoding="utf-8")
        record = copy.deepcopy(self.record)
        for stage in record["stages"].values() if isinstance(record.get("stages"), dict) else []:
            log = stage.get("log") if isinstance(stage, dict) else None
            if isinstance(log, dict) and log.get("sha256") is None and (self.folder / str(log["path"])).is_file():
                log["sha256"] = digest((self.folder / log["path"]).read_bytes())
        path = self.folder / "host.json"
        path.write_text(json.dumps(record, indent=2), encoding="utf-8")
        return path

    def validate(self, candidate=None):
        return tool.validate(self.repo["path"], candidate or self.repo["sha"], self.write())


@pytest.fixture
def bundle(tmp_path, repo):
    folder = tmp_path / "bundle"
    folder.mkdir()
    return Bundle(folder, repo)


def assert_invalid(check, fragment):
    assert check.outcomes == ["EVIDENCE_INVALID"], check.errors
    assert any(fragment in error for error in check.errors), check.errors


def test_a_complete_record_passes(bundle):
    check = bundle.validate()
    assert check.errors == []
    assert check.outcomes == ["PASS"]
    assert set(check.statuses.values()) == {"PASS"}


def test_a_record_under_the_former_repository_name_passes(bundle):
    # The 323260b record names danielep71/K-FRTB-SA, the name before #261
    bundle.record["repository"] = "danielep71/K-FRTB-SA"
    check = bundle.validate()
    assert check.errors == []
    assert check.outcomes == ["PASS"]


def test_a_record_of_another_repository_is_invalid(bundle):
    bundle.record["repository"] = "danielep71/another-repository"
    assert_invalid(bundle.validate(), "repository must be danielep71/K-FRTB-ASA")


def test_policy_matches_the_source():
    assert tool.check_policy(ROOT) == []
    assert tool.generated_items((ROOT / tool.SETUP_MODULE).read_text(encoding="cp1252")) == 52


def test_inventory_is_of_the_git_blobs(repo):
    items = tool.inventory(repo["path"], repo["sha"])
    assert items[-1]["path"] == tool.TEMPLATE
    blob = subprocess.run(["git", "-C", str(repo["path"]), "show", f"{repo['sha']}:src/modules/CORE_SBM.bas"],
                          capture_output=True, check=True).stdout
    assert {"path": "src/modules/CORE_SBM.bas", "sha256": digest(blob)} in items
    assert b"\r\n" not in blob


def test_sources_of_another_commit_are_invalid(bundle, repo):
    bundle.record["candidate_sha"] = repo["later"]
    for name in ("build-development.log", "build-production.log", "session.log"):
        bundle.logs[name] = bundle.logs[name].replace(repo["sha"], repo["later"])
    check = bundle.validate(repo["later"])
    assert_invalid(check, "sources must equal the --inventory")


def test_a_record_for_another_candidate_is_invalid(bundle, repo):
    assert_invalid(bundle.validate(repo["later"]), "is not the candidate")


@pytest.mark.parametrize(
    ("log", "old", "new", "fragment"),
    [
        ("build-production.log", "CANDIDATE=", "COMMIT=", "exactly one CANDIDATE= line, found 0"),
        ("session.log", "CANDIDATE=", "CANDIDATE=0000", "not the candidate"),
        ("build-development.log", "Saved ", "Not saved ", "must show the saved workbook"),
        ("build-development.log", "PASS (52 items)", "PASS (51 items)", "the policy at the candidate expects 52"),
        ("round-trip-production.log", "Mode:      production", "Mode:      development", "the stage needs production"),
        ("round-trip-development.log", "binding=VERIFIED", "binding=UNVERIFIED", "binding=VERIFIED"),
        ("round-trip-development.log", "RESULT=PASS", "RESULT=FAIL", "RESULT=PASS and binding=VERIFIED"),
        ("round-trip-development.log", "Candidate: ", "Candidate: 0", "not the candidate"),
    ],
    ids=["no candidate line", "other candidate", "unsaved build", "stale item count",
         "wrong mode", "unverified binding", "failed comparison recorded as pass", "other comparison candidate"],
)
def test_logs_that_contradict_the_record_are_invalid(bundle, log, old, new, fragment):
    assert old in bundle.logs[log]
    bundle.logs[log] = bundle.logs[log].replace(old, new, 1)
    assert_invalid(bundle.validate(), fragment)


def test_another_template_in_the_comparison_is_invalid(bundle):
    log = bundle.logs["round-trip-production.log"]
    template = next(item["sha256"] for item in bundle.inventory if item["path"] == tool.TEMPLATE)
    bundle.logs["round-trip-production.log"] = log.replace(template, "c" * 64)
    assert_invalid(bundle.validate(), "another template")


def test_a_tampered_log_is_invalid(bundle):
    path = bundle.write()
    (bundle.folder / "session.log").write_text(bundle.logs["session.log"] + "edited\n", encoding="utf-8")
    assert_invalid(tool.validate(bundle.repo["path"], bundle.repo["sha"], path), "changed: its SHA-256")


@pytest.mark.parametrize(
    ("path", "fragment"),
    [("missing.log", "not found in the bundle"), ("../outside.log", "must be relative"),
     ("/tmp/absolute.log", "must be relative")],
)
def test_a_missing_or_escaping_log_is_invalid(bundle, path, fragment):
    bundle.record["stages"]["compile_production"]["log"] = {"path": path, "sha256": "d" * 64}
    assert_invalid(bundle.validate(), fragment)


def test_a_symlinked_log_is_invalid(bundle, tmp_path):
    outside = tmp_path / "outside.log"
    outside.write_text(bundle.logs["session.log"], encoding="utf-8")
    (bundle.folder / "linked.log").symlink_to(outside)
    bundle.record["stages"]["compile_production"]["log"] = {"path": "linked.log", "sha256": digest(outside.read_bytes())}
    assert_invalid(bundle.validate(), "symbolic link")


def test_the_same_comparison_cited_twice_is_invalid(bundle):
    bundle.logs["round-trip-production.log"] = bundle.logs["round-trip-development.log"] + "\n\n"
    bundle.record["stages"]["round_trip_production"]["workbook_sha256"] = WORKBOOK_DEV
    check = bundle.validate()
    assert_invalid(check, "cites the same run as round_trip_development")


def test_a_failed_regression_is_test_failed(bundle):
    bundle.logs["build-development.log"] = bundle.logs["build-development.log"].replace(
        "Result: PASS (52 items)", "Result: FAIL (2 of 52 items)")
    bundle.record["stages"]["regression"].update(status="FAIL", items={"total": 52, "failures": 2})
    check = bundle.validate()
    assert check.errors == []
    assert check.outcomes == ["TEST_FAILED"]


def test_a_failed_regression_with_wrong_counts_is_invalid(bundle):
    bundle.logs["build-development.log"] = bundle.logs["build-development.log"].replace(
        "Result: PASS (52 items)", "Result: FAIL (2 of 52 items)")
    bundle.record["stages"]["regression"].update(status="FAIL", items={"total": 52, "failures": 1})
    assert_invalid(bundle.validate(), "do not match the log")


def test_a_failed_round_trip_is_round_trip_failed(bundle):
    template = next(item["sha256"] for item in bundle.inventory if item["path"] == tool.TEMPLATE)
    bundle.logs["round-trip-production.log"] = compare_log(bundle.repo["sha"], "production", template,
                                                           WORKBOOK_PROD, result="FAIL")
    bundle.record["stages"]["round_trip_production"]["status"] = "FAIL"
    check = bundle.validate()
    assert check.errors == []
    assert check.outcomes == ["ROUND_TRIP_FAILED"]


def test_a_failed_compile_blocks_the_regression(bundle):
    bundle.record["stages"]["compile_development"]["status"] = "FAIL"
    assert_invalid(bundle.validate(), "regression ran although compile_development did not pass")


def test_a_failed_compile_with_stages_not_run_is_compile_failed_and_incomplete(bundle):
    stages = bundle.record["stages"]
    stages["compile_development"]["status"] = "FAIL"
    stages["regression"] = {"status": "NOT_RUN", "detail": "compile failed", "log": None, "items": None}
    check = bundle.validate()
    assert check.errors == []
    assert check.outcomes == ["COMPILE_FAILED", "INCOMPLETE"]


def test_a_failed_build_is_build_failed(bundle):
    bundle.logs["build-production.log"] = bundle.logs["build-production.log"].replace("Saved", "Error: Excel denied access")
    stages = bundle.record["stages"]
    stages["build_production"]["status"] = "FAIL"
    for name in ("compile_production", "round_trip_production"):
        stages[name] = {"status": "NOT_RUN", "detail": "build failed", "log": None}
    check = bundle.validate()
    assert check.errors == []
    assert check.outcomes == ["BUILD_FAILED", "INCOMPLETE"]


def test_a_timed_out_comparison_is_execution_timeout(bundle):
    bundle.logs["round-trip-production.log"] = bundle.logs["round-trip-production.log"].split("RESULT=")[0]
    bundle.record["stages"]["round_trip_production"]["status"] = "TIMEOUT"
    check = bundle.validate()
    assert check.errors == []
    assert check.outcomes == ["EXECUTION_TIMEOUT"]


def test_a_timeout_that_finished_is_invalid(bundle):
    bundle.record["stages"]["round_trip_production"]["status"] = "TIMEOUT"
    assert_invalid(bundle.validate(), "so the comparison finished")


def test_a_stage_not_run_cites_no_log(bundle):
    bundle.record["stages"]["compile_production"]["status"] = "NOT_RUN"
    assert_invalid(bundle.validate(), "cites no log")


@pytest.mark.parametrize(
    ("key", "value", "fragment"),
    [
        ("os", "macOS 15", "Windows host"),
        ("office_bitness", "64", "32-bit or 64-bit"),
        ("references", ["Visual Basic For Applications"], "required references"),
        ("trust_changes", True, "trust_changes_detail must say"),
        ("trust_changes", "no", "trust_changes must be true or false"),
        ("decimal_separator", "", "decimal_separator must be a nonempty string"),
    ],
)
def test_environment_errors_are_invalid(bundle, key, value, fragment):
    bundle.record["environment"][key] = value
    assert_invalid(bundle.validate(), fragment)


def test_a_declared_trust_change_is_accepted(bundle):
    bundle.record["environment"].update(trust_changes=True, trust_changes_detail=(
        "Trust access to the VBA project object model turned on for the build and off afterwards"))
    check = bundle.validate()
    assert check.outcomes == ["PASS"], check.errors


def test_times_without_offset_or_reversed_are_invalid(bundle):
    bundle.record["started_at"] = "2026-10-20T09:00:00"
    assert_invalid(bundle.validate(), "UTC offset")
    bundle.record["started_at"] = "2026-10-20T11:00:00+02:00"
    assert_invalid(bundle.validate(), "precedes")


def test_a_missing_stage_is_invalid(bundle):
    del bundle.record["stages"]["round_trip_production"]
    assert_invalid(bundle.validate(), "stages must be exactly")


def test_an_unavailable_record(bundle):
    bundle.record.update(execution="unavailable", availability_reason="No Windows host with Excel",
                         operator=None, environment=None, sources=[], stages={})
    check = bundle.validate()
    assert check.errors == []
    assert check.outcomes == ["UNAVAILABLE"]
    bundle.record["availability_reason"] = ""
    assert_invalid(bundle.validate(), "availability_reason")


def test_utf16_logs_from_powershell_redirection_are_read(bundle):
    path = bundle.write()
    log = bundle.folder / "round-trip-development.log"
    log.write_bytes(b"\xff\xfe" + bundle.logs["round-trip-development.log"].encode("utf-16-le"))
    record = json.loads(path.read_text(encoding="utf-8"))
    record["stages"]["round_trip_development"]["log"]["sha256"] = digest(log.read_bytes())
    path.write_text(json.dumps(record), encoding="utf-8")
    check = tool.validate(bundle.repo["path"], bundle.repo["sha"], path)
    assert check.outcomes == ["PASS"], check.errors


def run_tool(*arguments):
    out, err = io.StringIO(), io.StringIO()
    with redirect_stdout(out), redirect_stderr(err):
        code = tool.main(list(arguments))
    return code, out.getvalue() + err.getvalue()


def test_exit_codes_and_reports(bundle, tmp_path):
    path = bundle.write()
    root = str(bundle.repo["path"])
    output, summary = tmp_path / "out" / "report.json", tmp_path / "out" / "report.md"
    code, text = run_tool("--root", root, "--candidate-sha", bundle.repo["sha"], "--evidence", str(path),
                          "--output", str(output), "--summary", str(summary))
    assert code == 0, text
    assert "OUTCOME=PASS" in text
    assert json.loads(output.read_text(encoding="utf-8"))["outcomes"] == ["PASS"]
    assert "**PASS**" in summary.read_text(encoding="utf-8")
    assert run_tool("--root", root, "--candidate-sha", bundle.repo["later"], "--evidence", str(path))[0] == 1
    assert run_tool("--root", root, "--candidate-sha", "not-a-commit", "--evidence", str(path))[0] == 2
    assert run_tool("--root", root, "--candidate-sha", bundle.repo["sha"], "--evidence", str(path),
                    "--output", str(bundle.folder / "report.json"))[0] == 2
    assert run_tool("--root", root, "--evidence", str(path))[0] == 2
    assert run_tool("--root", root, "--candidate-sha", bundle.repo["sha"], "--inventory")[0] == 0
    assert run_tool("--check-policy")[0] == 0


def test_the_validator_reads_the_lines_the_tools_print(tmp_path):
    # The build script and the regression test print the lines the build and
    # regression stages parse; a change to either format must fail here.
    build = (ROOT / "tools" / "Build-Workbook.ps1").read_text(encoding="utf-8")
    assert 'Write-Host "Regression test: $result"' in build
    assert 'Write-Host "Saved $OutputPath"' in build
    regression = (ROOT / "tests" / "modules" / "TEST_Regression.bas").read_text(encoding="cp1252")
    assert '"Result: PASS (" & total & " items)"' in regression
    assert '"Result: FAIL (" & failures & " of " & total & " items)"' in regression
    assert tool.REGRESSION_LINE.match("Regression test: Result: PASS (52 items)")
    assert tool.REGRESSION_LINE.match("Regression test: Result: FAIL (2 of 52 items)")
    # The comparison tool prints the lines the round-trip stages parse.
    import compare_vba_export
    export = tmp_path / "export"
    export.mkdir()
    for folder in compare_vba_export.MODE_FOLDERS["production"]:
        for path in (ROOT / folder).glob("*.bas"):
            data = subprocess.run(["git", "-C", str(ROOT), "show", f"HEAD:{folder}/{path.name}"],
                                  capture_output=True, check=True).stdout
            (export / path.name).write_bytes(data.replace(b"\n", b"\r\n"))
    out = io.StringIO()
    with redirect_stdout(out):
        assert compare_vba_export.main(["--root", str(ROOT), "--export", str(export), "--mode", "production"]) == 0
    lines = out.getvalue().splitlines()
    for pattern in (tool.COMPARE_CANDIDATE, tool.COMPARE_MODE, tool.COMPARE_TEMPLATE, tool.COMPARE_RESULT):
        assert sum(bool(pattern.match(line)) for line in lines) == 1, (pattern.pattern, lines)


def test_an_undecodable_utf16_log_is_invalid(bundle):
    path = bundle.write()
    log = bundle.folder / "round-trip-development.log"
    log.write_bytes(b"\xff\xfe" + "Candidate: x\n".encode("utf-16-le") + b"\x00")  # an odd byte count
    record = json.loads(path.read_text(encoding="utf-8"))
    record["stages"]["round_trip_development"]["log"]["sha256"] = digest(log.read_bytes())
    path.write_text(json.dumps(record), encoding="utf-8")
    assert_invalid(tool.validate(bundle.repo["path"], bundle.repo["sha"], path), "cannot be decoded")


def test_non_string_references_are_invalid(bundle):
    bundle.record["environment"]["references"] = ["Visual Basic For Applications", None]
    assert_invalid(bundle.validate(), "required references")


def test_a_comparison_that_hung_before_its_header_is_a_timeout(bundle):
    bundle.logs["round-trip-production.log"] = "Comparing...\n"
    bundle.record["stages"]["round_trip_production"]["status"] = "TIMEOUT"
    check = bundle.validate()
    assert check.errors == []
    assert check.outcomes == ["EXECUTION_TIMEOUT"]


def test_a_timed_out_comparison_still_names_the_right_candidate(bundle):
    log = bundle.logs["round-trip-production.log"].split("RESULT=")[0]
    bundle.logs["round-trip-production.log"] = log.replace("Candidate: ", "Candidate: 0", 1)
    bundle.record["stages"]["round_trip_production"]["status"] = "TIMEOUT"
    assert_invalid(bundle.validate(), "not the candidate")


def test_an_unexpected_record_shape_is_invalid_not_a_crash(bundle, monkeypatch):
    def broken(*_):
        raise TypeError("unexpected shape")
    monkeypatch.setattr(tool, "check_stage", broken)
    assert_invalid(bundle.validate(), "malformed record")
