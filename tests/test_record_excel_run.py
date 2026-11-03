"""Offline tests of tools/record_excel_run.py with synthetic bundles.

The bundles hold what tools/Record-FRTB-ExcelRun.ps1 writes: session.log
with KEY=value lines, and the build and comparison logs. They are synthetic:
they test that host.json follows the logs, and are never Excel evidence.
"""

import json

import pytest

import check_excel_evidence as validator
import record_excel_run as recorder
from test_excel_evidence import compare_log, repo  # noqa: F401  (repo is a fixture)

WORKBOOK_DEV = "c" * 64
WORKBOOK_PROD = "d" * 64
REFERENCES = ("Visual Basic For Applications|Microsoft Excel 16.0 Object Library|OLE Automation|"
              "Microsoft Office 16.0 Object Library")


def session(sha, **changes):
    values = {
        "CANDIDATE": sha, "STARTED": "2026-11-02T10:00:00+01:00", "OPERATOR": "daniele, WORKSTATION",
        "EXCEL_VERSION": "16.0", "EXCEL_BUILD": "16.0.19127.20000", "OFFICE_BITNESS": "64-bit",
        "OS": "Windows 11 Pro 24H2 (build 26100.4061)", "OS_ARCHITECTURE": "AMD64",
        "DECIMAL_SEPARATOR": ",", "LIST_SEPARATOR": ";",
        "MACRO_POLICY": "Disable VBA macros with notification", "VBA_PROJECT_ACCESS": "On",
        "BUILD_DEVELOPMENT_EXIT": "0", "COMPILE_DEVELOPMENT": "PASS",
        "COMPILE_DEVELOPMENT_DETAIL": "Debug > Compile VBAProject: no error", "REFERENCES": REFERENCES,
        "BUILD_PRODUCTION_EXIT": "0", "COMPILE_PRODUCTION": "PASS", "FINISHED": "2026-11-02T10:20:00+01:00",
    }
    values.update(changes)
    return "".join(f"{key}={value}\n" for key, value in values.items() if value is not None)


@pytest.fixture
def run(tmp_path, repo):  # noqa: F811
    """A passing bundle as the recorder leaves it; tests change it, then record it."""
    sha = repo["sha"]
    template = next(item["sha256"] for item in validator.inventory(repo["path"], sha)
                    if item["path"] == validator.TEMPLATE)
    bundle = tmp_path / "excel-run-test"
    bundle.mkdir()
    logs = {
        "session.log": session(sha),
        "build-development.log": (f"CANDIDATE={sha}\nImported CORE_Config.bas\n"
                                  "Regression test: Result: PASS (52 items)\n"
                                  "Saved C:\\FRTB\\dist\\K-FRTB-SA-development.xlsm\n"),
        "build-production.log": f"CANDIDATE={sha}\nImported CORE_Config.bas\nSaved C:\\FRTB\\dist\\prod.xlsm\n",
        "round-trip-development.log": compare_log(sha, "development", template, WORKBOOK_DEV),
        "round-trip-production.log": compare_log(sha, "production", template, WORKBOOK_PROD),
    }

    def record(**session_changes):
        if session_changes:
            logs["session.log"] = session(sha, **session_changes)
        for name, text in logs.items():
            if text is None:
                (bundle / name).unlink(missing_ok=True)
            else:
                (bundle / name).write_text(text, encoding="utf-8")
        code = recorder.main([str(bundle), "--candidate", sha, "--root", str(repo["path"])])
        host = json.loads((bundle / "host.json").read_text(encoding="utf-8")) if code != 2 else None
        return code, host

    return {"bundle": bundle, "logs": logs, "record": record, "sha": sha}


def statuses(host):
    return {name: stage["status"] for name, stage in host["stages"].items()}


def test_a_complete_run_is_recorded_and_passes(run, capsys):
    code, host = run["record"]()
    assert code == 0, capsys.readouterr()
    assert set(statuses(host).values()) == {"PASS"}
    assert host["candidate_sha"] == run["sha"]
    assert host["environment"]["references"] == REFERENCES.split("|")
    assert host["environment"]["trust_changes"] is False
    assert host["stages"]["regression"]["items"] == {"total": 52, "failures": 0}
    assert host["stages"]["round_trip_development"]["workbook_sha256"] == WORKBOOK_DEV
    reports = run["bundle"].parent / "excel-run-test-validation"
    assert json.loads((reports / "excel-evidence.json").read_text(encoding="utf-8"))["outcomes"] == ["PASS"]
    assert "Validator outcome: PASS" in capsys.readouterr().out


def test_a_trust_change_is_recorded(run):
    code, host = run["record"](TRUST_CHANGES="VBA project access turned on for the run and off after it")
    assert code == 0
    assert host["environment"]["trust_changes"] is True
    assert host["environment"]["trust_changes_detail"].startswith("VBA project access")


def test_a_failed_development_build_leaves_its_stages_not_run(run):
    run["logs"]["build-development.log"] = f"CANDIDATE={run['sha']}\nERROR: Imported nothing\n"
    run["logs"]["round-trip-development.log"] = None
    code, host = run["record"](BUILD_DEVELOPMENT_EXIT="1", COMPILE_DEVELOPMENT=None)
    assert code == 1
    found = statuses(host)
    assert found["build_development"] == "FAIL"
    assert found["compile_development"] == found["regression"] == found["round_trip_development"] == "NOT_RUN"
    assert host["stages"]["regression"]["items"] is None
    assert found["build_production"] == found["round_trip_production"] == "PASS"


def test_a_failed_compile_blocks_the_regression(run):
    code, host = run["record"](COMPILE_DEVELOPMENT="FAIL",
                               COMPILE_DEVELOPMENT_DETAIL="Debug > Compile VBAProject reported a compile error")
    assert code == 1
    assert statuses(host)["compile_development"] == "FAIL"
    assert statuses(host)["regression"] == "NOT_RUN"


def test_a_failed_regression_is_recorded_with_its_counts(run):
    run["logs"]["build-development.log"] = run["logs"]["build-development.log"].replace(
        "Result: PASS (52 items)", "Result: FAIL (2 of 52 items)")
    # Build-Workbook.ps1 exits 1 after a failed regression, and saves the workbook
    code, host = run["record"](BUILD_DEVELOPMENT_EXIT="1")
    assert code == 1
    assert statuses(host)["build_development"] == "PASS"
    assert host["stages"]["regression"]["status"] == "FAIL"
    assert host["stages"]["regression"]["items"] == {"total": 52, "failures": 2}
    report = json.loads((run["bundle"].parent / "excel-run-test-validation/excel-evidence.json").read_text())
    assert report["outcomes"] == ["TEST_FAILED"]


def test_exit_code_1_without_a_failed_regression_is_a_failed_build(run):
    code, host = run["record"](BUILD_PRODUCTION_EXIT="1")
    assert code == 1
    assert statuses(host)["build_production"] == "FAIL"
    assert statuses(host)["compile_production"] == statuses(host)["round_trip_production"] == "NOT_RUN"


def test_a_failed_comparison_is_recorded(run):
    run["logs"]["round-trip-production.log"] = run["logs"]["round-trip-production.log"].replace(
        "RESULT=PASS; modules=11; differences=0", "RESULT=FAIL; modules=11; differences=1")
    code, host = run["record"]()
    assert code == 1
    assert statuses(host)["round_trip_production"] == "FAIL"
    assert "workbook_sha256" not in host["stages"]["round_trip_production"]
    assert host["stages"]["round_trip_production"]["detail"].endswith("RESULT=FAIL; modules=11; differences=1; binding=VERIFIED")


def test_an_export_that_failed_before_the_comparison_names_the_log(run):
    run["logs"]["round-trip-development.log"] = "ERROR: export-manifest.json not found\n"
    code, host = run["record"]()
    assert code == 1
    assert statuses(host)["round_trip_development"] == "FAIL"
    assert host["stages"]["round_trip_development"]["detail"].endswith("0 RESULT lines; see round-trip-development.log")


def test_a_session_log_for_another_candidate_is_an_error(run, capsys):
    code, _ = run["record"](CANDIDATE="0" * 40)
    assert code == 2
    assert "session.log names CANDIDATE" in capsys.readouterr().err


def test_a_missing_session_log_is_an_error(tmp_path, repo, capsys):  # noqa: F811
    code = recorder.main([str(tmp_path), "--candidate", repo["sha"], "--root", str(repo["path"])])
    assert code == 2
    assert "session.log" in capsys.readouterr().err


def test_a_later_session_line_replaces_an_earlier_one(tmp_path):
    (tmp_path / "session.log").write_text("COMPILE_PRODUCTION=FAIL\nnot a key line\nCOMPILE_PRODUCTION=PASS\n",
                                          encoding="utf-8")
    assert recorder.read_session(tmp_path) == {"COMPILE_PRODUCTION": "PASS"}
