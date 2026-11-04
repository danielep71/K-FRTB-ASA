# SPDX-License-Identifier: MIT
"""Each quality gate fails on a deliberate defect (#219).

The fixtures run Ruff, mypy and actionlint with the repository's
configuration on small files in a temporary folder. Without a tool they are
skipped, unless REQUIRE_QUALITY_TOOLS=1, as in CI, where a missing tool fails.
"""

import os
import subprocess
import sys
import tempfile
import unittest
import unittest.mock
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools"))
import check_repository as checker  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
CONFIG = ROOT / "pyproject.toml"
REQUIRED = os.environ.get("REQUIRE_QUALITY_TOOLS") == "1"

COMPLEX = "def branches(x: int) -> int:\n" + "".join(
    f"    if x == {n}:\n        return {n}\n" for n in range(16)) + "    return -1\n"
UNREACHABLE = (
    "def kind(x: int) -> str:\n"
    "    if isinstance(x, int):\n"
    "        return 'int'\n"
    "    return 'other'\n"
)
INVALID_WORKFLOW = (
    "name: Broken\n"
    "on: push\n"
    "permissions: {}\n"
    "jobs:\n"
    "  build:\n"
    "    runs-on: ubuntu-24.04\n"
    "    timeout-minutes: 5\n"
    "    steps:\n"
    "      - run: echo ${{ github.no_such_context }}\n"
)


def tool(name: str) -> list[str]:
    command = checker.quality_command(name)
    if command is None:
        if REQUIRED:
            raise AssertionError(f"{name} is not installed and REQUIRE_QUALITY_TOOLS=1")
        raise unittest.SkipTest(f"{name} is not installed")
    return command


def run(command: list[str], cwd: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(command, cwd=cwd, capture_output=True, text=True, check=False)


class QualityGateTests(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.folder = Path(directory.name)

    def ruff(self, source: str) -> subprocess.CompletedProcess[str]:
        (self.folder / "sample.py").write_text(source, encoding="utf-8")
        return run(tool("ruff") + ["check", "--config", str(CONFIG), "--no-cache", "sample.py"], self.folder)

    def mypy(self, source: str) -> subprocess.CompletedProcess[str]:
        (self.folder / "sample.py").write_text(source, encoding="utf-8")
        return run(tool("mypy") + ["--config-file", str(CONFIG), "--no-incremental", "sample.py"], self.folder)

    def test_ruff_passes_clean_code_and_fails_an_unused_import(self) -> None:
        self.assertEqual(self.ruff("def one() -> int:\n    return 1\n").returncode, 0)
        result = self.ruff("import os\n")
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("F401", result.stdout)

    def test_ruff_fails_a_complexity_breach(self) -> None:
        result = self.ruff(COMPLEX)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("C901", result.stdout)

    def test_mypy_passes_typed_code_and_fails_a_type_error(self) -> None:
        self.assertEqual(self.mypy("def one() -> int:\n    return 1\n").returncode, 0, "clean fixture must pass")
        result = self.mypy("def one() -> int:\n    return 'one'\n")
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("[return-value]", result.stdout)

    def test_mypy_fails_an_unreachable_branch(self) -> None:
        result = self.mypy(UNREACHABLE)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("[unreachable]", result.stdout)

    def test_actionlint_fails_an_invalid_workflow(self) -> None:
        workflows = self.folder / ".github" / "workflows"
        workflows.mkdir(parents=True)
        (workflows / "broken.yml").write_text(INVALID_WORKFLOW, encoding="utf-8")
        arguments = ["-no-color", "-oneline", "-shellcheck=", "-pyflakes=", ".github/workflows/broken.yml"]
        result = run(tool("actionlint") + arguments, self.folder)
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)  # 1: lint errors, not a failed run
        self.assertIn("no_such_context", result.stdout)
        (workflows / "broken.yml").write_text(INVALID_WORKFLOW.replace("github.no_such_context", "github.sha"),
                                              encoding="utf-8")
        self.assertEqual(run(tool("actionlint") + arguments, self.folder).returncode, 0)

    def test_a_missing_tool_fails_only_when_required(self) -> None:
        with unittest.mock.patch.object(checker, "quality_command", return_value=None):  # an empty environment
            commands, failed, skipped = checker.quality_plan(require=True)
            self.assertEqual((commands, skipped), ([], []))
            self.assertEqual(len(failed), len(checker.QUALITY_GATES))
            commands, failed, skipped = checker.quality_plan(require=False)
            self.assertEqual((commands, failed), ([], []))
            self.assertEqual(skipped, [name for name, _, _ in checker.QUALITY_GATES])

if __name__ == "__main__":
    unittest.main()
