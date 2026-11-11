# SPDX-License-Identifier: MIT
"""The LibreOffice install step of the prototype checks falls back to
archive.ubuntu.com when the runner's Azure mirror does not answer (#263).

The step's script runs with stand-ins for sudo and apt-get, and with the apt
configuration paths moved into a temporary folder, so nothing is installed.
"""

import os
import shutil
import stat
import subprocess
import tempfile
import unittest
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parents[2]
WORKFLOW = ROOT / ".github" / "workflows" / "prototype-checks.yml"
STEP = "Install LibreOffice Calc and the Python UNO bridge"
AZURE = "http://azure.archive.ubuntu.com/ubuntu/\tpriority:1\nhttps://archive.ubuntu.com/ubuntu/\tpriority:2\n"


def step_script() -> str:
    workflow = yaml.safe_load(WORKFLOW.read_text(encoding="utf-8"))
    steps = workflow["jobs"]["checks"]["steps"]
    return str(next(step["run"] for step in steps if step.get("name") == STEP))


@unittest.skipUnless(shutil.which("bash"), "bash is not installed")
class MirrorFallbackTests(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.apt = self.root / "etc" / "apt"
        (self.apt / "sources.list.d").mkdir(parents=True)
        (self.apt / "sources.list.d" / "ubuntu.sources").write_text(
            "Types: deb\nURIs: mirror+file:/etc/apt/apt-mirrors.txt\n", encoding="utf-8")
        self.calls = self.root / "calls.log"
        bin_dir = self.root / "bin"
        bin_dir.mkdir()
        # apt-get fails while the mirror list still names the Azure mirror
        self.stub(bin_dir / "apt-get", f'echo "apt-get $*" >> "{self.calls}"\n'
                                       f'grep -q azure "{self.apt / "apt-mirrors.txt"}" && exit 100\nexit 0\n')
        self.stub(bin_dir / "sudo", 'exec "$@"\n')
        self.env = {**os.environ, "PATH": f"{bin_dir}{os.pathsep}{os.environ['PATH']}"}

    def stub(self, path: Path, body: str) -> None:
        path.write_text("#!/bin/bash\n" + body, encoding="utf-8")
        path.chmod(path.stat().st_mode | stat.S_IXUSR)

    def run_step(self, mirrors: str) -> subprocess.CompletedProcess[str]:
        (self.apt / "apt-mirrors.txt").write_text(mirrors, encoding="utf-8")
        script = step_script().replace("/etc/apt", str(self.apt))
        return subprocess.run(["bash", "-e", "-c", script], env=self.env, capture_output=True, text=True,
                              timeout=60, check=False)

    def test_a_working_mirror_installs_on_the_first_attempt(self) -> None:
        result = self.run_step("http://mirror.example/ubuntu/\tpriority:1\n")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        calls = self.calls.read_text(encoding="utf-8").splitlines()
        self.assertEqual(len(calls), 2)
        self.assertTrue(calls[0].endswith(" update") and " install " in calls[1])
        self.assertNotIn("::warning::", result.stdout)

    def test_a_silent_azure_mirror_falls_back_to_archive_ubuntu_com(self) -> None:
        result = self.run_step(AZURE)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("::warning::The runner's package mirror did not answer", result.stdout)
        self.assertNotIn("azure", (self.apt / "apt-mirrors.txt").read_text(encoding="utf-8"))
        calls = self.calls.read_text(encoding="utf-8").splitlines()
        self.assertEqual(sum(" install " in call for call in calls), 1)  # only the fallback got that far
        self.assertTrue(all("Acquire::http::Timeout=" in call for call in calls))

    def test_the_first_attempt_has_its_own_limits_within_the_step_timeout(self) -> None:
        workflow = yaml.safe_load(WORKFLOW.read_text(encoding="utf-8"))
        job = workflow["jobs"]["checks"]
        step = next(step for step in job["steps"] if step.get("name") == STEP)
        self.assertIn("timeout 120 apt-get", step["run"])
        self.assertIn("timeout 300 apt-get", step["run"])
        self.assertLess(7, step["timeout-minutes"])  # the bounded first attempt fits, with room for the fallback
        self.assertLess(step["timeout-minutes"], job["timeout-minutes"])


if __name__ == "__main__":
    unittest.main()
