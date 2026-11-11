# SPDX-License-Identifier: MIT
"""tests/requirements.txt pins the Windows-only dependencies of pytest (#246).

pip-compile on Linux drops a requirement whose marker excludes Linux, and a
hashed install on Windows then fails for want of its pin and hash. The check
reads the dependencies the installed pytest declares; without pytest it is
skipped, unless REQUIRE_QUALITY_TOOLS=1, as in CI, where that fails.
"""

import importlib.metadata
import os
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
REQUIRED = os.environ.get("REQUIRE_QUALITY_TOOLS") == "1"
WINDOWS = {"sys_platform": "win32", "platform_system": "Windows", "os_name": "nt"}
LINUX = {"sys_platform": "linux", "platform_system": "Linux", "os_name": "posix"}
# name==version, then an optional "; marker" before the hashes.
PIN = re.compile(r"(?m)^([A-Za-z0-9_.-]+)==([^\s;\\]+)[ \t]*(?:;([^\\\n]*))?")


def canonical(name: str) -> str:
    return re.sub(r"[-_.]+", "-", name).lower()


class WindowsRequirementTests(unittest.TestCase):
    def test_windows_only_pytest_dependencies_are_pinned(self) -> None:
        try:
            from packaging.markers import Marker
            from packaging.requirements import Requirement
            declared = importlib.metadata.requires("pytest") or []
        except (ImportError, importlib.metadata.PackageNotFoundError):
            if REQUIRED:
                raise
            self.skipTest("pytest is not installed")
        text = (ROOT / "tests/requirements.txt").read_text(encoding="utf-8")
        pins = {canonical(name): (version, marker.strip()) for name, version, marker in PIN.findall(text)}
        for requirement in map(Requirement, declared):
            marker = requirement.marker
            if marker is None or not marker.evaluate(WINDOWS) or marker.evaluate(LINUX):
                continue
            with self.subTest(requirement=str(requirement)):
                name = canonical(requirement.name)
                self.assertIn(name, pins, f"tests/requirements.txt lacks the Windows-only {requirement}")
                version, pinned_marker = pins[name]
                self.assertIn(version, requirement.specifier)
                self.assertTrue(not pinned_marker or Marker(pinned_marker).evaluate(WINDOWS))


if __name__ == "__main__":
    unittest.main()
