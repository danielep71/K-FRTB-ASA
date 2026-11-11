"""Run the workbook's VBA modules headless in LibreOffice Calc.

LibreOffice executes VBA source in compatibility mode (Option VBASupport 1),
which lets the engine run on Linux without Excel. The runner:

1. starts a private headless soffice instance,
2. opens an empty spreadsheet (or --input, for example the template) and
   inserts every module from src/modules and tests/modules,
3. calls the requested macros in order,
4. saves the result as .xlsx (no macros) for inspection.

This is a cross-check of the VBA logic, not a substitute for running the
workbook in Excel: LibreOffice's VBA support differs from Excel in places.

Usage:
    python3 tools/lo_runner.py OUTPUT.xlsx [--macro NAME ...] [--input BOOK.xlsx]

Without --macro the default sequence builds the workbook, runs the
calculation and runs the regression test.
"""

from __future__ import annotations

import argparse
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import time
from collections.abc import Callable, Mapping
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[1]
VBA_DIRS = [REPO_ROOT / "src" / "modules", REPO_ROOT / "tests" / "modules"]
TEMPLATE = REPO_ROOT / "src" / "workbook" / "FRTB_Template.xlsx"
DEFAULT_MACROS = (
    "M_Setup.BuildWorkbook",
    "M_Main.RunFRTB_Silent",
    "TEST_Regression.Test_RunRegression_Silent",
)


def module_paths() -> list[Path]:
    return sorted((p for d in VBA_DIRS for p in d.glob("*.bas")), key=lambda p: p.stem)


def probe_name(module: str) -> str:
    return f"LoCompileProbe_{module}"


def module_source(path: Path) -> str:
    """Return module text prepared for LibreOffice.

    Attribute lines are dropped and a probe function is appended. LibreOffice
    does not report compile errors through UNO: a module that fails to
    compile silently does nothing. Calling the probe proves the module
    compiled.
    """
    lines = path.read_text(encoding="cp1252").splitlines()
    # LibreOffice VBA compatibility omits this VBA intrinsic constant.
    body = [ln.replace("vbObjectError", "LoVbObjectError")
            for ln in lines if not ln.startswith("Attribute ")]
    declaration = "Private Const LoVbObjectError As Long = -2147221504"
    explicit = next(i for i, line in enumerate(body) if line.strip() == "Option Explicit")
    body.insert(explicit + 1, declaration)
    probe = [
        "",
        f"Public Function {probe_name(path.stem)}() As Long",
        f"    {probe_name(path.stem)} = 1",
        "End Function",
    ]
    return "Option VBASupport 1\n" + "\n".join(body + probe) + "\n"


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        port: int = s.getsockname()[1]
        return port


def run(  # noqa: C901 -- one LibreOffice session: start, load, run, save, shut down
    output: Path,
    macros: list[str],
    input_book: Path | None = None,
    post_process: Callable[[Any], object] | None = None,
    expected_results: Mapping[str, object] | None = None,
) -> list[str]:
    """Run macros and save the document as .xlsx.

    post_process, when given, is called with the UNO document after the
    macros and before saving (used to strip controls from the template).
    """
    import uno  # noqa: PLC0415 - only available where LibreOffice is installed
    from com.sun.star.beans import PropertyValue  # noqa: PLC0415

    def prop(name: str, value: object) -> Any:
        p = PropertyValue()
        p.Name = name
        p.Value = value
        return p

    port = free_port()
    profile = Path(tempfile.mkdtemp(prefix="lo-profile-"))
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("LibreOffice (soffice) not found on PATH")
    proc = subprocess.Popen(
        [
            soffice,
            "--headless",
            "--invisible",
            "--norestore",
            "--nologo",
            "--nodefault",
            f"-env:UserInstallation={profile.as_uri()}",
            f"--accept=socket,host=127.0.0.1,port={port};urp;",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    errors: list[str] = []
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local
        )
        ctx = None
        for _ in range(120):
            try:
                ctx = resolver.resolve(
                    f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
                )
                break
            except Exception:  # noqa: BLE001 - soffice still starting
                time.sleep(0.5)
        if ctx is None:
            raise RuntimeError("Could not connect to LibreOffice")
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx
        )
        url = input_book.resolve().as_uri() if input_book else "private:factory/scalc"
        doc = desktop.loadComponentFromURL(url, "_blank", 0, (prop("Hidden", True),))

        libs = doc.BasicLibraries
        lib = libs.getByName("Standard") if libs.hasByName("Standard") else libs.createLibrary("Standard")
        for path in module_paths():
            name = path.stem
            if lib.hasByName(name):
                lib.removeByName(name)
            lib.insertByName(name, module_source(path))

        provider = doc.getScriptProvider()
        for path in module_paths():
            script = provider.getScript(
                f"vnd.sun.star.script:Standard.{path.stem}.{probe_name(path.stem)}"
                "?language=Basic&location=document"
            )
            try:
                result = script.invoke((), (), ())[0]
            except Exception as exc:  # noqa: BLE001
                result = exc
            if result != 1:
                errors.append(f"{path.stem}: module did not compile ({result!r})")
        for macro in macros:
            script = provider.getScript(
                f"vnd.sun.star.script:Standard.{macro}?language=Basic&location=document"
            )
            try:
                result = script.invoke((), (), ())[0]
                if expected_results is not None and macro in expected_results:
                    if result != expected_results[macro]:
                        errors.append(f"{macro}: unexpected completion result {result!r}")
            except Exception as exc:  # noqa: BLE001 - report and continue
                errors.append(f"{macro}: {exc}")

        if post_process is not None:
            post_process(doc)
        output.parent.mkdir(parents=True, exist_ok=True)
        doc.storeToURL(output.resolve().as_uri(), (prop("FilterName", "Calc MS Excel 2007 XML"),))
        doc.close(True)
        try:
            desktop.terminate()
        except Exception:  # noqa: BLE001 - connection drops on terminate
            pass
        proc.wait(timeout=30)
    finally:
        if proc.poll() is None:
            proc.kill()
        shutil.rmtree(profile, ignore_errors=True)
    return errors


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("output", type=Path)
    parser.add_argument("--macro", action="append", dest="macros")
    parser.add_argument("--input", type=Path, default=None)
    args = parser.parse_args(argv)
    errors = run(args.output, args.macros or list(DEFAULT_MACROS), args.input)
    for e in errors:
        print("MACRO ERROR:", e, file=sys.stderr)
    return 1 if errors else 0


if __name__ == "__main__":
    os.environ.setdefault("SAL_USE_VCLPLUGIN", "svp")
    sys.exit(main(sys.argv[1:]))
