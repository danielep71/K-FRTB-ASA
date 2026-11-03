"""Compare the VBA exported from a built workbook with the repository source.

Part of the Excel round trip (#30, #79). A workbook built from the template
and the source of one commit, then exported with tools/Export-FRTB-VBA.ps1,
must hold exactly the standard modules of its mode:

* production: every module of src/modules;
* development: those and every module of tests/modules.

Each module must carry the same code and attributes as the source of that
commit, which is read from Git, not from the working tree. The comparison
ignores line endings, trailing blank lines, and the letter case of code
outside string literals and comments: VBA names are case-insensitive, and
the VBE re-cases identifiers on import. The re-cased lines are counted.
Document modules (ThisWorkbook and the sheets) are not part of the source.
They must hold no code beyond Option lines. Everything else must match. The
VBE's Export leaves the VB_Base attribute out of a document module, so the
tool takes the component type the exporter records in export-manifest.json
(100, a document module) and accepts a VB_Base line as well.

Source parity says nothing about which workbook the files came from. The
exporter writes export-manifest.json with the SHA-256 of the workbook it
opened and of every file it exported. With --workbook the tool checks the
workbook and the files against it. It reports the binding as VERIFIED, as
FAIL, or as UNVERIFIED for an export without a manifest or without
--workbook. A manifest is a record, not a signature: it binds an export to a
workbook, but it does not attest who made either.

The tool never runs Excel. Usage:

    python3 tools/compare_vba_export.py --export <folder> --mode production|development
        [--candidate-sha <commit>] [--workbook <file.xlsm>]

Exit codes: 0 PASS, 1 FAIL, 2 the inputs could not be read.
"""

from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import re
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
MODE_FOLDERS = {
    "production": ("src/modules",),
    "development": ("src/modules", "tests/modules"),
}
EXPORT_MANIFEST = "export-manifest.json"
EXPORT_FORMAT = "k-frtb-sa-vba-export-1"
VBA_PROJECT_PART = "xl/vbaProject.bin"
TEMPLATE = "src/workbook/FRTB_Template.xlsx"
VBA_ENCODING = "cp1252"
SUFFIXES = (".bas", ".cls", ".frm")
# A form's binary properties, which the VBE exports beside its .frm: not
# compared with the source, but bound to the workbook like the components.
RESOURCE_SUFFIX = ".frx"
# vbext_ct_Document: the VBComponent type of ThisWorkbook and the sheets
DOCUMENT_TYPE = 100
VB_NAME = re.compile(r'^Attribute VB_Name = "([^"]+)"$')
# The Excel Workbook, Worksheet and Chart classes a document module names in
# its VB_Base attribute; an ordinary class module has no VB_Base. The VBE's
# Export omits the line, so an export is read with its manifest's types.
DOCUMENT_BASE = re.compile(r'^Attribute VB_Base = "0\{0002082[01]-|^Attribute VB_Base = "0\{00020819-')
REM = re.compile(r"rem(?=\s|$)", re.IGNORECASE)
# Lines the VBE writes before the code of a class or form: the VERSION line,
# the BEGIN ... END property block and the Attribute lines.
CLASS_HEADER = re.compile(r"^(VERSION \d+\.\d+( CLASS)?|(?i:begin)(\s.*)?|(?i:end)|\s+\S.*)$")
OPTION_LINE = re.compile(r"^\s*Option\s+\w+(\s+\w+)?\s*$", re.IGNORECASE)


@dataclass(frozen=True)
class Component:
    name: str
    attributes: tuple[str, ...]
    body: tuple[str, ...]

    @property
    def document(self) -> bool:
        return any(DOCUMENT_BASE.match(line) for line in self.attributes)


def parse(raw: bytes, label: str) -> Component:
    """Split a VBA file into its VB_Name, its Attribute lines and its code."""
    lines = raw.decode(VBA_ENCODING).replace("\r\n", "\n").replace("\r", "\n").split("\n")
    index = 0
    attributes: list[str] = []
    while index < len(lines):
        line = lines[index]
        if line.startswith("Attribute "):
            attributes.append(line)
        elif not CLASS_HEADER.match(line) or attributes:
            break
        index += 1
    body = lines[index:]
    while body and body[-1] == "":
        body.pop()
    names = [match.group(1) for match in map(VB_NAME.match, attributes) if match]
    if len(names) != 1:
        raise ValueError(f"{label}: expected one Attribute VB_Name line, found {len(names)}")
    return Component(names[0], tuple(attributes), tuple(body))


def fold_case(line: str) -> str:
    """``line`` with its code lower-cased, string literals and the comment kept.

    A comment starts at an apostrophe outside a string, or at ``Rem`` where a
    statement starts: the start of the line, after a line number or after a
    colon."""
    folded: list[str] = []
    in_string = False
    statement_start = True
    index = 0
    while index < len(line):
        char = line[index]
        if in_string:
            in_string = char != '"'
        elif char == '"':
            in_string, statement_start = True, False
        elif char == "'":
            return "".join(folded) + line[index:]
        elif char == ":":
            statement_start = True
        elif statement_start and REM.match(line, index):
            return "".join(folded) + "rem" + line[index + 3:]
        elif statement_start and char.isdigit() and not "".join(folded).strip(" \t0123456789"):
            pass
        elif not char.isspace():
            statement_start = False
        folded.append(char if in_string or char == '"' else char.lower())
        index += 1
    return "".join(folded)


def git(root: Path, *args: str) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(["git", "-C", str(root), *args], capture_output=True, check=False)


def committed(root: Path, sha: str, path: str) -> bytes:
    completed = git(root, "show", f"{sha}:{path}")
    if completed.returncode != 0:
        raise ValueError(f"{path}: not found at {sha}")
    return completed.stdout


def expected_components(root: Path, sha: str, mode: str) -> dict[str, tuple[str, Component]]:
    """The standard modules of ``mode`` at ``sha``, by lower-case VB_Name."""
    expected: dict[str, tuple[str, Component]] = {}
    for folder in MODE_FOLDERS[mode]:
        listing = git(root, "ls-tree", "--name-only", f"{sha}:{folder}")
        if listing.returncode != 0:
            raise ValueError(f"{folder}: not found at {sha}")
        for name in sorted(listing.stdout.decode("utf-8").split()):
            if not name.lower().endswith(".bas"):
                continue
            path = f"{folder}/{name}"
            component = parse(committed(root, sha, path), path)
            key = component.name.lower()
            if key in expected:
                raise ValueError(f"{path}: VB_Name {component.name} also in {expected[key][0]}")
            expected[key] = (path, component)
    if not expected:
        raise ValueError(f"no modules found in {', '.join(MODE_FOLDERS[mode])} at {sha}")
    return expected


def exported_components(folder: Path) -> dict[str, tuple[str, Component]]:
    exported: dict[str, tuple[str, Component]] = {}
    for path in sorted(folder.iterdir()):
        if path.suffix.lower() not in SUFFIXES:
            continue
        if not path.is_file():
            raise ValueError(f"{path.name}: not a file")
        component = parse(path.read_bytes(), path.name)
        key = component.name.lower()
        if key in exported:
            raise ValueError(f"{path.name}: VB_Name {component.name} exported twice")
        exported[key] = (path.name, component)
    return exported


def first_difference(source: tuple[str, ...], export: tuple[str, ...]) -> str:
    for number, (left, right) in enumerate(zip(source, export), start=1):
        if fold_case(left) != fold_case(right):
            return f"line {number}: source {left!r}, export {right!r}"
    shorter = min(len(source), len(export))
    return f"line {shorter + 1}: source has {len(source)} lines, export {len(export)}"


def compare(expected: dict[str, tuple[str, Component]],
            exported: dict[str, tuple[str, Component]],
            documents: frozenset[str] = frozenset()) -> tuple[list[str], int]:
    """The differences between the source and the export, empty when they
    match, and the number of lines that differ only in letter case.

    ``documents`` holds the lower-case file names the manifest records as
    document modules. Such a module must still hold no code, so a wrong type
    in the manifest cannot hide code from the comparison."""
    findings = [f"missing from the export: {expected[key][1].name} ({expected[key][0]})"
                for key in sorted(set(expected) - set(exported))]
    for key in sorted(set(exported) - set(expected)):
        file, component = exported[key]
        if not (component.document or file.lower() in documents):
            findings.append(f"not in the source for this mode: {component.name} ({file})")
        elif any(line.strip() and not OPTION_LINE.match(line) for line in component.body):
            findings.append(f"document module {component.name} ({file}) holds code that is not in the source")
    recased = 0
    for key in sorted(set(expected) & set(exported)):
        path, source = expected[key]
        _, export = exported[key]
        if source.name != export.name:
            findings.append(f"{path}: VB_Name {source.name} exported as {export.name}")
        if source.attributes != export.attributes:
            changed = [line for line in difflib.unified_diff(
                source.attributes, export.attributes, lineterm="", n=0)
                if line[:1] in "+-" and line[:3] not in ("+++", "---")]
            findings.append(f"{path}: attributes differ: {'; '.join(changed)}")
        if [fold_case(line) for line in source.body] != [fold_case(line) for line in export.body]:
            findings.append(f"{path}: code differs at {first_difference(source.body, export.body)}")
        else:
            recased += sum(left != right for left, right in zip(source.body, export.body))
    return findings, recased


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


@dataclass(frozen=True)
class ExportRecord:
    workbook_name: str
    workbook_sha256: str
    files: dict[str, tuple[str, str, str]]  # lower-case file -> (file, VB_Name, sha256)
    documents: frozenset[str] = frozenset()  # lower-case files of type DOCUMENT_TYPE


def load_export_manifest(path: Path) -> ExportRecord:
    """export-manifest.json; ValueError when it is not a valid manifest."""
    try:
        record = json.loads(path.read_bytes().decode("utf-8-sig"))
        if not isinstance(record, dict) or record.get("format") != EXPORT_FORMAT:
            raise ValueError(f"format is not {EXPORT_FORMAT}")
        name = record["workbook"]["file"]
        if not isinstance(name, str) or not name:
            raise ValueError("workbook.file is not a file name")
        files = {item["file"].lower(): (item["file"], item["name"], item["sha256"].lower())
                 for item in record["components"]}
        if len(files) != len(record["components"]):
            raise ValueError("a file is listed twice")
        types = {item["file"].lower(): item.get("type") for item in record["components"]}
        if any(value is not None and (isinstance(value, bool) or not isinstance(value, int))
               for value in types.values()):
            raise ValueError("a component type is not an integer")
        documents = frozenset(file for file, value in types.items() if value == DOCUMENT_TYPE)
        return ExportRecord(name, record["workbook"]["sha256"].lower(), files, documents)
    except (ValueError, KeyError, TypeError, AttributeError) as error:
        raise ValueError(f"{EXPORT_MANIFEST} is not a valid export manifest: {error}") from error


def file_findings(folder: Path, record: ExportRecord) -> list[str]:
    """The exported files that are missing, unlisted, changed or renamed
    against the manifest."""
    listed = record.files
    present = {item.name.lower(): item for item in folder.iterdir()
               if item.is_file() and item.suffix.lower() in (*SUFFIXES, RESOURCE_SUFFIX)}
    findings = [f"{EXPORT_MANIFEST} lists {listed[key][0]}, which is not in the export"
                for key in sorted(set(listed) - set(present))]
    findings += [f"{present[key].name} is in the export but not in {EXPORT_MANIFEST}"
                 for key in sorted(set(present) - set(listed))]
    for key in sorted(set(listed) & set(present)):
        name, component, digest = listed[key]
        data = present[key].read_bytes()
        if sha256(data) != digest:
            findings.append(f"{name} changed after the export: its SHA-256 is not the one in {EXPORT_MANIFEST}")
        elif present[key].suffix.lower() != RESOURCE_SUFFIX:
            try:
                exported_name = parse(data, name).name
            except ValueError:
                continue  # reported by the source comparison
            if exported_name != component:
                findings.append(f"{name} declares VB_Name {exported_name}, {EXPORT_MANIFEST} names {component}")
    return findings


def is_macro_workbook(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path) as archive:
            return VBA_PROJECT_PART in archive.namelist()
    except (zipfile.BadZipFile, OSError):
        return False


def check_binding(folder: Path, workbook: Path | None) -> tuple[str, list[str], list[str]]:
    """The binding of the export to the workbook, VERIFIED, UNVERIFIED or
    FAIL, the findings that make it FAIL and notes that do not.

    A manifest that does not match the exported files fails even without
    --workbook; without either, or without a manifest, the binding is
    UNVERIFIED. The workbook is bound by its bytes: a workbook renamed after
    the export, as a download often is, gets a note, not a failure."""
    path = folder / EXPORT_MANIFEST
    if not path.is_file():
        return "UNVERIFIED", [], []
    try:
        record = load_export_manifest(path)
    except ValueError as error:
        return "FAIL", [str(error)], []
    findings = file_findings(folder, record)
    notes: list[str] = []
    if workbook is not None:
        data = workbook.read_bytes()
        same_bytes = sha256(data) == record.workbook_sha256
        if not same_bytes:
            findings.append(f"{workbook.name} is not the workbook the export was made from: "
                            f"sha256 {sha256(data)}, {EXPORT_MANIFEST} records {record.workbook_sha256}")
        if not is_macro_workbook(workbook):
            findings.append(f"{workbook.name} is not a macro-enabled workbook ({VBA_PROJECT_PART} not found)")
        if same_bytes and workbook.name.lower() != record.workbook_name.lower():
            notes.append(f"{EXPORT_MANIFEST} names the workbook {record.workbook_name}; "
                         "it was renamed since the export")
    if findings:
        return "FAIL", findings, notes
    return ("VERIFIED" if workbook is not None else "UNVERIFIED"), [], notes


def manifest_documents(folder: Path) -> frozenset[str]:
    """The files the manifest records as document modules; none without a
    valid manifest, whose absence or fault the binding reports."""
    path = folder / EXPORT_MANIFEST
    if not path.is_file():
        return frozenset()
    try:
        return load_export_manifest(path).documents
    except ValueError:
        return frozenset()


def binding_note(binding: str, folder: Path, workbook: Path | None) -> str:
    if binding == "VERIFIED":
        return f"{EXPORT_MANIFEST} records this workbook and every exported file"
    if binding == "FAIL":
        return f"the export, {EXPORT_MANIFEST} and the workbook do not agree"
    if not (folder / EXPORT_MANIFEST).is_file():
        return f"no {EXPORT_MANIFEST}: the export is not bound to a workbook"
    return "no --workbook given: the export is not checked against a workbook"


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=REPO_ROOT)
    parser.add_argument("--export", type=Path, required=True, help="folder written by Export-FRTB-VBA.ps1")
    parser.add_argument("--mode", choices=tuple(MODE_FOLDERS), required=True)
    parser.add_argument("--candidate-sha", default="HEAD", help="commit the workbook was built from")
    parser.add_argument("--workbook", type=Path,
                        help="built .xlsm the export was made from, checked against export-manifest.json")
    options = parser.parse_args(arguments)
    try:
        resolved = git(options.root, "rev-parse", "--verify", f"{options.candidate_sha}^{{commit}}")
        if resolved.returncode != 0:
            raise ValueError(f"unknown commit: {options.candidate_sha}")
        sha = resolved.stdout.decode().strip()
        if not options.export.is_dir():
            raise ValueError(f"export folder not found: {options.export}")
        if options.workbook is not None and not options.workbook.is_file():
            raise ValueError(f"workbook not found: {options.workbook}")
        expected = expected_components(options.root, sha, options.mode)
        exported = exported_components(options.export)
        findings, recased = compare(expected, exported, manifest_documents(options.export))
        binding, binding_findings, binding_notes = check_binding(options.export, options.workbook)
        template = sha256(committed(options.root, sha, TEMPLATE))
        workbook_digest = sha256(options.workbook.read_bytes()) if options.workbook is not None else ""
    except (ValueError, OSError, UnicodeDecodeError) as error:
        # An input that cannot be read is not a comparison result: exit 2.
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    print(f"Candidate: {sha}")
    print(f"Mode:      {options.mode} ({len(expected)} modules expected, {len(exported)} components exported)")
    print(f"Template:  {TEMPLATE} sha256 {template}")
    if options.workbook is not None:
        print(f"Workbook:  {options.workbook.name} sha256 {workbook_digest}")
    print(f"Re-cased:  {recased} line(s) differ only in the letter case of code (ignored)")
    print(f"Binding:   {binding} - {binding_note(binding, options.export, options.workbook)}")
    for note in binding_notes:
        print(f"  NOTE {note}")
    for finding in findings + binding_findings:
        print(f"  FAIL {finding}")
    passed = not findings and binding != "FAIL"
    print(f"RESULT={'PASS' if passed else 'FAIL'}; modules={len(expected)}; "
          f"differences={len(findings)}; binding={binding}")
    return 0 if passed else 1


if __name__ == "__main__":
    sys.exit(main())
