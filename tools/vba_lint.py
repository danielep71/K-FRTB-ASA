"""Static checks for the VBA modules in src/modules and tests/modules.

LibreOffice compiles VBA leniently: undeclared variables and calls to
missing procedures only fail when the line runs. Excel's compiler rejects
them up front under Option Explicit. This checker approximates that part of
Excel's "Debug > Compile" for the subset of VBA used in this project:

* every module declares Option Explicit and an Attribute VB_Name matching
  its file name; modules in src/modules are named CORE_* (engine, with
  Option Private Module) or M_* (workbook macros); modules in tests/modules
  are named TEST_*;
* every identifier is a keyword, a built-in, or declared in scope
  (procedure parameters and locals, module-level names, public names of any
  module);
* no duplicate declarations in one scope;
* private procedures are not called from other modules;
* files are ASCII, so the Windows-1252 import in Excel is lossless.

It is a heuristic, not a VBA parser. Usage: python3 tools/vba_lint.py
"""

from __future__ import annotations

import re
import sys
from collections.abc import Iterator
from dataclasses import dataclass, field
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
VBA_DIRS = [REPO_ROOT / "src" / "modules", REPO_ROOT / "tests" / "modules"]

KEYWORDS = {
    "and", "as", "attribute", "boolean", "byref", "byval", "call", "case", "collection", "const",
    "date", "dim", "do", "double", "each", "else", "elseif", "empty", "end", "error", "exit",
    "explicit", "false", "for", "function", "goto", "if", "in", "integer", "is", "let", "like",
    "long", "loop", "me", "mod", "new", "next", "not", "nothing", "null", "object", "on", "option",
    "optional", "or", "preserve", "private", "public", "redim", "resume", "select", "set",
    "single", "step", "string", "sub", "then", "to", "true", "until", "variant", "while", "with",
    "xor", "byte", "vb_name", "base", "compare", "text", "binary",
}

BUILTINS = {
    "abs", "application", "array", "cbool", "cdbl", "chr", "clng", "cstr", "debug", "dir", "fix", "int",
    "err", "exp", "format", "iif", "isarray", "isempty", "iserror", "isnumeric", "isobject", "isnull", "cverr",
    "lbound", "lcase", "lcase$", "left", "left$", "len", "log", "mid", "mid$", "msgbox", "now",
    "replace", "rgb", "right", "right$", "round", "sgn", "split", "sqr", "thisworkbook", "trim",
    "trim$", "typename", "ubound", "ucase", "ucase$", "val", "vartype", "join", "instr",
    "vbboolean", "vbcritical", "vbcrlf", "vbdate", "vbdirectory", "vbexclamation", "vbinformation",
    "vbobjecterror", "vbstring", "vbnullstring", "vbyesno", "vbyes", "vbno",
}

DECL_RE = re.compile(r"^\s*(?:(Public|Private|Dim|Static)\s+)?(?:(Const)\s+)?(.*)$", re.I)
PROC_RE = re.compile(
    r"^\s*(?:(Public|Private)\s+)?(Sub|Function)\s+([A-Za-z_]\w*)\s*\((.*)\)\s*(?:As\s+\w+)?\s*$", re.I
)
END_PROC_RE = re.compile(r"^\s*End\s+(Sub|Function)\b", re.I)
IDENT_RE = re.compile(r"(?<![\w.])([A-Za-z_]\w*\$?)")
LABEL_RE = re.compile(r"^\s*([A-Za-z_]\w*):\s*$")


@dataclass
class Proc:
    name: str
    public: bool
    params: set[str] = field(default_factory=set)
    body: list[tuple[int, str]] = field(default_factory=list)  # (lineno, text)


@dataclass
class Module:
    name: str
    path: Path
    procs: dict[str, Proc] = field(default_factory=dict)
    module_names: set[str] = field(default_factory=set)
    public_names: set[str] = field(default_factory=set)
    module_lines: list[tuple[int, str]] = field(default_factory=list)


def strip_code(line: str) -> str:
    """Remove string literals and comments."""
    out, in_str, i = [], False, 0
    while i < len(line):
        ch = line[i]
        if in_str:
            if ch == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    i += 2
                    continue
                in_str = False
            i += 1
            continue
        if ch == '"':
            in_str = True
            out.append('""')
        elif ch == "'":
            break
        else:
            out.append(ch)
        i += 1
    return "".join(out)


def logical_lines(text: str) -> Iterator[tuple[int, str]]:
    """Yield (first line number, code) with continuations joined."""
    buf, start = "", None
    for n, raw in enumerate(text.splitlines(), start=1):
        code = strip_code(raw).rstrip()
        if start is None:
            start = n
        if code.endswith(" _"):
            buf += code[:-2] + " "
            continue
        yield start, buf + code
        buf, start = "", None


def declared_names(decl_body: str) -> list[str]:
    """Names declared by 'a As Long, b() As String, c' (Dim/Const bodies)."""
    names, depth, cur = [], 0, ""
    for ch in decl_body:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            names.append(cur)
            cur = ""
        else:
            cur += ch
    names.append(cur)
    out = []
    for part in names:
        m = re.match(r"\s*(?:ByVal\s+|ByRef\s+|Optional\s+)*([A-Za-z_]\w*\$?)", part, re.I)
        if m:
            out.append(m.group(1))
    return out


def parse(  # noqa: C901 -- one pass over the lines; a branch per statement kind
    path: Path,
) -> tuple[Module, list[str]]:
    errors = []
    raw = path.read_bytes()
    if any(b > 127 for b in raw):
        errors.append(f"{path.name}: non-ASCII characters")
    text = raw.decode("cp1252")
    mod = Module(path.stem, path)
    m = re.match(r'Attribute VB_Name = "([^"]+)"', text)
    if not m or m.group(1) != path.stem:
        errors.append(f"{path.name}: first line must be Attribute VB_Name = \"{path.stem}\"")
    if not re.search(r"^Option Explicit\s*$", text, re.M):
        errors.append(f"{path.name}: missing Option Explicit")
    in_tests = path.parent.parent.name == "tests"
    if in_tests and not path.stem.startswith("TEST_"):
        errors.append(f"{path.name}: modules in tests/modules are named TEST_*")
    if not in_tests and not path.stem.startswith(("CORE_", "M_")):
        errors.append(f"{path.name}: modules in src/modules are named CORE_* or M_*")
    if path.stem.startswith("CORE_") and not re.search(r"^Option Private Module\s*$", text, re.M):
        errors.append(f"{path.name}: CORE_ modules declare Option Private Module")
    if not re.search(r"^' MODULE: " + re.escape(path.stem) + r"\s*$", text, re.M):
        errors.append(f"{path.name}: missing module banner (' MODULE: {path.stem})")

    current = None
    for n, code in logical_lines(text):
        if not code.strip():
            continue
        pm = PROC_RE.match(code)
        if pm:
            vis, _kind, name, params = pm.groups()
            current = Proc(name, (vis or "Public").lower() == "public")
            for p in declared_names(params) if params.strip() else []:
                current.params.add(p.lower())
            if name.lower() in mod.procs:
                errors.append(f"{path.name}:{n}: duplicate procedure {name}")
            mod.procs[name.lower()] = current
            mod.module_names.add(name.lower())
            if current.public:
                mod.public_names.add(name.lower())
            continue
        if END_PROC_RE.match(code):
            current = None
            continue
        if current is not None:
            current.body.append((n, code))
            continue
        mod.module_lines.append((n, code))
        dm = re.match(r"^\s*(Public|Private|Dim)\s+(Const\s+)?(.*)$", code, re.I)
        if dm:
            body = dm.group(3)
            if dm.group(2):
                body = body.split("=")[0]
            for name in declared_names(body):
                low = name.lower()
                if low in mod.module_names:
                    errors.append(f"{path.name}:{n}: duplicate module-level declaration {name}")
                mod.module_names.add(low)
                if dm.group(1).lower() == "public":
                    mod.public_names.add(low)
    return mod, errors


def check_proc(mod: Module, proc: Proc, all_public: dict[str, str], private_procs: dict[str, str]) -> list[str]:
    errors = []
    locals_ = set(proc.params)
    labels = set()
    for n, code in proc.body:
        lm = LABEL_RE.match(code)
        if lm:
            labels.add(lm.group(1).lower())
            continue
        for stmt in code.split(":"):
            dm = re.match(r"^\s*(Dim|Const|Static)\s+(.*)$", stmt, re.I)
            if dm:
                body = dm.group(2)
                if dm.group(1).lower() == "const":
                    body = body.split("=")[0]
                for name in declared_names(body):
                    low = name.lower()
                    if low in locals_:
                        errors.append(f"{mod.path.name}:{n}: duplicate declaration {name} in {proc.name}")
                    locals_.add(low)

    known = KEYWORDS | BUILTINS | locals_ | labels | mod.module_names | set(all_public)
    for n, code in proc.body:
        if LABEL_RE.match(code):
            continue
        # Named arguments (After:=) and type names after As are not identifiers in scope
        code = re.sub(r"\b[A-Za-z_]\w*\s*:=", " ", code)
        code = re.sub(r"\bAs\s+(New\s+)?[A-Za-z_]\w*", " ", code, flags=re.I)
        code = re.sub(r"\bGoTo\s+[A-Za-z_]\w*", " ", code, flags=re.I)
        code = re.sub(r"\b\d+(\.\d+)?(E[+-]?\d+)?#?", " ", code, flags=re.I)
        for ident in IDENT_RE.findall(code):
            low = ident.lower()
            if low not in known:
                errors.append(f"{mod.path.name}:{n}: undeclared identifier '{ident}' in {proc.name}")
            elif low in private_procs and private_procs[low] != mod.name and low not in mod.module_names:
                errors.append(f"{mod.path.name}:{n}: private procedure '{ident}' of {private_procs[low]} used")
    return errors


def main() -> int:
    modules, errors = [], []
    for path in sorted((p for d in VBA_DIRS for p in d.glob("*.bas")), key=lambda p: p.stem):
        mod, errs = parse(path)
        modules.append(mod)
        errors.extend(errs)
    all_public: dict[str, str] = {}
    private_procs: dict[str, str] = {}
    for mod in modules:
        for name in mod.public_names:
            if name in all_public:
                errors.append(f"{mod.path.name}: public name '{name}' also declared in {all_public[name]}")
            all_public[name] = mod.name
        for pname, proc in mod.procs.items():
            if not proc.public:
                private_procs.setdefault(pname, mod.name)
    for mod in modules:
        for proc in mod.procs.values():
            errors.extend(check_proc(mod, proc, all_public, private_procs))
        # Module-level statements must be declarations or options
        for n, code in mod.module_lines:
            if not re.match(r"^\s*(Option|Attribute|Public|Private|Dim|Const)\b", code, re.I):
                errors.append(f"{mod.path.name}:{n}: executable statement outside a procedure")
    for e in errors:
        print(e)
    print(f"vba_lint: {len(modules)} modules, {len(errors)} problem(s)")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
