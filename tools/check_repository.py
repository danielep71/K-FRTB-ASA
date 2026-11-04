#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Offline, tracked-file repository checks; no Excel or regulatory validation."""

import argparse
from datetime import datetime, timezone
import html
import json
import os
from pathlib import Path
import re
import unicodedata
import subprocess
import sys
import time
from urllib.parse import unquote, urlsplit

import yaml


class ConfigLoader(yaml.SafeLoader):
    """Reject duplicate keys and preserve GitHub's YAML `on` key as text."""


ConfigLoader.yaml_implicit_resolvers = {
    key: [(tag, pattern) for tag, pattern in values if tag != "tag:yaml.org,2002:bool"]
    for key, values in yaml.SafeLoader.yaml_implicit_resolvers.items()
}
ConfigLoader.add_implicit_resolver(
    "tag:yaml.org,2002:bool", re.compile(r"^(?:true|false)$", re.IGNORECASE), list("tTfF")
)


def unique_pairs(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate key: {key}")
        result[key] = value
    return result


def yaml_mapping(loader, node):
    return unique_pairs((loader.construct_object(key), loader.construct_object(value)) for key, value in node.value)


ConfigLoader.add_constructor("tag:yaml.org,2002:map", yaml_mapping)


def load_yaml(text):
    return yaml.load(text, Loader=ConfigLoader)


def reject_json_constant(value):
    raise ValueError(f"non-standard JSON constant: {value}")


def binary_paths(root, paths):
    result = subprocess.check_output(
        ["git", "check-attr", "-z", "--stdin", "text"],
        input=("\0".join(paths) + "\0").encode(), cwd=root,
    ).decode().rstrip("\0").split("\0")
    return {result[index] for index in range(0, len(result), 3) if result[index + 2] == "unset"}


def _outside_fences(text):
    """The lines of a Markdown text outside fenced code blocks."""
    fence = None
    for line in text.splitlines():
        match = re.match(r"^\s*(`{3,}|~{3,})", line)
        if match:
            marker = match[1][0]
            fence = None if fence == marker else marker if fence is None else fence
            continue
        if fence is None:
            yield line


def heading_slug(heading):
    """GitHub's anchor for a heading: the rendered text lower-cased, every
    character but letters, marks, digits, connector punctuation, spaces and
    hyphens removed, and each space turned into a hyphen.

    Code spans render their text literally, so they are set aside first and
    restored after links, HTML, emphasis and character references are
    resolved in the rest."""
    spans = []

    def set_aside(match):
        spans.append(match[2].strip(" ") if match[2].strip(" ") else match[2])
        return chr(0xE000 + len(spans) - 1)  # a private-use placeholder

    text = re.sub(r"(`+)(.+?)(?<!`)\1(?!`)", set_aside, heading)
    text = re.sub(r"!?\[([^\]]*)\]\([^)]*\)", r"\1", text)  # links and images: their text
    text = re.sub(r"<[^>]+>", "", text)  # inline HTML tags
    # emphasis delimiters: * anywhere, _ only at word edges (snake_case keeps its _)
    text = re.sub(r"(\*{1,3})(?=\S)(.+?)(?<=\S)\1", r"\2", text)
    text = re.sub(r"(?<!\w)(_{1,3})(?=\S)(.+?)(?<=\S)\1(?!\w)", r"\2", text)
    text = html.unescape(text)  # character references: the character they render
    text = re.sub("[\uE000-\uF8FF]", lambda match: spans[ord(match[0]) - 0xE000], text)
    kept = []
    for char in text.lower():
        category = unicodedata.category(char)
        if char in " -" or category[0] in "LMN" or category == "Pc":
            kept.append(char)
    return "".join(kept).replace(" ", "-")


def setext_text(line):
    """Whether a line can be the text of a Setext heading: not blank and not
    another block (heading, list item, quote, table row, HTML or a rule)."""
    stripped = line.strip()
    return bool(stripped) and not line.startswith("    ") and not re.match(
        r"(#|[-*+][ \t]|\d+[.)][ \t]|>|\||<|(?:[-*_][ \t]*){3,}$)", stripped)


def markdown_anchors(text):
    """The anchors a Markdown file defines, and its duplicated explicit ids.

    Headings give GitHub's slugs: a repeat takes the next free -1, -2 ...
    suffix, checked against every slug emitted so far, so "Scope", "Scope"
    and "Scope-1" give scope, scope-1 and scope-1-1. An HTML id or name
    attribute outside code gives an explicit anchor."""
    counts, explicit, duplicates = {}, set(), set()
    previous = ""
    for line in _outside_fences(text):
        heading = re.match(r"^ {0,3}#{1,6}[ \t]+(.*?)(?:[ \t]+#+)?[ \t]*$", line)
        title = heading[1] if heading else None
        # a Setext heading: a paragraph line underlined with = or -
        if title is None and re.fullmatch(r" {0,3}(=+|-+)[ \t]*", line) and setext_text(previous):
            title = previous.strip()
        previous = line
        if title is not None:
            base = slug = heading_slug(title)
            while slug in counts:
                counts[base] += 1
                slug = f"{base}-{counts[base]}"
            counts[slug] = 0
        for found in re.findall(r"<[A-Za-z][^>]*?\s(?:id|name)\s*=\s*[\"']([^\"']+)[\"']",
                                re.sub(r"`[^`]*`", "", line)):
            if found in explicit:
                duplicates.add(found)
            explicit.add(found)
    return set(counts) | explicit, duplicates


def markdown_targets(text):
    # Common inline links/images and reference definitions outside fenced code.
    body = "\n".join(re.sub(r"`[^`]*`", "", line) for line in _outside_fences(text))
    targets = re.findall(r"\]\(\s*(<[^>]+>|[^\s)]+)(?:\s+[^)]*)?\)", body)
    targets += re.findall(r"^\s*\[[^\]]+\]:\s*(<[^>]+>|\S+)", body, re.MULTILINE)
    return [target.strip("<>") for target in targets]


def validate_files(root, paths, binary=frozenset()):
    errors = []
    tracked = set(paths)
    anchor_cache = {}

    def anchors_of(relative):
        if relative not in anchor_cache:
            anchor_cache[relative] = markdown_anchors((root / relative).read_text(encoding="utf-8"))[0]
        return anchor_cache[relative]

    for name in paths:
        path = root / name
        try:
            data = path.read_bytes()
        except OSError as error:
            errors.append(f"{name}: {error}")
            continue
        if name in binary:
            continue
        if b"\0" in data:
            errors.append(f"{name}: unexpected NUL in text")
            continue
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            if path.suffix.lower() in {".bas", ".cls", ".frm"}:
                text = data.decode("cp1252")
            else:
                errors.append(f"{name}: text must be UTF-8")
                continue
        for number, line in enumerate(text.splitlines(), 1):
            trailing = line[len(line.rstrip(" \t")):]
            hard_break = path.suffix == ".md" and len(trailing) >= 2 and set(trailing) == {" "}
            if trailing and not hard_break:
                errors.append(f"{name}:{number}: trailing whitespace")
        if text and not text.endswith("\n"):
            errors.append(f"{name}: missing final newline")
        if re.search(r"\n[\r\n]+\Z", text):
            errors.append(f"{name}: extra blank line at EOF")
        try:
            if path.suffix == ".json":
                json.loads(text, object_pairs_hook=unique_pairs, parse_constant=reject_json_constant)
            elif path.suffix in {".yml", ".yaml"}:
                document = load_yaml(text)
                if not isinstance(document, dict):
                    raise ValueError("configuration must be a mapping")
                if name.startswith(".github/workflows/"):
                    validate_workflow(document, Path(name).name)
            elif path.suffix == ".md":
                anchors, duplicates = markdown_anchors(text)
                anchor_cache[name] = anchors
                if duplicates:
                    raise ValueError(f"duplicate explicit anchor: {', '.join(sorted(duplicates))}")
                for target in markdown_targets(text):
                    parsed = urlsplit(target)
                    if parsed.scheme or parsed.netloc:
                        continue
                    fragment = unquote(parsed.fragment)
                    if not parsed.path:
                        if fragment and fragment not in anchors and fragment.lower() not in anchors:
                            raise ValueError(f"link fragment not found in this file: {target}")
                        continue
                    destination = (root / parsed.path.lstrip("/")) if parsed.path.startswith("/") else path.parent / unquote(parsed.path)
                    destination = destination.resolve()
                    if not destination.is_relative_to(root.resolve()):
                        raise ValueError(f"link escapes repository: {target}")
                    relative = destination.relative_to(root.resolve()).as_posix()
                    if relative not in tracked and not any(item.startswith(relative.rstrip("/") + "/") for item in tracked):
                        raise ValueError(f"link target is not tracked: {target}")
                    if fragment and relative.endswith(".md") and relative in tracked:
                        known = anchors_of(relative)
                        if fragment not in known and fragment.lower() not in known:
                            raise ValueError(f"link fragment not found in {relative}: {target}")
        except (ValueError, TypeError, yaml.YAMLError) as error:
            errors.append(f"{name}: {error}")
    return errors


# Jobs that may hold a write permission, and the scopes they may write.
WRITE_ALLOWED = {
    ("labels-sync.yml", "reconcile"): {"issues"},
    ("traffic-export.yml", "export"): {"contents"},
    ("traffic-export.yml", "alert"): {"issues"},
    ("pr-milestones.yml", "reconcile"): {"pull-requests"},
}
DEFAULT_BRANCH_ONLY = "github.ref == format('refs/heads/{0}', github.event.repository.default_branch)"
LABEL_RECONCILE_CONDITION = "github.event_name != 'pull_request' && " + DEFAULT_BRANCH_ONLY


def written_scopes(permissions):
    """The scopes a permissions value grants write access to."""
    if permissions in ("write-all", "write"):
        return {"*"}
    if isinstance(permissions, dict):
        return {scope for scope, level in permissions.items() if level == "write"}
    if permissions in (None, "read-all", "read"):
        return set()
    raise ValueError(f"unrecognized permissions value: {permissions!r}")


def validate_workflow(document, name=None):
    """Check pinning, credentials, permissions and timeouts of one workflow.

    ``name`` is the file name under .github/workflows; write permissions are
    allowed only for the jobs listed in WRITE_ALLOWED.
    """
    if not document.get("name") or not document.get("on") or not isinstance(document.get("jobs"), dict):
        raise ValueError("workflow requires name, on, and jobs")
    if "permissions" not in document:
        raise ValueError("workflow must declare top-level permissions")
    for job_id, job in document["jobs"].items():
        if not isinstance(job, dict):
            raise ValueError("workflow job must be a mapping")
        if "timeout-minutes" not in job:
            raise ValueError(f"job {job_id} must set timeout-minutes")
        effective = job["permissions"] if "permissions" in job else document["permissions"]
        extra = written_scopes(effective) - WRITE_ALLOWED.get((name, job_id), set())
        if extra:
            raise ValueError(f"job {job_id} must not get write permission: {', '.join(sorted(extra))}")
        for step in job.get("steps", []):
            action = step.get("uses", "")
            if action and not action.startswith("./") and not re.fullmatch(r"[^\s@]+@[0-9a-f]{40}", action):
                raise ValueError(f"action must use a full commit SHA: {action}")
            if action.startswith("actions/checkout@") and (step.get("with") or {}).get("persist-credentials") is not False:
                raise ValueError(f"job {job_id}: actions/checkout must set persist-credentials: false")
            run = step.get("run", "")
            if "pip install" in run and ("--require-hashes" not in run or "--only-binary=:all:" not in run):
                raise ValueError(f"job {job_id}: pip install must use --require-hashes --only-binary=:all:")
    if name == "labels-sync.yml" and document["jobs"].get("reconcile", {}).get("if") != LABEL_RECONCILE_CONDITION:
        raise ValueError("labels-sync reconcile must run only outside pull requests on the default branch")


def validate_controls(root):
    milestone_workflow = load_yaml((root / ".github/workflows/pr-milestones.yml").read_text())
    validate_milestone_workflow(milestone_workflow)
    workflow = load_yaml((root / ".github/workflows/integrity.yml").read_text())
    for event in ("pull_request", "push"):
        config = workflow["on"][event]
        if set(config) != {"branches"} or set(config["branches"]) != {"main", "release/**"}:
            raise ValueError(f"integrity {event} must cover main and release/** without path filters")
    job = workflow["jobs"]["integrity"]
    if job["name"] != "Repository integrity" or "if" in job or "strategy" in job:
        raise ValueError("integrity must provide one unconditional stable Repository integrity check")
    if workflow["permissions"] != {"contents": "read"} or job["permissions"] != {"contents": "read"}:
        raise ValueError("integrity must use read-only contents permission")
    validate_integrity_reproducibility(workflow)
    milestone_group = milestone_workflow.get("concurrency", {}).get("group", "")
    if "github.event.pull_request.number" not in milestone_group:
        raise ValueError("PR milestone events need one concurrency group per PR")
    for name in (".github/requirements.txt", "tests/requirements.txt"):
        validate_hashed_requirements((root / name).read_text(encoding="utf-8"), name)
    updates = load_yaml((root / ".github/dependabot.yml").read_text())
    if updates["version"] != 2:
        raise ValueError("Dependabot requires version 2")
    expected = {
        ("github-actions", "/", "release/v0.2.0"),
        ("pip", "/.github", "release/v0.2.0"),
        ("pip", "/tests", "release/v0.2.0"),
    }
    actual = [(item["package-ecosystem"], item["directory"], item["target-branch"]) for item in updates["updates"]]
    if set(actual) != expected or len(actual) != len(expected):
        raise ValueError("Dependabot must cover Actions, CI parser and prototype packages once each on release/v0.2.0")
    for item in updates["updates"]:
        if "milestone" in item:
            raise ValueError("Dependabot PRs must have no milestone setting")
        if item["assignees"] != ["danielep71"] or [label for label in item["labels"] if label in {"P1", "P2", "P3"}] != ["P3"]:
            raise ValueError("Dependabot requires owner and exactly P3")


INTEGRITY_CONCURRENCY = {
    "group": "${{ github.event_name == 'pull_request' && format('integrity-pr-{0}', "
             "github.event.pull_request.number) || format('integrity-run-{0}', github.run_id) }}",
    "cancel-in-progress": "${{ github.event_name == 'pull_request' }}",
}
EXACT_VERSION = re.compile(r"\d+\.\d+\.\d+")


def validate_integrity_reproducibility(workflow):
    """Exact runtimes, hashed installs, a report always uploaded, and push
    runs that no later push cancels."""
    if workflow.get("concurrency") != INTEGRITY_CONCURRENCY:
        raise ValueError("integrity must cancel only superseded PR runs, never push runs")
    steps = workflow["jobs"]["integrity"]["steps"]

    def step(action):
        found = [item for item in steps if item.get("uses", "").startswith(action + "@")]
        if len(found) != 1:
            raise ValueError(f"integrity must use {action} once")
        return found[0]

    for action, key in (("actions/setup-python", "python-version"), ("actions/setup-node", "node-version")):
        if not EXACT_VERSION.fullmatch(str((step(action).get("with") or {}).get(key, ""))):
            raise ValueError(f"integrity {action} must set an exact {key}")
    installs = [item["run"] for item in steps if "pip install" in item.get("run", "")]
    if not installs or any("--require-hashes" not in run or "--only-binary=:all:" not in run for run in installs):
        raise ValueError("integrity must install Python packages by hash, from wheels only")
    upload = step("actions/upload-artifact")
    options = upload.get("with") or {}
    if upload.get("if") != "always()" or options.get("if-no-files-found") != "error" or not options.get("retention-days"):
        raise ValueError("integrity must always upload its report and fail when it is missing")
    if not any("--report" in item.get("run", "") for item in steps):
        raise ValueError("integrity must write its report")


def validate_hashed_requirements(text, name):
    """Every requirement is pinned with == and carries at least one hash."""
    requirements = re.split(r"(?<!\\)\n(?=\S)", re.sub(r"(?m)^\s*#.*\n", "", text))
    for requirement in filter(str.strip, requirements):
        if not re.match(r"[A-Za-z0-9_.-]+==[^\s;\\]+", requirement) or "--hash=sha256:" not in requirement:
            raise ValueError(f"{name}: every requirement needs an exact pin and a hash: {requirement.split()[0]}")


MILESTONE_RECONCILE_CONDITION = (
    "github.event_name == 'pull_request_target' || "
    "github.event_name == 'schedule' || " + DEFAULT_BRANCH_ONLY
)


def validate_milestone_workflow(document):
    events = document["on"]
    if set(events) != {"push", "pull_request_target", "schedule", "workflow_dispatch"} or events["push"] != {"branches": ["main"], "paths": [".github/workflows/pr-milestones.yml"]} or events["pull_request_target"] != {"types": ["opened", "reopened", "edited", "milestoned"]} or events["schedule"] != [{"cron": "17 3 * * *"}]:
        raise ValueError("PR milestone workflow requires event and full-scan coverage")
    job = document["jobs"]["reconcile"]
    if job.get("if") != MILESTONE_RECONCILE_CONDITION:
        raise ValueError("PR milestone reconcile must run only on pull_request_target, schedule or the default branch")
    if document["permissions"] != {} or job["permissions"] != {"pull-requests": "write"}:
        raise ValueError("PR milestone workflow permissions must be limited to metadata")
    steps = job["steps"]
    if len(steps) != 1 or set(steps[0]) != {"name", "env", "run"} or steps[0]["env"] != {"GH_TOKEN": "${{ github.token }}"} or "${{" in steps[0]["run"]:
        raise ValueError("PR milestone workflow must use trusted inline code without checkout or event interpolation")


def run_checks(root, checks):
    """Run every check in order, appending one record per check to ``checks``.

    It stops at the first failure, as before; the record of the failed check
    holds its errors or exit code. Returns the process exit code."""
    def run(name, command):
        record = {"name": name, "status": "FAIL"}
        checks.append(record)
        started = time.monotonic()
        try:
            subprocess.run(command, cwd=root, check=True)
        finally:
            record["seconds"] = round(time.monotonic() - started, 3)
        record["status"] = "PASS"

    paths = subprocess.check_output(["git", "ls-files", "-z"], cwd=root).decode().rstrip("\0").split("\0")
    errors = validate_files(root, paths, binary_paths(root, paths))
    try:
        validate_controls(root)
    except (OSError, KeyError, TypeError, ValueError, yaml.YAMLError) as error:
        errors.append(f"CI controls: {error}")
    checks.append({"name": "configuration, tracked whitespace, local Markdown targets and CI controls",
                   "status": "FAIL" if errors else "PASS", "detail": errors or f"{len(paths)} files"})
    if errors:
        print("\n".join(errors), file=sys.stderr)
        return 1
    print(f"PASS: configuration, tracked whitespace and local Markdown targets ({len(paths)} files)", flush=True)
    commands = [
        ["node", ".github/scripts/labels-sync.mjs", "--self-test"],
        ["node", ".github/scripts/labels-sync.mjs", "--mode", "validate"],
        ["node", ".github/scripts/labels-drift.mjs", "--self-test"],
        [sys.executable, "tools/check_issues.py", "--check-config"],
        [sys.executable, "-m", "unittest", "discover", "-s", "tests/tooling", "-p", "test_*.py", "-v"],
    ]
    for command in commands:
        run(" ".join("python" if part == sys.executable else part for part in command), command)
    print("PASS: repository integrity (maintenance only; no Excel or regulatory certification)")
    return 0


def checked_commit(root):
    """The commit of the checked repository: GITHUB_SHA in CI, else its HEAD.

    It never falls back to the caller's working directory; a failed lookup
    is recorded as "unknown", not as an empty or unrelated commit."""
    if os.environ.get("GITHUB_SHA"):
        return os.environ["GITHUB_SHA"]
    try:
        found = subprocess.run(["git", "rev-parse", "HEAD"], cwd=root, capture_output=True, text=True, check=True)
    except (OSError, subprocess.CalledProcessError):
        return "unknown"
    return found.stdout.strip() or "unknown"


def write_report(path, checks, code, versions, root=Path(__file__).resolve().parents[1]):
    """The JSON report of one run, bound to its commit, run and attempt."""
    commit = checked_commit(root)
    report = {
        # the commit tested (on a PR, GitHub's merge commit) and the head it was built from
        "commit": commit,
        "head_commit": os.environ.get("REPORT_HEAD_SHA") or commit,
        "run_id": os.environ.get("GITHUB_RUN_ID", ""),
        "run_attempt": os.environ.get("GITHUB_RUN_ATTEMPT", ""),
        "finished_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "result": "PASS" if code == 0 else "FAIL",
        "runtimes": versions,
        "checks": checks,
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")


def runtime_versions():
    def version(command):
        try:
            return subprocess.run(command, capture_output=True, text=True, check=False).stdout.strip()
        except OSError:
            return "not found"
    return {"python": sys.version.split()[0], "node": version(["node", "--version"]),
            "git": version(["git", "--version"]), "pyyaml": yaml.__version__}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path,
                        help="also write a JSON report of every check run, whatever the result")
    args = parser.parse_args(argv)
    root = Path(__file__).resolve().parents[1]
    checks, code = [], 1
    try:
        code = run_checks(root, checks)
        return code
    except Exception as error:
        detail = f"exit code {error.returncode}" if isinstance(error, subprocess.CalledProcessError) else str(error)
        if checks and checks[-1]["status"] == "FAIL" and "detail" not in checks[-1]:
            checks[-1]["detail"] = detail
        else:
            checks.append({"name": "repository integrity", "status": "FAIL", "detail": detail})
        raise
    finally:
        if args.report:
            write_report(args.report, checks, code, runtime_versions(), root)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (OSError, subprocess.CalledProcessError) as error:
        print(f"Repository integrity failed: {error}", file=sys.stderr)
        sys.exit(1)
