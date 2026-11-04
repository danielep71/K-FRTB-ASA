#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Check the external links of the tracked Markdown files; never blocking.

tools/check_repository.py checks local links and their anchors offline, in
the required gate. This script requests each http(s) link once and sorts the
results into three groups:

* verified: a 2xx answer, after redirects;
* unverifiable: 403, 429, a 5xx answer or a network error, which say
  nothing about the page; reported as a warning, never as a pass;
* broken: any other answer, such as 404 or 410; the run fails.

Links to this repository on github.com are skipped: the local gate covers
them. Each link is requested with a 25-second timeout, up to two retries
for an unverifiable answer, and at most three requests at a time. The
weekly and manual workflow link-check.yml runs it; it is not a required check.

    python3 tools/check_external_links.py [--summary FILE]
"""

from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
import os
from pathlib import Path
import subprocess
import sys
import time
from typing import Callable
import urllib.error
import urllib.request
from urllib.parse import urlsplit, urlunsplit

sys.path.insert(0, str(Path(__file__).resolve().parent))
from check_repository import markdown_targets  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OWN_REPOSITORY = "github.com/danielep71/k-frtb-sa"
TIMEOUT = 25
RETRIES = 2
WORKERS = 3
USER_AGENT = "K-FRTB-SA link check"

VERIFIED, UNVERIFIABLE, BROKEN = "verified", "unverifiable", "broken"


@dataclass(frozen=True)
class Result:
    url: str
    group: str
    detail: str


def own_link(url: str) -> bool:
    parts = urlsplit(url)
    path = (parts.netloc + parts.path).lower().rstrip("/")
    return path == OWN_REPOSITORY or path.startswith(OWN_REPOSITORY + "/")


def external_links(root: Path) -> list[str]:
    """The http(s) links of the tracked Markdown files, without fragments, once each."""
    listing = subprocess.run(["git", "-C", str(root), "ls-files", "-z", "*.md"],
                             capture_output=True, check=True).stdout.decode()
    links = set()
    for name in filter(None, listing.split("\0")):
        for target in markdown_targets((root / name).read_text(encoding="utf-8")):
            parts = urlsplit(target)
            if parts.scheme in ("http", "https") and parts.netloc and not own_link(target):
                links.add(urlunsplit(parts._replace(fragment="")))
    return sorted(links)


def classify(status: int | None) -> str:
    if status is None or status in (403, 429) or status >= 500:
        return UNVERIFIABLE
    return VERIFIED if 200 <= status < 300 else BROKEN


def request(url: str) -> int:
    """The HTTP status of a GET for ``url`` after redirects; raises on network errors."""
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as response:  # noqa: S310 (http(s) only)
            response.read(1)
            return int(response.status)
    except urllib.error.HTTPError as error:
        return int(error.code)


def check(url: str, fetch: Callable[[str], int] = request,
          sleep: Callable[[float], None] = time.sleep) -> Result:
    detail = ""
    for attempt in range(RETRIES + 1):
        try:
            status: int | None = fetch(url)
            detail = f"HTTP {status}"
        except (urllib.error.URLError, OSError, ValueError) as error:
            status, detail = None, f"network error: {type(error).__name__}"
        group = classify(status)
        if group != UNVERIFIABLE or attempt == RETRIES:
            return Result(url, group, detail)
        sleep(2 ** attempt)
    return Result(url, UNVERIFIABLE, detail)  # pragma: no cover


def check_all(urls: list[str], fetch: Callable[[str], int] = request,
              sleep: Callable[[float], None] = time.sleep) -> list[Result]:
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        return list(pool.map(lambda url: check(url, fetch, sleep), urls))


def report(results: list[Result]) -> str:
    def cell(text: str) -> str:
        return text.replace("|", "\\|").replace("\r", " ").replace("\n", " ")

    counts = {group: sum(r.group == group for r in results) for group in (VERIFIED, UNVERIFIABLE, BROKEN)}
    lines = ["# External links", "",
             f"- Verified: **{counts[VERIFIED]}**",
             f"- Unverifiable (warning, not a pass): **{counts[UNVERIFIABLE]}**",
             f"- Broken: **{counts[BROKEN]}**", ""]
    for group in (BROKEN, UNVERIFIABLE):
        listed = [r for r in results if r.group == group]
        if listed:
            lines += [f"## {group.capitalize()}", "", "| Link | Result |", "| --- | --- |"]
            lines += [f"| <{cell(r.url)}> | {cell(r.detail)} |" for r in listed]
            lines.append("")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--summary", type=Path, help="also append the report to this file")
    args = parser.parse_args(argv)
    results = check_all(external_links(ROOT))
    text = report(results)
    print(text)
    target = args.summary or (Path(os.environ["GITHUB_STEP_SUMMARY"]) if os.environ.get("GITHUB_STEP_SUMMARY") else None)
    if target:
        with target.open("a", encoding="utf-8") as output:
            output.write(text + "\n")
    for result in results:
        if result.group == UNVERIFIABLE:
            print(f"::warning::Unverifiable link {result.url}: {result.detail}")
    return 1 if any(r.group == BROKEN for r in results) else 0


if __name__ == "__main__":
    sys.exit(main())
