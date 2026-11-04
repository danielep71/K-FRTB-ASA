#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Daniele Penza
"""Atomically publish the merged traffic history to its data-only branch.

tools/traffic_history.py merges a collection into a checkout of the
traffic-history branch. This script publishes the files the merge changed
under data/ as one commit on the checked-out parent, through the Git Data API
with the workflow token: one tree, one commit, and a non-forced ref update.
A competing update makes the ref update fail; it is never forced or retried.
It never receives the traffic App token (docs/TRAFFIC_HISTORY.md).
"""

import argparse
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request
from http.client import HTTPMessage
from pathlib import Path
from typing import IO, Any, NoReturn

BRANCH = "traffic-history"
DATA = "data/"
MAX_FILE = 8 * 1024 * 1024
SHA = re.compile(r"[a-f0-9]{40}")
REPOSITORY = re.compile(r"[A-Za-z0-9_-][A-Za-z0-9_.-]*/[A-Za-z0-9_-][A-Za-z0-9_.-]*")
DATA_PATH = re.compile(r"data/[A-Za-z0-9_.-]+(/[A-Za-z0-9_.-]+)?")


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req: urllib.request.Request, fp: IO[bytes], code: int, msg: str,
                         headers: HTTPMessage, newurl: str) -> NoReturn:
        raise RuntimeError("Refusing an API redirect with credentials")


class GitHub:
    def __init__(self, token: str) -> None:
        if not token:
            raise ValueError("Required workflow credential is missing")
        self.token = token
        self.opener = urllib.request.build_opener(NoRedirect())

    def request(self, method: str, path: str, data: dict[str, Any] | None = None) -> Any:
        if not path.startswith("/repos/") or ".." in path:
            raise ValueError("Invalid API path")
        payload = None if data is None else json.dumps(data).encode("utf-8")
        request = urllib.request.Request(
            "https://api.github.com" + path, data=payload, method=method,
            headers={"Authorization": "Bearer " + self.token,
                     "Accept": "application/vnd.github+json",
                     "X-GitHub-Api-Version": "2022-11-28",
                     "Content-Type": "application/json", "User-Agent": "FRTB-traffic-history"})
        for attempt in range(3):
            try:
                with self.opener.open(request, timeout=30) as response:
                    content = response.read(MAX_FILE + 1)
                if len(content) > MAX_FILE:
                    raise ValueError("API response exceeds the size limit")
                return json.loads(content)
            except urllib.error.HTTPError as error:
                if method == "GET" and error.code in (429, 500, 502, 503, 504) and attempt < 2:
                    time.sleep(2 ** attempt)
                    continue
                # Never echo response bodies, headers, or token values.
                raise RuntimeError(f"GitHub {method} failed with HTTP {error.code}") from None
        raise RuntimeError(f"GitHub {method} failed")


def git(history: Path, *args: str) -> bytes:
    return subprocess.run(["git", "-C", str(history), *args], check=True,
                          capture_output=True).stdout


def changed_files(history: Path) -> dict[str, str]:
    """The files under data/ that differ from the checked-out commit.

    The merge never deletes a file, so a deletion is refused rather than
    published.
    """
    listing = git(history, "status", "--porcelain=v1", "-z", "--untracked-files=all", "--", DATA)
    files: dict[str, str] = {}
    for entry in filter(None, listing.decode("utf-8").split("\0")):
        status, path = entry[:2], entry[3:]
        if "D" in status or "R" in status:
            raise ValueError("Refusing to publish a deleted or renamed history file")
        if not DATA_PATH.fullmatch(path):
            raise ValueError("Unexpected history path")
        content = (Path(history) / path).read_bytes()
        if len(content) > MAX_FILE:
            raise ValueError("History file exceeds the size limit")
        files[path] = content.decode("utf-8")
    return files


def publish(writer: GitHub, repository: str, parent: str, files: dict[str, str], message: str) -> str | None:
    """One tree and one commit on the expected parent; the ref is never forced."""
    if not REPOSITORY.fullmatch(repository):
        raise ValueError("Invalid repository identity")
    if not SHA.fullmatch(parent):
        raise ValueError("A full parent SHA is required")
    if not files:
        return None
    base = "/repos/" + repository
    ref = writer.request("GET", base + "/git/ref/heads/" + BRANCH)
    if ref["object"]["sha"] != parent:
        raise RuntimeError("traffic-history moved since the checkout; rerun the collection")
    commit = writer.request("GET", base + "/git/commits/" + parent)
    entries = [{"path": path, "mode": "100644", "type": "blob", "content": content}
               for path, content in sorted(files.items())]
    tree = writer.request("POST", base + "/git/trees", {
        "base_tree": commit["tree"]["sha"], "tree": entries})
    new = writer.request("POST", base + "/git/commits", {
        "message": message, "tree": tree["sha"], "parents": [parent]})
    # A competing update yields a non-fast-forward rejection; never force or reset.
    writer.request("PATCH", base + "/git/refs/heads/" + BRANCH,
                   {"sha": new["sha"], "force": False})
    sha: str = new["sha"]
    return sha


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--history", type=Path, required=True,
                        help="checkout of traffic-history holding the merged data/")
    parser.add_argument("--collected-at", required=True, help="UTC time of the collection")
    args = parser.parse_args(argv)
    if os.environ.get("GITHUB_REF") != "refs/heads/main":
        raise ValueError("Privileged publication is restricted to main")
    repository = os.environ.get("GITHUB_REPOSITORY", "")
    if not REPOSITORY.fullmatch(repository):
        raise ValueError("Invalid repository identity")
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z", args.collected_at):
        raise ValueError("Invalid collection time")
    parent = git(args.history, "rev-parse", "HEAD").decode("ascii").strip()
    files = changed_files(args.history)
    sha = publish(GitHub(os.environ.get("GITHUB_TOKEN", "")), repository, parent, files,
                  "analytics: collect traffic " + args.collected_at)
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if sha is None:
        print("No change to the history; nothing published.")
        return 0
    print("Published " + str(len(files)) + " history files to " + BRANCH + " at " + sha)
    if summary:
        with open(summary, "a", encoding="utf-8") as output:
            output.write("\nPublished commit: [" + sha[:7] + "](https://github.com/"
                         + repository + "/tree/" + sha + ")\n")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (ValueError, RuntimeError, subprocess.CalledProcessError) as error:
        print("ERROR: " + str(error), file=sys.stderr)
        sys.exit(1)
