# SPDX-License-Identifier: MIT
"""Offline tests of tools/check_external_links.py: no request leaves the machine."""

import subprocess
import sys
import tempfile
import unittest
import urllib.error
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools"))
import check_external_links as links


class ExternalLinkTests(unittest.TestCase):
    def test_classification(self):
        for status, group in [(200, links.VERIFIED), (204, links.VERIFIED), (404, links.BROKEN), (410, links.BROKEN),
                              (400, links.BROKEN), (403, links.UNVERIFIABLE), (429, links.UNVERIFIABLE),
                              (500, links.UNVERIFIABLE), (503, links.UNVERIFIABLE), (None, links.UNVERIFIABLE)]:
            with self.subTest(status=status):
                self.assertEqual(links.classify(status), group)

    def test_unverifiable_answers_are_retried_and_never_pass(self):
        calls: list[str] = []
        sleeps: list[float] = []

        def flaky(url):
            calls.append(url)
            return 503

        result = links.check("https://example.com/a", flaky, sleeps.append)
        self.assertEqual((result.group, len(calls), sleeps), (links.UNVERIFIABLE, links.RETRIES + 1, [1, 2]))

        answers = iter([429, 200])
        result = links.check("https://example.com/b", lambda url: next(answers), lambda _: None)
        self.assertEqual(result.group, links.VERIFIED)

        def offline(url):
            raise urllib.error.URLError("no route")

        result = links.check("https://example.com/c", offline, lambda _: None)
        self.assertEqual(result.group, links.UNVERIFIABLE)
        self.assertIn("network error", result.detail)

    def test_broken_answers_are_not_retried(self):
        calls: list[str] = []

        def gone(url):
            calls.append(url)
            return 404

        result = links.check("https://example.com/gone", gone, lambda _: None)
        self.assertEqual((result.group, len(calls)), (links.BROKEN, 1))

    def test_links_are_collected_once_without_fragments_or_own_repository(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            (root / "a.md").write_text(
                "[One](https://example.com/page#part) [Again](https://example.com/page)\n"
                "[Own](https://github.com/danielep71/K-FRTB-SA/issues/1) [Local](docs/b.md)\n"
                "```\n[Fenced](https://example.org/hidden)\n```\n[Other](http://example.net/x)\n")
            subprocess.run(["git", "-C", str(root), "add", "a.md"], check=True)
            self.assertEqual(links.external_links(root), ["http://example.net/x", "https://example.com/page"])

    def test_report_lists_broken_and_unverifiable_links(self):
        results = [links.Result("https://a.example", links.VERIFIED, "HTTP 200"),
                   links.Result("https://b.example", links.BROKEN, "HTTP 404"),
                   links.Result("https://c.example/x|y", links.UNVERIFIABLE, "HTTP 429")]
        text = links.report(results)
        self.assertIn("- Verified: **1**", text)
        self.assertIn("- Broken: **1**", text)
        self.assertIn("| <https://b.example> | HTTP 404 |", text)
        self.assertIn("| <https://c.example/x\\|y> | HTTP 429 |", text)

    def test_main_fails_only_on_broken_links(self):
        original = (links.external_links, links.check_all)
        try:
            links.external_links = lambda root: ["https://a.example"]
            with tempfile.TemporaryDirectory() as directory:
                summary = Path(directory) / "summary.md"
                links.check_all = lambda urls, *_: [links.Result(urls[0], links.UNVERIFIABLE, "HTTP 503")]
                self.assertEqual(links.main(["--summary", str(summary)]), 0)
                links.check_all = lambda urls, *_: [links.Result(urls[0], links.BROKEN, "HTTP 404")]
                self.assertEqual(links.main(["--summary", str(summary)]), 1)
                self.assertIn("# External links", summary.read_text())
        finally:
            links.external_links, links.check_all = original


if __name__ == "__main__":
    unittest.main()
