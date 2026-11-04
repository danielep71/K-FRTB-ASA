# SPDX-License-Identifier: MPL-2.0
# This Source Code Form is subject to the terms of the Mozilla Public License,
# v. 2.0. If a copy of the MPL was not distributed with this file, You can
# obtain one at https://mozilla.org/MPL/2.0/.
# See docs/licenses/THIRD_PARTY_NOTICES.txt.
"""Tests for the traffic history merge (tools/traffic_history.py)."""

from __future__ import annotations

import contextlib
import csv
import io
import json
import sys
import tempfile
import unittest
from collections.abc import Iterator, Sequence
from datetime import UTC, date, datetime, timedelta
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools"))

import check_issues  # noqa: E402
import traffic_history  # noqa: E402
from traffic_history import DAILY, PATHS, REFERRERS, SNAPSHOTS, TrafficError  # noqa: E402

Day = tuple[str, int, int]


def window(last: str, values: Sequence[tuple[int, int]]) -> list[Day]:
    """Days ending on ``last``, oldest first, one per value pair."""
    end = date.fromisoformat(last)
    return [((end - timedelta(days=len(values) - 1 - offset)).isoformat(), total, unique)
            for offset, (total, unique) in enumerate(values)]


def series(key: str, days: Sequence[Day]) -> dict[str, Any]:
    return {
        "count": sum(total for _, total, _ in days),
        "uniques": max((unique for _, _, unique in days), default=0),
        key: [{"timestamp": f"{day}T00:00:00Z", "count": total, "uniques": unique}
              for day, total, unique in days],
    }


def stats(stars: int = 1, forks: int = 0, watchers: int = 1, open_issues: int = 4) -> dict[str, Any]:
    """The fields of a repository API response that a collection keeps."""
    return {"stargazers_count": stars, "forks_count": forks,
            "subscribers_count": watchers, "open_issues_count": open_issues, "name": "r"}


def at(stamp: str) -> datetime:
    return datetime.fromisoformat(stamp).replace(tzinfo=UTC)


class Workspace:
    """A raw response directory and a history directory under one temporary root."""

    def __init__(self, root: Path) -> None:
        self.raw = root / "raw"
        self.data = root / "data"
        self.raw.mkdir()

    def collect(self, views: Sequence[Day], clones: Sequence[Day] | None = None,
                referrers: Any = None, paths: Any = None, repository: Any = None) -> None:
        responses: dict[str, Any] = {
            "views.json": series("views", views),
            "clones.json": series("clones", views if clones is None else clones),
            "referrers.json": [{"referrer": "github.com", "count": 3, "uniques": 1}]
            if referrers is None else referrers,
            "paths.json": [{"path": "/o/r", "title": "r", "count": 3, "uniques": 1}]
            if paths is None else paths,
            "repository.json": stats() if repository is None else repository,
        }
        for name, body in responses.items():
            (self.raw / name).write_text(json.dumps(body), encoding="utf-8")

    def merge(self, stamp: str) -> traffic_history.Report:
        return traffic_history.merge(self.raw, self.data, at(stamp))

    def rows(self, name: str) -> list[dict[str, str]]:
        with (self.data / name).open(encoding="utf-8", newline="") as handle:
            return list(csv.DictReader(handle))

    def daily(self) -> dict[str, tuple[str, str, str]]:
        """Views, clones and the views' collection time, by date."""
        return {row["date"]: (row["views"], row["clones"], row["views_collected_at"])
                for row in self.rows(DAILY)}

    def snapshot(self) -> dict[str, str]:
        return {str(path.relative_to(self.data)): path.read_text(encoding="utf-8")
                for path in sorted(self.data.rglob("*")) if path.is_file()}

    def badge(self, name: str) -> dict[str, Any]:
        result: dict[str, Any] = json.loads(
            (self.data / "badges" / f"{name}.json").read_text(encoding="utf-8"))
        return result


@contextlib.contextmanager
def workspace() -> Iterator[Workspace]:
    with tempfile.TemporaryDirectory() as directory:
        yield Workspace(Path(directory))


class MergeTests(unittest.TestCase):
    def test_first_collection_creates_every_file(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 14))
            report = ws.merge("2026-10-14T06:17:00")
            self.assertEqual(sorted(report.changed),
                             sorted([*traffic_history.HEADERS, "badges/clones.json",
                                     "badges/visitors.json", "badges/visits.json"]))
            self.assertEqual(len(ws.rows(DAILY)), 14)
            snapshot = ws.rows(SNAPSHOTS)[0]
            self.assertEqual((snapshot["first_date"], snapshot["last_date"]),
                             ("2026-10-01", "2026-10-14"))
            self.assertEqual(snapshot["collected_at"], "2026-10-14T06:17:00Z")
            self.assertEqual(ws.rows(REFERRERS)[0]["rank"], "1")

    def test_overlapping_windows_keep_older_history_and_take_newer_values(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 13 + [(2, 1)]))
            ws.merge("2026-10-14T06:00:00")
            # A week later: 2026-10-14 is now complete (5 views), and the
            # window has moved on by seven days.
            ws.collect(window("2026-10-21", [(1, 1)] * 6 + [(5, 2)] + [(3, 1)] * 7))
            ws.merge("2026-10-21T06:00:00")
            daily = ws.daily()
            self.assertEqual(len(daily), 21)
            self.assertEqual(daily["2026-10-01"], ("1", "1", "2026-10-14T06:00:00Z"))
            self.assertEqual(daily["2026-10-14"], ("5", "5", "2026-10-21T06:00:00Z"))
            # An overlapping day with unchanged values keeps its first collection time.
            self.assertEqual(daily["2026-10-08"][2], "2026-10-14T06:00:00Z")
            self.assertEqual(len(ws.rows(SNAPSHOTS)), 2)

    def test_repeated_run_changes_nothing(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(2, 1)] * 14))
            ws.merge("2026-10-14T06:00:00")
            before = ws.snapshot()
            report = ws.merge("2026-10-14T09:30:00")
            self.assertEqual(report.changed, {})
            self.assertEqual(ws.snapshot(), before)

    def test_same_day_run_with_new_values_replaces_that_day_only(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-13", [(1, 1)] * 14))
            ws.merge("2026-10-13T06:00:00")
            ws.collect(window("2026-10-14", [(1, 1)] * 14))
            ws.merge("2026-10-14T06:00:00")
            ws.collect(window("2026-10-14", [(1, 1)] * 13 + [(4, 2)]),
                       referrers=[{"referrer": "example.org", "count": 4, "uniques": 2}])
            ws.merge("2026-10-14T18:00:00")
            snapshots = ws.rows(SNAPSHOTS)
            self.assertEqual([row["snapshot_date"] for row in snapshots],
                             ["2026-10-13", "2026-10-14"])
            self.assertEqual(snapshots[1]["collected_at"], "2026-10-14T18:00:00Z")
            referrers = ws.rows(REFERRERS)
            self.assertEqual([(row["snapshot_date"], row["referrer"]) for row in referrers],
                             [("2026-10-13", "github.com"), ("2026-10-14", "example.org")])

    def test_replaying_an_older_collection_keeps_newer_values(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 13 + [(2, 1)]))
            ws.raw.rename(ws.raw.parent / "morning")
            ws.raw.mkdir()
            ws.collect(window("2026-10-14", [(1, 1)] * 13 + [(5, 2)]),
                       referrers=[{"referrer": "example.org", "count": 5, "uniques": 2}])
            ws.merge("2026-10-14T18:00:00")
            evening = ws.snapshot()
            # The 06:00 collection is merged after the 18:00 one, as when a
            # saved collection is replayed or rebuilt on a newer branch tip.
            report = traffic_history.merge(ws.raw.parent / "morning", ws.data,
                                           at("2026-10-14T06:00:00"))
            self.assertEqual(report.changed, {})
            self.assertEqual(ws.snapshot(), evening)
            self.assertEqual(ws.daily()["2026-10-14"], ("5", "5", "2026-10-14T18:00:00Z"))
            # An older collection still fills dates the newer one never reported.
            (ws.raw.parent / "morning" / "views.json").write_text(json.dumps(
                series("views", window("2026-10-13", [(1, 1)] * 14))), encoding="utf-8")
            traffic_history.merge(ws.raw.parent / "morning", ws.data, at("2026-10-14T06:00:00"))
            self.assertEqual(ws.daily()["2026-09-30"][0], "1")
            self.assertEqual(ws.daily()["2026-10-14"][:2], ("5", "5"))

    def test_older_collection_fills_a_metric_the_newer_one_did_not_report(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 14),
                       clones=window("2026-10-14", [(1, 1)] * 13 + [(2, 1)]))
            ws.raw.rename(ws.raw.parent / "morning")
            ws.raw.mkdir()
            # The evening collection reports views for 2026-10-14 but no clones.
            ws.collect(window("2026-10-14", [(1, 1)] * 13 + [(5, 2)]),
                       clones=window("2026-10-13", [(1, 1)] * 13))
            ws.merge("2026-10-14T18:00:00")
            traffic_history.merge(ws.raw.parent / "morning", ws.data, at("2026-10-14T06:00:00"))
            row = next(row for row in ws.rows(DAILY) if row["date"] == "2026-10-14")
            self.assertEqual((row["views"], row["views_collected_at"]),
                             ("5", "2026-10-14T18:00:00Z"))
            self.assertEqual((row["clones"], row["clones_collected_at"]),
                             ("2", "2026-10-14T06:00:00Z"))

    def test_replay_keeps_newer_rankings_when_totals_are_unchanged(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 14))
            ws.raw.rename(ws.raw.parent / "morning")
            ws.raw.mkdir()
            traffic_history.merge(ws.raw.parent / "morning", ws.data, at("2026-10-14T06:00:00"))
            ws.collect(window("2026-10-14", [(1, 1)] * 14),
                       referrers=[{"referrer": "example.org", "count": 4, "uniques": 2}])
            ws.merge("2026-10-14T18:00:00")
            self.assertEqual(ws.rows(SNAPSHOTS)[0]["collected_at"], "2026-10-14T18:00:00Z")
            evening = ws.snapshot()
            report = traffic_history.merge(ws.raw.parent / "morning", ws.data,
                                           at("2026-10-14T06:00:00"))
            self.assertEqual(report.changed, {})
            self.assertEqual(ws.snapshot(), evening)
            self.assertEqual(ws.rows(REFERRERS)[0]["referrer"], "example.org")

    def test_zero_activity_is_stored_and_uncollected_dates_are_unavailable(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(0, 0)] * 14))
            ws.merge("2026-10-14T06:00:00")
            ws.collect(window("2026-11-10", [(0, 0)] * 13 + [(1, 1)]))
            report = ws.merge("2026-11-10T06:00:00")
            daily = ws.daily()
            self.assertEqual(daily["2026-10-01"][:2], ("0", "0"))
            self.assertNotIn("2026-10-20", daily)
            self.assertEqual(report.gaps, [("2026-10-15", "2026-10-27")])
            self.assertIn("2026-10-15 to 2026-10-27",
                          traffic_history.summary(report, at("2026-11-10T06:00:00")))

    def test_metric_reported_for_one_series_only_leaves_the_other_blank(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 14),
                       clones=window("2026-10-13", [(1, 1)] * 13))
            ws.merge("2026-10-14T06:00:00")
            self.assertEqual(ws.daily()["2026-10-14"][:2], ("1", ""))


class RepositoryCountTests(unittest.TestCase):
    def test_counts_are_stored_in_the_snapshot(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 14),
                       repository=stats(stars=3, forks=1, watchers=2, open_issues=7))
            ws.merge("2026-10-14T06:00:00")
            row = ws.rows(SNAPSHOTS)[0]
            self.assertEqual([row[key] for key in ("stars", "forks", "watchers", "open_issues")],
                             ["3", "1", "2", "7"])

    def test_a_count_change_alone_replaces_that_day(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 14))
            ws.merge("2026-10-14T06:00:00")
            ws.collect(window("2026-10-14", [(1, 1)] * 14), repository=stats(stars=2))
            ws.merge("2026-10-14T18:00:00")
            row = ws.rows(SNAPSHOTS)[0]
            self.assertEqual((row["stars"], row["collected_at"]), ("2", "2026-10-14T18:00:00Z"))

    def test_legacy_snapshots_are_kept_with_blank_counts(self) -> None:
        legacy = ("snapshot_date,collected_at,first_date,last_date,"
                  "views_14d,unique_views_14d,clones_14d,unique_clones_14d\n"
                  "2026-10-13,2026-10-13T06:00:00Z,2026-09-30,2026-10-13,7,1,10,3\n")
        with workspace() as ws:
            ws.data.mkdir()
            (ws.data / SNAPSHOTS).write_text(legacy, encoding="utf-8")
            ws.collect(window("2026-10-14", [(1, 1)] * 14))
            ws.merge("2026-10-14T06:00:00")
            old, new = ws.rows(SNAPSHOTS)
            self.assertEqual((old["views_14d"], old["stars"], old["open_issues"]), ("7", "", ""))
            self.assertEqual(new["stars"], "1")
            self.assertEqual((ws.data / SNAPSHOTS).read_text(encoding="utf-8").splitlines()[0],
                             ",".join(traffic_history.HEADERS[SNAPSHOTS]))

    def test_invalid_repository_response_is_refused(self) -> None:
        for body, message in [([], "repository: expected an object"),
                              (stats() | {"forks_count": None}, "repository.forks_count"),
                              (stats() | {"stargazers_count": -1}, "non-negative")]:
            with self.subTest(message=message), workspace() as ws:
                ws.collect(window("2026-10-14", [(1, 1)] * 14), repository=body)
                with self.assertRaisesRegex(TrafficError, message):
                    ws.merge("2026-10-14T06:00:00")
                self.assertFalse(ws.data.exists())


class BadgeTests(unittest.TestCase):
    def test_badges_count_clones_visitors_and_visits(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(2, 1)] * 14),
                       clones=window("2026-10-14", [(100, 3)] * 14))
            ws.merge("2026-10-14T06:00:00")
            self.assertEqual(ws.badge("clones"), {"schemaVersion": 1, "label": "clones",
                                                  "message": "1.4k", "color": "217346"})
            # The latest snapshot's 14-day unique count, not a sum of daily uniques.
            self.assertEqual(ws.badge("visitors")["message"], "1")
            self.assertEqual(ws.badge("visitors")["label"], "visitors (14d)")
            self.assertEqual(ws.badge("visits")["message"], "14")
            # A week later the visits badge keeps the older days and adds the new ones.
            ws.collect(window("2026-10-21", [(2, 1)] * 14),
                       clones=window("2026-10-21", [(100, 3)] * 14))
            ws.merge("2026-10-21T06:00:00")
            self.assertEqual(ws.badge("visits")["message"], "21")
            self.assertEqual(ws.badge("clones")["message"], "2.1k")

    def test_badges_are_rewritten_only_when_they_change(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(2, 1)] * 14))
            ws.merge("2026-10-14T06:00:00")
            (ws.data / "badges" / "visits.json").unlink()
            report = ws.merge("2026-10-14T09:00:00")
            self.assertEqual(report.changed, {"badges/visits.json": 1})
            self.assertEqual(ws.merge("2026-10-14T12:00:00").changed, {})

    def test_compact_numbers(self) -> None:
        self.assertEqual([traffic_history.compact(value)
                          for value in (0, 999, 1000, 1250, 11_000, 2_500_000)],
                         ["0", "999", "1k", "1.2k", "11k", "2.5M"])


class AlertTests(unittest.TestCase):
    def history(self, ws: Workspace, before: tuple[int, int, int], after: tuple[int, int, int],
                referrers: Any = None, milestone_due: str = "2026-12-20") -> dict[str, str] | None:
        """Two daily snapshots of (views, clones, stars); the second may add referrers."""
        for day, (views, clones, stars), extra in (("2026-10-13", before, None),
                                                  ("2026-10-14", after, referrers)):
            ws.collect([(day, views, 1)], clones=[(day, clones, 1)],
                       repository=stats(stars=stars), referrers=extra)
            ws.merge(f"{day}T06:00:00")
        return traffic_history.alert(ws.data, milestone_due)

    def test_no_alert_when_nothing_stands_out(self) -> None:
        with workspace() as ws:
            ws.collect([("2026-10-13", 5, 1)])
            ws.merge("2026-10-13T06:00:00")
            self.assertIsNone(traffic_history.alert(ws.data, "2026-12-20"))
            self.assertIsNone(self.history(ws, (50, 20, 1), (99, 29, 1)))

    def test_spikes_and_more_stars_raise_an_alert(self) -> None:
        with workspace() as ws:
            issue = self.history(ws, (50, 20, 1), (150, 40, 2))
            assert issue is not None
            self.assertEqual(issue["title"], "Traffic alert: 2026-10-14")
            self.assertIn("views spike, clones spike, stars increased.", issue["body"])
            self.assertIn("| Stars | 2 | 1 |", issue["body"])
            for line in ("- **Planned start:** 2026-10-14", "- **Target completion date:** 2026-10-21",
                         "- **Milestone target:** 2026-12-20",
                         "## Dependencies and blocking status\n\nNone."):
                self.assertIn(line, issue["body"])
            # issue triage reads the alert's target date from its planning line
            self.assertEqual(check_issues.target_date(issue["body"]), "2026-10-21")

    def test_the_alert_target_never_passes_the_milestone_due_date(self) -> None:
        with workspace() as ws:
            issue = self.history(ws, (50, 20, 1), (150, 40, 2), milestone_due="2026-10-20")
            assert issue is not None
            self.assertIn("- **Target completion date:** 2026-10-20", issue["body"])
            self.assertEqual(check_issues.target_date(issue["body"]), "2026-10-20")

    def test_a_spike_needs_double_and_the_minimum_growth(self) -> None:
        with workspace() as ws:
            # Clones double (4 to 8) but grow by less than 10: no spike.
            self.assertIsNone(self.history(ws, (1, 4, 1), (1, 8, 1)))

    def test_new_referrer_needs_enough_views(self) -> None:
        few = [{"referrer": "github.com", "count": 3, "uniques": 1},
               {"referrer": "example.org", "count": 4, "uniques": 2}]
        with workspace() as ws:
            self.assertIsNone(self.history(ws, (10, 1, 1), (10, 1, 1), few))
        many = [{"referrer": "github.com", "count": 30, "uniques": 1},
                {"referrer": "example.org", "count": 5, "uniques": 2}]
        with workspace() as ws:
            issue = self.history(ws, (10, 1, 1), (10, 1, 1), many)
            assert issue is not None
            self.assertIn("- example.org: 5 views, 2 unique", issue["body"])
            self.assertNotIn("New referrers** (at least 5 views):\n\n- github.com", issue["body"])

    def test_counts_missing_from_the_older_snapshot_raise_nothing(self) -> None:
        legacy = ("snapshot_date,collected_at,first_date,last_date,"
                  "views_14d,unique_views_14d,clones_14d,unique_clones_14d\n"
                  "2026-10-13,2026-10-13T06:00:00Z,2026-10-13,2026-10-13,10,1,1,1\n")
        with workspace() as ws:
            ws.data.mkdir()
            (ws.data / SNAPSHOTS).write_text(legacy, encoding="utf-8")
            ws.collect([("2026-10-14", 10, 1)], clones=[("2026-10-14", 1, 1)],
                       repository=stats(stars=5))
            ws.merge("2026-10-14T06:00:00")
            self.assertIsNone(traffic_history.alert(ws.data, "2026-12-20"))

    def test_alert_command_writes_the_issue_only_when_due(self) -> None:
        with workspace() as ws:
            self.history(ws, (50, 20, 1), (50, 20, 1))
            output = ws.raw.parent / "alert.json"
            arguments = ["alert", "--data", str(ws.data), "--milestone-due", "2026-12-20",
                         "--output", str(output)]
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(traffic_history.main(arguments), 0)
            self.assertFalse(output.exists())
            ws.collect([("2026-10-15", 500, 1)], clones=[("2026-10-15", 20, 1)])
            ws.merge("2026-10-15T06:00:00")
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(traffic_history.main(arguments), 0)
            self.assertEqual(json.loads(output.read_text(encoding="utf-8"))["title"],
                             "Traffic alert: 2026-10-15")


class FailureTests(unittest.TestCase):
    def assert_refused(self, ws: Workspace, message: str) -> None:
        before = ws.snapshot()
        with self.assertRaisesRegex(TrafficError, message):
            ws.merge("2026-10-15T06:00:00")
        self.assertEqual(ws.snapshot(), before)

    def test_invalid_or_partial_collection_leaves_history_unchanged(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 14))
            ws.merge("2026-10-14T06:00:00")
            ws.collect(window("2026-10-15", [(1, 1)] * 14), paths={"message": "Bad credentials"})
            self.assert_refused(ws, "paths: expected a list")
            ws.collect(window("2026-10-15", [(1, 1)] * 14))
            (ws.raw / "referrers.json").unlink()
            self.assert_refused(ws, "referrers.json: response missing")
            (ws.raw / "referrers.json").write_text("{not json", encoding="utf-8")
            self.assert_refused(ws, "unreadable response")

    def test_malformed_values_are_rejected(self) -> None:
        cases = [
            ([{"timestamp": "2026-10-14T05:00:00Z", "count": 1, "uniques": 1}], "UTC midnight"),
            ([{"timestamp": "2026-10-16T00:00:00Z", "count": 1, "uniques": 1}], "after"),
            ([{"timestamp": "2026-10-14T00:00:00Z", "count": -1, "uniques": 1}], "non-negative"),
            ([{"timestamp": "2026-10-14T00:00:00Z", "count": True, "uniques": 1}], "non-negative"),
            ([{"timestamp": "2026-10-14T00:00:00Z", "count": 1, "uniques": 1}] * 2, "duplicate"),
        ]
        for views, message in cases:
            with self.subTest(message=message), workspace() as ws:
                ws.collect(window("2026-10-14", [(1, 1)]))
                (ws.raw / "views.json").write_text(
                    json.dumps({"count": 1, "uniques": 1, "views": views}), encoding="utf-8")
                self.assert_refused(ws, message)

    def test_corrupt_history_file_is_never_overwritten(self) -> None:
        for content, message in [("date,views\n2026-10-01,1\n", "header"),
                                 (traffic_history.render(DAILY, [])
                                  + "2026-10-01,x,1,2026-10-01T06:00:00Z,1,1,2026-10-01T06:00:00Z\n", "non-numeric"),
                                 (traffic_history.render(DAILY, [])
                                  + "yesterday,1,1,2026-10-01T06:00:00Z,1,1,2026-10-01T06:00:00Z\n", "invalid date"),
                                 (traffic_history.render(DAILY, []) + "2026-10-01,1\n",
                                  "number of fields"),
                                 (traffic_history.render(DAILY, [])
                                  + "2026-10-01,1,1,t,1,1,2026-10-01T06:00:00Z\n", "invalid views_collected_at"),
                                 (traffic_history.render(DAILY, [])
                                  + "2026-10-01,1,,,1,1,2026-10-01T06:00:00Z\n", "non-numeric unique_views"),
                                 (traffic_history.render(DAILY, [])
                                  + "2026-10-01,1,1,2026-10-01T06:00:00Z,1,1,2026-10-01T06:00:00Z\n"
                                  + "2026-10-01,2,1,2026-10-02T06:00:00Z,1,1,2026-10-02T06:00:00Z\n",
                                  "duplicate date 2026-10-01")]:
            with self.subTest(message=message), workspace() as ws:
                ws.data.mkdir()
                (ws.data / DAILY).write_text(content, encoding="utf-8")
                ws.collect(window("2026-10-15", [(1, 1)] * 14))
                self.assert_refused(ws, message)

    def test_more_uniques_than_the_count_is_refused(self) -> None:
        cases = {
            "a daily entry": {"views.json": series("views", [("2026-10-15", 1, 9)])},
            "a window total": {"views.json": {"count": 1, "uniques": 9, "views": []}},
            "a referrer": {"referrers.json": [{"referrer": "x", "count": 1, "uniques": 9}]},
            "a path": {"paths.json": [{"path": "/o/r", "title": "r", "count": 1, "uniques": 9}]},
        }
        for case, responses in cases.items():
            with self.subTest(case=case), workspace() as ws:
                ws.collect(window("2026-10-14", [(1, 1)] * 14))
                ws.merge("2026-10-14T06:00:00")
                ws.collect(window("2026-10-15", [(1, 1)] * 14))
                for name, body in responses.items():
                    (ws.raw / name).write_text(json.dumps(body), encoding="utf-8")
                self.assert_refused(ws, "9 uniques exceed the count 1")

    def test_duplicate_ranks_in_rankings_are_refused(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 14))
            ws.merge("2026-10-14T06:00:00")
            path = ws.data / REFERRERS
            path.write_text(path.read_text(encoding="utf-8")
                            + path.read_text(encoding="utf-8").splitlines()[1] + "\n", encoding="utf-8")
            ws.collect(window("2026-10-15", [(1, 1)] * 14))
            self.assert_refused(ws, "duplicate snapshot_date, rank 2026-10-14 1")


def legacy_snapshot(collected_at: str, views: Sequence[Day],
                    clones: Sequence[Day] | None = None) -> dict[str, Any]:
    """A snapshot of the former layout: the four responses in one JSON file."""
    return {"schema_version": 1, "repository": "o/r", "collected_at_utc": collected_at,
            "source_commit_sha": "a" * 40,
            "views": series("views", views), "clones": series("clones", views if clones is None else clones),
            "referrers": [{"referrer": "github.com", "count": 3, "uniques": 1}],
            "paths": [{"path": "/o/r", "title": "r", "count": 3, "uniques": 1}]}


class LegacyImportTests(unittest.TestCase):
    def write(self, folder: Path, name: str, body: Any) -> None:
        folder.mkdir(exist_ok=True)
        (folder / name).write_text(json.dumps(body), encoding="utf-8")

    def test_replay_follows_collection_order_and_leaves_counts_blank(self) -> None:
        with workspace() as ws:
            legacy = ws.raw.parent / "snapshots"
            # The file names sort in the wrong order: the collection time decides.
            self.write(legacy, "b.json", legacy_snapshot(
                "2026-10-14T06:00:00.123456+00:00", window("2026-10-14", [(2, 1)] * 14)))
            self.write(legacy, "a.json", legacy_snapshot(
                "2026-10-15T06:00:00+00:00", window("2026-10-15", [(5, 2)] * 14)))
            stamps = traffic_history.import_legacy(legacy, ws.data)
            self.assertEqual(stamps, ["2026-10-14T06:00:00Z", "2026-10-15T06:00:00Z"])
            daily = ws.daily()
            self.assertEqual(daily["2026-10-01"], ("2", "2", "2026-10-14T06:00:00Z"))
            self.assertEqual(daily["2026-10-15"], ("5", "5", "2026-10-15T06:00:00Z"))
            snapshots = ws.rows(SNAPSHOTS)
            self.assertEqual([row["snapshot_date"] for row in snapshots], ["2026-10-14", "2026-10-15"])
            self.assertTrue(all(row[key] == "" for row in snapshots
                                for key in traffic_history.REPOSITORY_COUNTS))
            # 2026-10-01 from the first, the 14 later days from the second
            self.assertEqual(ws.badge("clones")["message"], "72")

    def test_a_new_collection_continues_a_replayed_history(self) -> None:
        with workspace() as ws:
            legacy = ws.raw.parent / "snapshots"
            self.write(legacy, "2026-10-14.json", legacy_snapshot(
                "2026-10-14T06:00:00+00:00", window("2026-10-14", [(2, 1)] * 14)))
            traffic_history.import_legacy(legacy, ws.data)
            ws.collect(window("2026-10-15", [(3, 1)] * 14), repository=stats(stars=2))
            ws.merge("2026-10-15T06:00:00")
            old, new = ws.rows(SNAPSHOTS)
            self.assertEqual((old["stars"], new["stars"]), ("", "2"))
            self.assertIsNone(traffic_history.alert(ws.data, "2026-11-22"))

    def test_an_invalid_legacy_snapshot_changes_nothing(self) -> None:
        with workspace() as ws:
            legacy = ws.raw.parent / "snapshots"
            self.write(legacy, "a.json", legacy_snapshot(
                "2026-10-14T06:00:00+00:00", window("2026-10-14", [(2, 1)] * 14)))
            broken = legacy_snapshot("2026-10-15T06:00:00+00:00", window("2026-10-15", [(2, 1)] * 14))
            broken["views"]["views"][0]["count"] = -1
            self.write(legacy, "b.json", broken)
            with self.assertRaises(TrafficError):
                traffic_history.import_legacy(legacy, ws.data)
            self.assertFalse(ws.data.exists())
            self.write(legacy, "b.json", {**broken, "collected_at_utc": "2026-10-15T06:00:00"})
            with self.assertRaisesRegex(TrafficError, "UTC offset"):
                traffic_history.import_legacy(legacy, ws.data)
            with self.assertRaisesRegex(TrafficError, "no legacy snapshot"):
                traffic_history.import_legacy(ws.raw, ws.data)

    def test_import_legacy_command_reports_and_refuses(self) -> None:
        with workspace() as ws:
            legacy = ws.raw.parent / "snapshots"
            self.write(legacy, "a.json", legacy_snapshot(
                "2026-10-14T06:00:00+00:00", window("2026-10-14", [(2, 1)] * 14)))
            arguments = ["import-legacy", "--snapshots", str(legacy), "--data", str(ws.data)]
            with contextlib.redirect_stdout(io.StringIO()) as stdout:
                self.assertEqual(traffic_history.main(arguments), 0)
            self.assertIn("imported 1 legacy snapshots", stdout.getvalue())
            with contextlib.redirect_stderr(io.StringIO()) as stderr:
                self.assertEqual(traffic_history.main(
                    ["import-legacy", "--snapshots", str(ws.raw), "--data", str(ws.data)]), 2)
            self.assertIn("the history was not changed", stderr.getvalue())


class CommandTests(unittest.TestCase):
    def test_init_creates_headers_and_keeps_existing_files(self) -> None:
        with workspace() as ws:
            self.assertEqual(traffic_history.initialize(ws.data), list(traffic_history.HEADERS))
            self.assertEqual(ws.rows(PATHS), [])
            (ws.data / DAILY).write_text(traffic_history.render(DAILY, [
                {"date": "2026-10-01", "views": "1", "unique_views": "1", "clones": "",
                 "unique_clones": "", "views_collected_at": "2026-10-01T06:00:00Z",
                 "clones_collected_at": ""}]),
                encoding="utf-8")
            self.assertEqual(traffic_history.initialize(ws.data), [])
            self.assertEqual(len(ws.rows(DAILY)), 1)

    def test_merge_command_reports_and_exits_two_on_refusal(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 14))
            summary = ws.raw.parent / "summary.md"
            arguments = ["merge", "--raw", str(ws.raw), "--data", str(ws.data),
                         "--collected-at", "2026-10-14T06:00:00Z", "--summary", str(summary)]
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(traffic_history.main(arguments), 0)
            self.assertIn("| `daily.csv` | 14 |", summary.read_text(encoding="utf-8"))
            (ws.raw / "clones.json").write_text("[]", encoding="utf-8")
            with contextlib.redirect_stderr(io.StringIO()) as stderr:
                self.assertEqual(traffic_history.main(arguments), 2)
            self.assertIn("the history was not changed", stderr.getvalue())

    def test_timestamps_need_an_offset(self) -> None:
        self.assertEqual(traffic_history.parse_time("2026-10-14T08:00:00+02:00"),
                         at("2026-10-14T06:00:00"))
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            traffic_history.main(["merge", "--raw", ".", "--data", ".",
                                  "--collected-at", "2026-10-14T06:00:00"])


if __name__ == "__main__":
    unittest.main()
