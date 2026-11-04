# 📈 Repository traffic history

[![Purpose: analytics](https://img.shields.io/badge/purpose-analytics-1D76DB)](../README.md)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](../LICENSE)

## 🎯 What is collected

**Export repository traffic** keeps this repository's GitHub traffic beyond
the API's 14-day window. Once a day it collects views, clones, the top
referrers and popular paths, and the repository's stars, forks, watchers and
open issues; merges them into the CSV files of the data-only
[traffic-history branch](https://github.com/danielep71/K-FRTB-SA/tree/traffic-history);
writes the README traffic badges there; and opens an issue when traffic
stands out (*Traffic alert* below). It is repository analytics, not CI.
No visitor-level identities are collected.

## 🗄️ Stored data

The branch holds `README.md`, `LICENSE` and these files under `data/`. It is
never merged into main or a release branch, and its history is not
rewritten.

| File | Key | Contents |
| --- | --- | --- |
| `daily.csv` | `date` | Views and unique views, clones and unique clones per UTC day. `views_collected_at` and `clones_collected_at` give the collection that last changed each pair. |
| `snapshots.csv` | `snapshot_date` | One row per collection day: the first and last day the API reported, the 14-day totals, the repository's `stars`, `forks`, `watchers` and `open_issues` at collection time, and `collected_at`, the collection that last changed that day's snapshot, referrers or paths. |
| `referrers.csv` | `snapshot_date`, `rank` | The top 10 referrers over the 14 days before each collection. |
| `paths.csv` | `snapshot_date`, `rank` | The top 10 repository paths over the same window. |
| `badges/*.json` | — | The README badges, rewritten from the stored data after each merge (*Badges* below). |
| `legacy/` | — | The files of the former layout, kept as an archive: the old `daily.csv`, `latest.json` and one JSON snapshot per collection day from 2026-10-18 to 2026-10-31. |

Referrers and paths are kept per collection day, because the API gives them
only as 14-day totals, never per day.

The repository counts come from the repository API (`stargazers_count`,
`forks_count`, `subscribers_count`, `open_issues_count`). `open_issues` is
GitHub's count, which includes open pull requests. Snapshots replayed from
`legacy/` have them blank: a blank count was not recorded, not zero.

### Badges

The README shows three badges read by shields.io from `data/badges/`, in its
*endpoint* format:

| Badge | Counts |
| --- | --- |
| `clones` | Every clone in `daily.csv`, summed since tracking began. |
| `visitors (14d)` | Unique visitors over the API's 14-day window, from the latest snapshot. Daily unique counts cannot be summed into a unique count. |
| `visits` | Daily unique visitors summed since tracking began: a visitor counts once per day they visit. |

Large numbers are shortened (`1.2k`, `11k`). A badge file changes only when
its value does, so a rerun with the same data changes nothing.

### Merge rules

`tools/traffic_history.py` merges each collection idempotently:

- **Overlapping windows:** for every date the collection reports, its values
  replace the stored ones. The newest collection wins, so a day that was
  incomplete when first collected is corrected later. Views and clones are
  merged separately, each with its own collection time.
- **Older history:** dates the collection does not report keep their values.
  Nothing is deleted.
- **No double counting:** values are replaced, never added. Unique counts are
  never summed across days or windows; a 14-day unique count is only in
  `snapshots.csv`.
- **Same-day runs:** a collection day's snapshot, referrers and paths are one
  unit. A second collection on the same UTC day replaces the unit only if any
  of the three changed.
- **Order:** "newest" means the latest collection time. Merging an older
  collection after a newer one that changed the values never replaces them;
  it can only fill a date or metric that nothing newer reported. A stored time
  is that of the collection that last *changed* the values, so that a rerun
  with the same data changes nothing. An older collection replayed after a
  newer but unchanged one can therefore replace its values: replay saved
  collections in collection order, as `import-legacy` does. The workflow
  merges each collection once, in order.

### Zero or unavailable

The API reports each day of its window, days without activity included, as
`0`. So **`0`** means the day was collected and had no activity; **no row**
means no collection covered the day; **a blank metric** means one of the two
series was not reported for that day. The step summary of each run lists the
dates between the first and the last stored day that have no row.

The branch is public, as the repository is. The data is aggregate (counts,
referrers and paths) and identifies no visitor. The counts include my own
visits and every clone made by CI checkouts. Analytics describe repository
activity, not user adoption or calculation quality.

## 🚨 Traffic alert

After each successful collection the **Traffic alert** job compares the
latest snapshot with the one before (`traffic_history.py alert`) and opens
one issue, titled `Traffic alert: <snapshot date>`, when any of these holds:

- **Views or clones spike:** the 14-day total at least doubles and grows by
  at least 100 views or 10 clones;
- **Stars or forks increase;**
- **New referrer:** a referrer never stored before brings at least 5 views.

A count blank in the older snapshot raises nothing. At most one issue is
opened per snapshot date: the job skips a date whose issue exists, open or
closed. Like every issue it is assigned to me, labelled `P3` and
`repository`, and filed under the open milestone with the earliest due date.
Its target completion date is seven days after the snapshot, or the
milestone's due date if that is sooner, so it passes *Issue triage*.

The job uses the workflow token with `issues: write` and `contents: read`. It
has no environment, so it never receives the App key or token.

## ⚙️ Workflows and trust boundaries

- **Export repository traffic** runs daily at 04:37 UTC and supports manual
  dispatch. Collection runs only on `main`, using the `analytics` environment.
  Default-branch schedules may be delayed or disabled by GitHub after extended
  repository inactivity; inspect Actions rather than assuming collection ran.
- **Validate traffic tooling** runs the offline tests on relevant PRs and
  pushes to main and release branches, with read-only contents access and no
  traffic credential.
- A short-lived GitHub App installation token reads the five API responses in
  a separate curl-only step. The merge step reads local files only and
  receives no token. The publication step receives only the built-in
  `GITHUB_TOKEN`, with `contents: write`. Checkout credentials are not
  persisted.
- Every response and every history file validates before anything is
  written. `tools/traffic_publish.py` publishes the changed files as one
  commit on the checked-out parent, with a non-forced ref update. If the
  branch moved in between, the run fails rather than overwrite another
  commit; the next run collects the same 14 days.
- Protect `traffic-history` against deletion and force pushes without requiring
  PRs for its ordinary automated data commits. Do not execute workflow code
  from the data branch.

## 🔑 GitHub App configuration

I configure the App and credentials before approving the migration merge:

1. Create a dedicated traffic GitHub App, disable its webhook and grant only
   **Administration: Read-only** (GitHub adds Metadata: Read-only).
2. Install it on this repository. Generate a dedicated private key.
3. Keep the `analytics` environment restricted to `main`. Add the App Client ID
   as variable **TRAFFIC_APP_CLIENT_ID** and the complete PEM key as secret
   **TRAFFIC_APP_PRIVATE_KEY** in that environment, not at repository scope.
4. Confirm configuration before merging the migration. After integration into
   the default branch, dispatch **Export repository traffic** on `main` and
   verify a successful run and a new `traffic-history` commit.
5. Only after that successful run, I delete the old `TRAFFIC_TOKEN` environment
   secret and revoke the personal token. There is no fallback to that token.

Never put the private key in source, logs, issues, screenshots or chat.
The configuration check sees only whether the key exists. Only the pinned
App-token action receives its value; only the curl API step receives the
installation token. The action scopes the token to this repository with
Administration: read, expires it after an hour and revokes it at job end.
Publication retains the built-in `GITHUB_TOKEN` with contents: write.

### Key rotation

Generate a new dedicated key, replace the `analytics` environment secret,
then dispatch on main and verify the resulting history commit. Only after
success, delete the old key in the App settings, checking its fingerprint.
Rotate promptly after suspected exposure. Missed data outside GitHub's
14-day window cannot be recovered.

### Failure diagnosis

| Failure | Recovery |
| --- | --- |
| App client ID or private key not set | Configure both values in `analytics`; check environment access. |
| App token creation fails | Verify Client ID, key and installation on this repository. |
| curl HTTP 401 | Check installation and key, then rerun. |
| curl HTTP 403/404 | Verify Administration: read and repository selection; accept permission changes. |
| curl HTTP 429/5xx or timeout | Retry later; curl retries transient errors before failing. |
| Missing, malformed or oversized response | Inspect the failing endpoint; no history is published. |
| *Merge*: `ERROR: ...; the history was not changed` | Inspect the response or the named file; a file edited by hand must keep its header and columns. |
| *Check out the traffic history* fails | The branch was deleted: restore it from the last known commit; never recreate it empty over existing history. |
| *Publish*: `traffic-history moved since the checkout` or non-fast-forward | Check branch rules and competing updates, then rerun; never force. |
| *Traffic alert*: `No open milestone with a due date` | Open the next milestone with a due date, then rerun; the history is already saved. |
| *Traffic alert*: `gh issue create` fails | Restore the label, the assignee or the milestone, then rerun. |

## 🧪 Validation and operations

Python 3.12 on the workflow runner uses only the standard library.

```sh
python3 -m unittest discover -s tests/tooling -p "test_traffic_*.py" -v
```

`test_traffic_history.py` covers the merge rules, the repository counts, the
badges, the alert thresholds, the legacy replay and the refusal of invalid
input. `test_traffic_publish.py` covers the changed-file selection, the
atomic commit on the expected parent, the moved-branch and non-fast-forward
failures, the restriction to main, and the isolation of the App credentials
between workflow steps.

To repair the branch, add a new commit made with
`python3 tools/traffic_history.py merge` or by hand, keeping every header.
`python3 tools/traffic_history.py init --data data` recreates missing files
with their headers and never overwrites an existing one. Never reset the
branch to make a failed publication pass.

### Replay of the former layout

The former layout kept `data/daily.csv` with other columns, `latest.json` and
one JSON snapshot per day. Right after the main integration of #185 and
before the first scheduled run, I move those files to `data/legacy/` and
replay them:

```sh
python3 tools/traffic_history.py import-legacy --snapshots data/legacy/snapshots --data data
```

Every daily count of the replay must equal the old `daily.csv`, then the
result is pushed as one commit. Offline, the replay of the 14 snapshots of
2026-10-18 to 2026-10-31 reproduces all 26 days.

This is operational analytics. Static tests, traffic counts, and successful
collection do not run Excel or validate capital calculations.

## 📋 Live configuration record

Verified on October 18, 2026 (Europe/Rome): the data-only branch exists;
[Protect traffic history](https://github.com/danielep71/K-FRTB-SA/rules/25321321)
is active for its exact ref, blocks deletion and non-fast-forward updates,
and has no bypass actors. The `analytics` environment permits the main branch
only (one branch, zero tags) and holds the `TRAFFIC_TOKEN` environment
secret. The first real export succeeded in
[run 39105087826](https://github.com/danielep71/K-FRTB-SA/actions/runs/39105087826)
and published data commit `af6a68a` on `traffic-history`; issue #23 is closed.
This dated record describes the former personal-token setup.

Updated on October 31, 2026 (Europe/Rome): the GitHub App migration (#181)
is in operation. The main run
[#17](https://github.com/danielep71/K-FRTB-SA/actions/runs/41216703015)
passed with the App token and published data commit `3cc986c`. After that
run I deleted the `TRAFFIC_TOKEN` secret from `analytics` and revoked the
personal token; the workflow has no fallback to it. This record is my
confirmation; environment secrets cannot be read back through the public API.
