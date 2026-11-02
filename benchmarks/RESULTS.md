# BTA OSG cache and mirror benchmark

Measured on 2026-11-02 against the real `EmbeddedDB` cache and `DesktopDataCoordinator`
mirror code. The benchmark does not simulate those components; it calls them directly.

## Environment

| Item | Value |
|---|---|
| OS | Microsoft Windows 10.0.26100 |
| Runtime | .NET 10.0.11 |
| CPU | Intel64 Family 6 Model 151 Stepping 5, 12 logical processors |
| SQL | Disabled for this run (offline cache benchmark) |
| Build | Release |

## Method

- Data set: N documents, 2N routing logs, N storage movements, N/2 directives, 1,000 audit
  rows (the audit mirror cap), and one heartbeat row.
- Every operation is preceded by a warm-up and reported as the median of 3 iterations.
- Included: cache save, cache load, the per-tick snapshot merge, and the local-time
  normalization added in the 2026-11-02 hardening batch.
- Excluded: SQL read and network time, UI rendering, and grid refresh. Those are
  environment-specific and should be measured on the office network.
- A repeat run varied by up to about 15 percent on the smallest case (cold disk and
  antivirus warm-up). Treat the numbers as indicative, not exact.

## Results

### 5,000 documents, 10,000 routing logs, 5,000 movements, 2,500 directives, 1,000 audit rows

| Metric | Median |
|---|---|
| Cache save (full XML, temp + replace) | 222.30 ms |
| Cache load (schema + XML + identity resync) | 262.61 ms |
| Quiet mirror merge, all tables | 41.38 ms |
| One-row change merge, Documents | 13.68 ms |
| Local-time normalization, 4 log tables | 26.15 ms |
| Cache file | 12.71 MB |
| Previous generation (.bak) | 12.71 MB |
| Managed heap after load | 17.17 MB |

### 20,000 documents, 40,000 routing logs, 20,000 movements, 10,000 directives, 1,000 audit rows

| Metric | Median |
|---|---|
| Cache save (full XML, temp + replace) | 358.59 ms |
| Cache load (schema + XML + identity resync) | 403.59 ms |
| Quiet mirror merge, all tables | 77.73 ms |
| One-row change merge, Documents | 23.68 ms |
| Local-time normalization, 4 log tables | 64.88 ms |
| Cache file | 50.01 MB |
| Previous generation (.bak) | 50.01 MB |
| Managed heap after load | 65.23 MB |

## What the numbers mean

**The office scale.** The thesis estimates about 3,000 documents per quarter. The 5,000
row case is roughly six months of documents plus their logs. The 20,000 row case is
roughly twenty months of documents with about two years of logs. The application is
comfortable at 5,000 and still usable at 20,000.

**Saving.** A full cache save costs about 222 ms at 5,000 rows and 359 ms at 20,000 rows.
Routine mutations use the 1.5 second deferred flush, so this work happens on a background
timer and does not block the UI. Explicit saves do pay it: application close, a void, and
the per-document save during offline replay. The replay path is the one to watch: at
20,000 rows, replaying 50 offline documents costs about 18 seconds of cache serialization
on the sync thread. It does not freeze the UI, but it is slow enough to notice. If replay
volume grows, batch the replay save every N documents while preserving the existing
no-double-mint rule.

**Merging.** The snapshot merge is the UI-thread cost of a sync tick: about 41 ms at 5,000
rows and 78 ms at 20,000 rows. With a 5 to 30 second sync interval, this is a short,
periodic hitch, not a sustained load. The one-row change merge costs almost the same as a
quiet merge (14 ms and 24 ms), which shows the cost is the full scan, not the changed row.

**Local-time normalization.** Converting the four mirrored log tables to local time adds
about 26 ms at 5,000 rows and 65 ms at 20,000 rows. For the regular sync tick this runs on
the background thread. After a connected write, the post-write refresh runs on the UI
thread, so the worst case at 20,000 rows is roughly the 78 ms merge plus the 65 ms
normalization, plus the grid refresh. Under about 150 ms total, it remains acceptable.

**Disk and memory.** The cache is 12.71 MB at 5,000 rows and 50.01 MB at 20,000 rows. The
atomic save keeps one previous generation, so steady-state disk use is about twice the
cache size. During a save, a temporary generation exists as well, so peak disk use is about
three times the cache size: roughly 150 MB at 20,000 rows. The loaded cache occupies about
17 MB and 65 MB of managed heap respectively.

**Scaling ceiling.** The existing full-pull design reads every document and every child log
on each tick. These measurements show that the client merge and normalization are still
small at 20,000 documents, but they grow linearly. When the log tables reach roughly
20,000 to 50,000 rows, the UI hitch and the serialization time become worth addressing. The
next step is the incremental pull already described in the code: pull by ID above the
last seen value and merge append-only.

## Thesis-ready summary

At the expected first-year office volume, the desktop cache costs about 220 to 260 ms per
full save, 260 to 300 ms at startup, 40 ms per sync-tick merge, and 26 ms for local-time
normalization. At a two-year volume of 20,000 documents and about 70,000 child rows, those
figures rise to about 360 ms per save, 400 ms at startup, 78 ms per merge, and 65 ms for
normalization. The dominant costs scale with total rows because the current design mirrors
full snapshots; the client-side work is still under about 150 ms per tick at that volume,
and the system remains usable. The full-pull design is the scaling item to replace before
the log tables reach roughly 20,000 to 50,000 rows.

## Reproduce

```powershell
dotnet run --project benchmarks/BTA_OSG_Benchmarks/BTA_OSG_Benchmarks.vbproj -c Release -- 5000 20000
```
