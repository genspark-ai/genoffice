# First open of a large XLSX

## Baseline at `4649adf9`

Fixture `mytest/alldata-synthetic.xlsx`, 652 MiB compressed and 3.37 GiB of
worksheet XML. Four sheets contain about one million rows each. The compact
first viewport is served long before its sheet finishes indexing.

Local Linux Electron build, fresh user-data directories, file passed on the
command line. The OS file cache was left alone. Two frames after the opening
overlay disappeared, the grid showed `Record Number` in Sheet1!A1.

| Stage | Baseline |
| --- | ---: |
| Process launch to visible cells | 8.30–9.29 s |
| Process launch to `workbook:select` | 1.59–1.63 s |
| Session snapshot copy | 0.29–1.13 s |
| Sidecar `open` in Electron | 4.83 s |
| Sidecar response to visible cells | 1.59–1.71 s |
| Standalone sidecar `open` | 4.44 s |
| Standalone first 100 rows, 1,800 cells | 8 ms |

The snapshot SHA-256 took about 1.02 s and finished while `open` was
running. A separately instrumented release sidecar spent 3.97 s waiting
for the four metadata scans. The slowest scan spent 1.00 s reading and
inflating its ZIP entry and roughly 2.97 s walking XML. The scan of the
first sheet finished at 3.80 s. A subsequent cell-count pass over the first
two million stored cells took 0.90 s. The Electron renderer CPU profile
attributed about 1.20 s to Univer's row-height cache, including 0.71 s
inside the table-filter lookup.

## Parallel metadata and cell caching

The first progressive implementation displayed cells after 5.85 s and
kept the active sheet read-only until its full index finished at 18.11 s.
The current implementation scans the worksheet metadata independently
from the cell cache. It starts the active sheet first and uses bounded
background workers to scan and write the other sheets' caches.

Two fresh Electron launches with the same synthetic workbook displayed
the first cells in 4.14–4.23 s. Sheet1 was already editable in both
runs. Before switching sheets, all four had 279–329 chunks each on disk.
Switching to Sheet2, Sheet3 and Sheet4 took 45–99 ms each until their
first cells were visible and editable.

A third launch with a faster snapshot copy showed cells in 3.23 s and
accepted an edit at 3.25 s. After switching to the other three tabs, the
first edits succeeded in 96–130 ms. The source XLSX was left untouched.

In a separate release-sidecar run, metadata for all four sheets was
available after 0.81–0.93 s. Their full indexes finished after 14.55–15.76 s.
That run produced 15,313 cache chunks. Electron verification also covered
a saved non-first active tab, a cached cross-sheet formula, a trailing merge,
and a genuinely protected cell.
