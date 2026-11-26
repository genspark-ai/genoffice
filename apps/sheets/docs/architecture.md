# Architecture

## Trust boundaries

```text
Renderer (React spreadsheet shell + Univer)
  -> typed preload bridge
  -> validated Electron IPC
  -> Rust XLSX sidecar
  -> temporary row-chunk index
  -> copy-on-write OOXML gateway

Renderer
  -> context extractor
  -> local privacy policy
  -> cloud planner
  -> untrusted command DSL
  -> local validation and dry-run
  -> user approval
  -> atomic commit and audit
```

The document core is the only workbook writer. The renderer cannot access disk, model credentials, or subprocesses. A cloud model cannot invoke native capabilities or commit files.

The Electron main process owns the XLSX sidecar lifecycle. The renderer receives an opaque session ID and can request only validated, size-limited worksheet ranges. External workbooks are read-only while this streaming path is being validated.

## Renderer composition

React owns the desktop title bar, Ribbon navigation, AI panel, async status, and file metadata. Univer owns the formula bar, worksheet canvas, sheet tabs, selection, scrolling, and zoom controls.

The AI panel is a collapsible peer of the worksheet rather than a permanent overlay. Collapsing it expands the worksheet column without recreating the Univer runtime or workbook session.

## Styles and visual objects

The sidecar resolves `styles.xml` cell formats and carries the source style index with each streamed cell. The renderer converts fonts, fills, alignment, and number formats into Univer cell style data.

Worksheet relationship files locate drawing parts without parsing worksheet bodies. Drawing anchors then resolve chart and image relationships. Charts use their OOXML cached categories and values and render as Apache-licensed Univer floating DOM components; this avoids the license-gated Univer Pro chart package. Embedded image data is loaded only when its anchored object is installed, through a validated 20MB media endpoint. Floating objects are reinstalled for the active sheet and viewport so off-screen drawings do not create an unbounded DOM.

## Large workbook reads

The sidecar reads ZIP metadata, workbook relationships, worksheet dimensions, and shared strings without constructing a JavaScript workbook snapshot. On the first range request for a sheet, it streams worksheet XML once into temporary 256-row chunks. Range requests wait only for the required chunk, and parsing continues in the background.

Univer starts with worksheet dimensions and empty sparse cell data. Scroll and active-sheet events request the visible range plus a buffer. Before another window is installed, the previous range and formats are cleared, so renderer memory does not grow with the total workbook size. Visual discovery reads only relationship, drawing, chart, style, and media metadata parts; it never builds a DOM for a large worksheet.

### Worksheet indexing tiers

| Tier | Worksheets                                                                              | Scheduling and cache policy                                                                                                                                          |
| ---- | --------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1    | The active worksheet                                                                    | Highest priority. Full indexing without the prefetch cache cap.                                                                                                      |
| 2    | The two nearest visible tabs and the active worksheet's transitive formula dependencies | Lower priority than Tier 1. Full indexing without the prefetch cache cap.                                                                                            |
| 3    | All other worksheets                                                                    | Best-effort prefetch within available worker slots and a shared 256 MiB row-chunk cache budget. A worksheet can remain unstarted or stop at its share of the budget. |

**A worksheet can remain partially indexed or unstarted for the entire user session.** The renderer requests full indexing only when the current operation needs that worksheet. A passive index read reports available data and completion without starting a worker or releasing its prefetch cap.

Live formula activation for a progressive workbook starts at the active worksheet and follows its statically resolvable formula references transitively. The renderer installs a closure only after indexing the required sheets and reading their precedent cells within the renderer's separate cell budget. Unrelated sheets can be opened and analyzed later. A sheet whose closure cannot be resolved or exceeds the renderer budget keeps cached values and can use the existing recalculation fallback when eligible.

The renderer loads row outlines for each grouped worksheet on its first visit, including the initial or restored tab. Each request follows that worksheet's index independently and can finish after the user switches tabs.

## Workbook state

Each document eventually owns:

1. The original XLSX package.
2. A normalized editor snapshot.
3. A revisioned operation journal.

The PoC implements the normalized snapshot and in-memory journal. The OOXML gateway operates independently to prove that untouched package entries survive a surgical cell edit.

## Adapter boundary

Product code depends on `WorkbookAdapter`, not Univer APIs. The adapter exposes:

- `getSnapshot`
- `plan`
- `apply`
- `undo`

This keeps AI planning, transaction safety, and audit behavior replaceable if the editor changes.

## XLSX preservation

The gateway:

- validates ZIP path safety and entry count;
- limits total uncompressed data;
- inventories each uncompressed entry by SHA-256;
- resolves a worksheet through workbook relationships;
- rewrites only the target worksheet or workbook metadata;
- refuses unknown sheet mappings and stale file hashes;
- writes through a temporary sibling file and atomic rename.

The current gateway does not claim full OOXML compatibility. Its contract is that unlisted entries retain identical uncompressed bytes.

## Production gaps

- Durable SQLite transaction and audit journal.
- Full XLSX-to-editor import model.
- High-fidelity borders, theme colors, conditional formatting, editable drawings and editable native charts.
- Tables, names, validation, merged cells, filters, and frozen panes.
- Formula calculation and differential Excel validation.
- Local privacy classifier and cloud model gateway.
- Signed updates, crash recovery, telemetry, and enterprise controls.
