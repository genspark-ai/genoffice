# Structural changes guide (structure)

## Operation definitions

- `{op:"insert_rows", sheetId, row:2, count:1}` — row is 1-based; new rows are inserted **before** that row.
- `{op:"delete_rows", sheetId, row, count}` — deletes count rows starting at row.
- `{op:"insert_cols", sheetId, column:"C", count}` — new columns are inserted before that column; column uses the column letter.
- `{op:"delete_cols", sheetId, column, count}`
- `{op:"add_sheet", name, rows?, columns?}` — creates a blank worksheet (no sheetId needed). The grid defaults to **1000 rows × 20 columns**; writes beyond the grid are rejected and formula spills truncate at its edge, so pass `rows`/`columns` sized to the expected data (plus a little headroom) when it will hold more — e.g. `{op:"add_sheet", name:"ja", rows:6700}` before copying ~6600 rows into it.
- `{op:"delete_sheet", sheetId}` — deletes a worksheet; the last remaining sheet cannot be deleted. Formulas on other sheets referencing the deleted sheet become #REF! (listed in the preview warnings); for real xlsx files there is an additional fail-closed guard on save (sheets referenced by formulas/charts/defined names refuse deletion).
- `{op:"duplicate_sheet", sheetId, name?}` — copies a whole worksheet (contents and formats); name is auto-generated when omitted. Unavailable in large-file streaming mode; sheets containing pivot tables cannot be duplicated.
- `{op:"set_sheet_hidden", sheetId, hidden:true|false}` — hides/shows a worksheet; at least one sheet must stay visible.
- `{op:"move_sheet", sheetId, position:1}` — moves to the 1-based tab position.

## Sheet protection

`{op:"protect_sheet", sheetId, protected:true|false}` — **layout-class**, can share a batch with content/format operations (exempt from the batching discipline below). Written into the file on save (passwords not supported); password-protected sheets cannot be unprotected. The editor itself does not enforce the lock.

## Mandatory batching discipline

Structural operations move cell addresses and **cannot appear in the same batch as content/format/sort-layout operations**. The correct rhythm:

1. Submit the structural changes alone first (multiple structural ops may share a batch, applied in order, effective on submit)
2. Call get_workbook_context / read_range again for the shifted layout
3. Then submit the content/format changes

## Formula references

- On row/column insertion/deletion, all A1 references in the workbook's formulas are rewritten automatically for the shift (including absolute $ references and cross-sheet prefixed references).
- Formulas whose reference target is deleted entirely become **#REF!** — the preview warnings list the affected cells precisely. When this happens, ask the user to confirm, or proactively suggest a fix.
- Range references (like SUM(B2:B10)) shrink automatically on partial deletion — you don't need to repair formulas by hand.

## Common mistakes

- ❌ insert_rows followed by writing into the "new rows" in the same batch — rejected by the mixed-batch rule, and the new rows' addresses don't exist before apply.
- ❌ Continuing to write with old addresses right after a structural change applies — re-read the context first.
- ❌ Using delete_rows to clear data — when you only want to clear contents, use clear_range (structure unchanged, formula references unaffected).
