/// Pivot layout growth: when a refresh sends an expanded output area (new row
/// or column items entered the source data), the pivotTableDefinition's
/// <location ref="…"> must be widened to match, and the new cells must be
/// clear of pre-existing non-pivot content.  This guard ensures the file is
/// always consistent on disk — an unupdated ref would confuse Excel about
/// which cells belong to the pivot.

import { columnIndex, columnLabel } from '../domain/cell-address'
import type { MutablePackage } from './xlsx-drawing-add'
import { buildCacheDefinitionXml, buildPivotTableXml, type PivotAddition } from './xlsx-pivot-add'

export class PivotExpandError extends Error {}

/// Carries the intended new output area for one recomputed pivot.
export interface PivotRefreshUpdate {
  /// xl/pivotCache/pivotCacheDefinitionN.xml
  readonly cachePath: string
  /// xl/worksheets/sheetN.xml — the sheet that hosts the pivot output.
  /// May be omitted when the renderer only knows the sheet name; then
  /// planCellEditsToXlsx resolves it from sheetName before calling
  /// applyPivotLayoutExpansions.
  readonly worksheetPath?: string | undefined
  /// Sheet name used when worksheetPath is absent (one of the two is required).
  readonly sheetName?: string | undefined
  /// New full output ref (including header row + grand-total row/col),
  /// e.g. "F1:G7".  May equal the current ref (no-op), larger, or (after a
  /// layout edit) smaller.
  readonly newOutputRef: string
  /// Layout edit (A3): the full edited layout in the pivot-addition shape.
  /// The pivotTableDefinition and pivotCacheDefinition are regenerated from
  /// it — the original name, cacheId, and records rel id stay; the emptied
  /// records part is rebuilt by Excel via refreshOnLoad. The payload's own
  /// `name` is a renderer placeholder and is ignored.
  readonly relayout?: Omit<PivotAddition, 'worksheetPath'> | undefined
}

// ─── Reference helpers ──────────────────────────────────────────────────────

interface CellRef {
  col: number // 0-based column index
  row: number // 0-based row index
}

interface AreaRef {
  start: CellRef
  end: CellRef
}

function parseCellRef(ref: string): CellRef {
  const m = /^([A-Z]+)(\d+)$/.exec(ref.trim().toUpperCase())
  if (!m) throw new PivotExpandError(`Cannot parse cell address "${ref}".`)
  return { col: columnIndex(m[1]!), row: Number(m[2]!) - 1 }
}

function parseAreaRef(ref: string): AreaRef {
  const parts = ref.trim().toUpperCase().split(':')
  if (parts.length !== 2) throw new PivotExpandError(`Cannot parse area ref "${ref}".`)
  return { start: parseCellRef(parts[0]!), end: parseCellRef(parts[1]!) }
}

function areaRefToString(area: AreaRef): string {
  return (
    `${columnLabel(area.start.col)}${area.start.row + 1}` +
    `:${columnLabel(area.end.col)}${area.end.row + 1}`
  )
}

// ─── Package helpers ─────────────────────────────────────────────────────────

/// Find the pivotTable part path linked to a given pivotCacheDefinition via
/// its own rels (pivot table → cache, so we do a reverse scan).
async function findPivotTablePathForCache(
  pkg: MutablePackage,
  cachePath: string,
): Promise<string | null> {
  const allPaths = await pkg.paths()
  const tableRelsPattern = /^xl\/pivotTables\/_rels\/[^/]+\.rels$/
  for (const relsPath of allPaths) {
    if (!tableRelsPattern.test(relsPath)) continue
    const relsXml = await pkg.readText(relsPath)
    // Each pivot table rels file has exactly one pivotCacheDefinition relationship
    const targetMatch =
      /Target="([^"]+)"[^>]*Type="[^"]*pivotCacheDefinition[^"]*"/.exec(relsXml) ??
      /Type="[^"]*pivotCacheDefinition[^"]*"[^>]*Target="([^"]+)"/.exec(relsXml)
    if (!targetMatch?.[1]) continue
    // Resolve the target relative to the rels file's directory
    // rels path: xl/pivotTables/_rels/pivotTable1.xml.rels
    // Target: ../pivotCache/pivotCacheDefinition1.xml
    // Resolved: xl/pivotCache/pivotCacheDefinition1.xml
    const relsDir = relsPath.replace(/\/_rels\/[^/]+$/, '')
    const resolved = resolveRelativePath(relsDir, targetMatch[1])
    if (resolved !== cachePath) continue
    // The pivot table path: strip _rels/ and .rels suffix
    const tableFilename = relsPath.replace(/.*\/_rels\//, '').replace(/\.rels$/, '')
    return `${relsDir}/${tableFilename}`
  }
  return null
}

/// Minimal relative-path resolver: walk ".." segments from the base directory.
function resolveRelativePath(baseDir: string, target: string): string {
  const segments = baseDir.split('/')
  for (const part of target.split('/')) {
    if (part === '..') segments.pop()
    else if (part !== '.') segments.push(part)
  }
  return segments.join('/')
}

/// Read the current <location ref="…"> from a pivot table XML.
function extractLocationRef(pivotTableXml: string): string | null {
  return /<location\b[^>]*\bref="([^"]+)"/.exec(pivotTableXml)?.[1] ?? null
}

/// Replace the ref attribute on the <location …> element.
function updateLocationRef(pivotTableXml: string, newRef: string): string {
  return pivotTableXml.replace(
    /<location\b([^>]*)\bref="[^"]*"([^>]*)>/,
    (_, before, after) => `<location${before}ref="${newRef}"${after}>`,
  )
}

// ─── Conflict detection ───────────────────────────────────────────────────────

/// True when the worksheet XML has any non-empty cell in the given area.
function worksheetHasContentInArea(worksheetXml: string, area: AreaRef): boolean {
  // Walk every <row r="N"> in the range, then every <c r="XX"> in the column band.
  const { start, end } = area
  const rowPattern = /<row\b([^>]*)>([\s\S]*?)<\/row>/g
  for (const rowMatch of worksheetXml.matchAll(rowPattern)) {
    const rAttr = /\br="(\d+)"/.exec(rowMatch[1] ?? '')
    if (!rAttr) continue
    const rowIdx = Number(rAttr[1]) - 1 // 0-based
    if (rowIdx < start.row || rowIdx > end.row) continue
    // Check each cell in this row
    const cellPattern = /<c\b([^>]*)(?:\/>|>[\s\S]*?<\/c>)/g
    for (const cellMatch of rowMatch[2]!.matchAll(cellPattern)) {
      const rCell = /\br="([A-Z]+\d+)"/.exec(cellMatch[1] ?? '')
      if (!rCell) continue
      const colMatch = /^([A-Z]+)/.exec(rCell[1]!)
      if (!colMatch) continue
      const colIdx = columnIndex(colMatch[1]!)
      if (colIdx < start.col || colIdx > end.col) continue
      // Cell exists in area — is it non-empty?
      const cellXml = cellMatch[0]!
      // Self-closing <c .../> or <c with no <v> = empty
      if (cellXml.endsWith('/>')) continue
      if (!/<v[^>]*>/.test(cellXml) && !/<is>/.test(cellXml)) continue
      return true
    }
  }
  return false
}

/// Subtract area B from area A, returning the set of sub-areas in A that are
/// not covered by B. Used to identify the newly-added cells when the pivot
/// grows right or down.
function subtractAreas(outer: AreaRef, inner: AreaRef): AreaRef[] {
  // Only handles the case where inner is fully inside outer (which it always
  // is here: old ⊆ new).
  const result: AreaRef[] = []
  if (
    inner.start.col > outer.start.col ||
    inner.end.col < outer.end.col ||
    inner.start.row > outer.start.row ||
    inner.end.row < outer.end.row
  ) {
    // Generic subtraction: rows below old area
    if (inner.end.row < outer.end.row) {
      result.push({
        start: { col: outer.start.col, row: inner.end.row + 1 },
        end: { col: outer.end.col, row: outer.end.row },
      })
    }
    // Columns to the right of old area (in rows covered by old)
    if (inner.end.col < outer.end.col) {
      result.push({
        start: { col: inner.end.col + 1, row: outer.start.row },
        end: { col: outer.end.col, row: Math.min(inner.end.row, outer.end.row) },
      })
    }
  }
  return result
}

// ─── Main export ─────────────────────────────────────────────────────────────

/// For each PivotRefreshUpdate whose newOutputRef is larger than the current
/// location ref, this function:
///  1. Checks that the newly-added rows/columns in the *original* worksheet
///     XML are empty (fail-closed on conflict).
///  2. Updates the pivotTableDefinition <location ref="…"> to newOutputRef.
///
/// No-ops (same ref) are silently skipped. Shrinking refs are also silently
/// skipped (the frontend never sends a smaller ref intentionally, but it is
/// not an error).
///
/// Called from planCellEditsToXlsx after structural shifts but before cell
/// edits are applied, so the worksheet XML reflects the pre-edit state and
/// the conflict check sees pre-existing content.
export async function applyPivotLayoutExpansions(
  pkg: MutablePackage,
  updates: readonly PivotRefreshUpdate[],
  touchedEntries: Set<string>,
): Promise<void> {
  for (const update of updates) {
    const pivotTablePath = await findPivotTablePathForCache(pkg, update.cachePath)
    if (!pivotTablePath) {
      // Cache not yet linked to a table (edge case: bare cache from old save) — skip.
      continue
    }

    const pivotTableXml = await pkg.readText(pivotTablePath)
    const currentRef = extractLocationRef(pivotTableXml)
    if (!currentRef) continue

    const oldArea = parseAreaRef(currentRef)
    const newArea = parseAreaRef(update.newOutputRef)
    const sameArea =
      oldArea.start.col === newArea.start.col &&
      oldArea.start.row === newArea.start.row &&
      oldArea.end.col === newArea.end.col &&
      oldArea.end.row === newArea.end.row

    // Same size and no layout edit: no-op.
    if (sameArea && update.relayout === undefined) continue

    // Moved origin (shouldn't happen on refresh, but fail safe).
    if (oldArea.start.col !== newArea.start.col || oldArea.start.row !== newArea.start.row) {
      throw new PivotExpandError(
        `Pivot "${update.newOutputRef}" cannot change its top-left cell during refresh.`,
      )
    }

    const grows = newArea.end.row > oldArea.end.row || newArea.end.col > oldArea.end.col
    if (grows) {
      if (update.worksheetPath === undefined) {
        throw new PivotExpandError('Pivot expansion needs a resolved worksheet path.')
      }
      // Growth — check added cells for pre-existing content. Shrinking needs
      // no check: the renderer already cleared the leftover cells and saved
      // them as ordinary cell edits.
      const worksheetXml = await pkg.readText(update.worksheetPath)
      const addedAreas = subtractAreas(newArea, oldArea)
      for (const added of addedAreas) {
        if (worksheetHasContentInArea(worksheetXml, added)) {
          const addedRef = areaRefToString(added)
          throw new PivotExpandError(
            `Pivot expansion into ${addedRef} conflicts with existing worksheet content. ` +
              'Clear the area first or move the pivot.',
          )
        }
      }
    }

    if (update.relayout === undefined) {
      const updatedTableXml = updateLocationRef(pivotTableXml, update.newOutputRef)
      pkg.write(pivotTablePath, updatedTableXml)
      touchedEntries.add(pivotTablePath)
      continue
    }
    await applyPivotRelayout(pkg, update, pivotTablePath, pivotTableXml, touchedEntries)
  }
}

/// Layout edit: regenerate the pivotTableDefinition and pivotCacheDefinition
/// from the edited layout, keeping the original name, cacheId, and records
/// rel id. The records part empties (recordCount 0) — refreshOnLoad, set via
/// the same save's pivotCacheRefreshPaths, has Excel rebuild it on open.
async function applyPivotRelayout(
  pkg: MutablePackage,
  update: PivotRefreshUpdate,
  pivotTablePath: string,
  pivotTableXml: string,
  touchedEntries: Set<string>,
): Promise<void> {
  const relayout = update.relayout!
  const name = unescapeAttribute(
    /<pivotTableDefinition\b[^>]*?\bname="([^"]*)"/.exec(pivotTableXml)?.[1] ?? '',
  )
  const cacheId = Number(/\bcacheId="(\d+)"/.exec(pivotTableXml)?.[1])
  if (!name || !Number.isInteger(cacheId)) {
    throw new PivotExpandError('The pivot table definition is missing its name or cacheId.')
  }
  const addition: PivotAddition = {
    ...relayout,
    name,
    worksheetPath: update.worksheetPath ?? '',
    location: refToArea(update.newOutputRef),
  }
  pkg.write(pivotTablePath, buildPivotTableXml(cacheId, addition))
  touchedEntries.add(pivotTablePath)

  const cacheXml = await pkg.readText(update.cachePath)
  const recordsRelId = /<pivotCacheDefinition\b[^>]*?\br:id="([^"]+)"/.exec(cacheXml)?.[1]
  if (!recordsRelId) {
    throw new PivotExpandError('The pivot cache definition is missing its records relationship.')
  }
  pkg.write(update.cachePath, buildCacheDefinitionXml(recordsRelId, addition, 0))
  touchedEntries.add(update.cachePath)

  // Empty the records part (found through the cache's rels); Excel rebuilds
  // it from the source on open.
  const cacheRelsPath = update.cachePath.replace(/\/([^/]+)$/, '/_rels/$1.rels')
  if (await pkg.has(cacheRelsPath)) {
    const relsXml = await pkg.readText(cacheRelsPath)
    const target =
      /Target="([^"]+)"[^>]*Type="[^"]*pivotCacheRecords[^"]*"/.exec(relsXml)?.[1] ??
      /Type="[^"]*pivotCacheRecords[^"]*"[^>]*Target="([^"]+)"/.exec(relsXml)?.[1]
    if (target) {
      const recordsPath = resolveRelativePath(update.cachePath.replace(/\/[^/]+$/, ''), target)
      pkg.write(
        recordsPath,
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
          '<pivotCacheRecords ' +
          'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
          'count="0"/>',
      )
      touchedEntries.add(recordsPath)
    }
  }
}

function refToArea(ref: string): PivotAddition['location'] {
  const area = parseAreaRef(ref)
  return {
    startRow: area.start.row,
    startColumn: area.start.col,
    endRow: area.end.row,
    endColumn: area.end.col,
  }
}

function unescapeAttribute(input: string): string {
  return input
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&amp;', '&')
}
