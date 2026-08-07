import type {
  WorkbookCellEdit,
  WorkbookChartEdit,
  WorkbookHyperlinkEdit,
  WorkbookRichRun,
  WorkbookSheetOp,
  WorkbookStructuralOp,
  WorkbookPivotAdd,
  WorkbookStyleEdit,
  WorkbookTableAdd,
  WorkbookVisualAdd,
  WorkbookVisualEdit,
  WorkbookVisualObject,
} from '../shared/desktop-api'
import { columnLabel, parseRange } from '../domain/cell-address'
import { splitSheetRef } from '../domain/chart-visual'
import { ADDABLE_SHAPE_TYPES } from '../shared/shape-types'
import { INDENT_STEP_PX } from './selection-format'

/// Tracks the user's cell edits on a streamed external workbook. Streaming
/// evicts and re-installs viewport cells, so the journal is both the save
/// payload and the overlay that re-applies edits after each re-install.
export interface JournalEntry {
  readonly row: number
  readonly column: number
  /// true when the user changed the cell's content (not just its style).
  readonly hasValue: boolean
  readonly value: string | number | boolean | null
  readonly formula?: string
  readonly style?: WorkbookStyleEdit
  /// Per-run styling of a rich-text cell; value holds the concatenated text.
  readonly rich?: readonly WorkbookRichRun[]
  /// Clear Formats / Clear All: the cell returns to the default style before
  /// any remaining `style` delta applies.
  readonly styleReset?: boolean
}

export type StructuralJournalOp =
  | {
      readonly kind: 'insert-rows' | 'remove-rows' | 'insert-cols' | 'remove-cols'
      readonly index: number
      readonly count: number
    }
  | {
      readonly kind: 'merge-cells' | 'unmerge-cells'
      readonly range: {
        readonly startRow: number
        readonly endRow: number
        readonly startColumn: number
        readonly endColumn: number
      }
    }
  | {
      readonly kind: 'set-row-size' | 'set-col-size'
      readonly start: number
      readonly end: number
      /// File units: points for rows, character width for columns. null resets
      /// to the sheet default (auto height / default width).
      readonly size: number | null
    }
  | {
      readonly kind: 'set-rows-hidden' | 'set-cols-hidden'
      readonly start: number
      readonly end: number
      readonly hidden: boolean
    }
  | {
      readonly kind: 'set-rows-outline' | 'set-cols-outline'
      readonly start: number
      readonly end: number
      /// Absolute outline level 0-7; 0 removes the attribute. An omitted
      /// collapsed leaves the file's collapsed flag untouched.
      readonly level: number
      readonly collapsed?: boolean
    }

/// Normalized sheet-level operations. Removing a sheet only marks it (its
/// cell entries stay, so an undo that re-inserts the sheet loses nothing);
/// the save payload skips everything belonging to marked sheets.
export interface SheetJournal {
  /// Univer sheet id → user-visible name, for sheets created this session.
  /// Duplicates carry the id of the sheet they were copied from.
  readonly added: Map<string, { name: string; sourceSheetId?: string }>
  /// Removed sheet ids — original sheets and added ones alike.
  readonly removed: Set<string>
  /// Original sheet id → new name (dropped when renamed back).
  readonly renamed: Map<string, string>
  /// Sheet id → desired visibility (dropped when toggled back to original).
  readonly hidden: Map<string, boolean>
  /// The tab order changed; the save sends the final on-screen order.
  orderDirty: boolean
}

export interface EditJournal {
  readonly cells: Map<string, Map<string, JournalEntry>>
  /// Ordered per sheet; cell entries are kept in post-operation coordinates.
  readonly structuralOps: Map<string, StructuralJournalOp[]>
  /// Keyed by chart part path; successive edits to one chart merge.
  readonly chartEdits: Map<string, Omit<WorkbookChartEdit, 'chartPath'>>
  /// Floating visuals created this session; the save writes each as a new
  /// drawing/chart part pair.
  readonly visualAdds: WorkbookVisualObject[]
  /// Edits to visuals already in the file, keyed by visual id: an anchor
  /// rewrite (move/resize) or a removal, located by (drawingPath, index).
  readonly visualEdits: Map<string, VisualEditEntry>
  /// Tables created this session; the save writes each as a new xl/tables
  /// part registered on its worksheet.
  readonly tableAdds: WorkbookTableAdd[]
  /// Pivots created this session; the baked cells ride the normal cell
  /// journal, the save additionally writes native pivot parts.
  readonly pivotAdds: WorkbookPivotAdd[]
  /// Sparkline groups created this session (x14 extLst on save).
  readonly sparklineAdds: SparklineAddEntry[]
  readonly sheets: SheetJournal
  /// Sheets whose filter changed; the save snapshots their live filter model
  /// instead of replaying mutations.
  readonly filterDirty: Set<string>
  /// Sheets whose conditional formatting changed; the save snapshots the full
  /// rule set declaratively (same recipe as filters).
  readonly cfDirty: Set<string>
  /// Sheets whose data validation changed; saved like cfDirty.
  readonly dvDirty: Set<string>
  /// Desired sheet-protection state (dropped when toggled back to original).
  readonly sheetProtection: Map<string, boolean>
  /// The defined-name set changed; the save snapshots the full model.
  readonly definedNames: { dirty: boolean }
  /// sheetId → "row:column" → link target ('#Sheet!A1' internal, URL
  /// external), or null for a removed link. Coordinates live in the same
  /// post-operation space as cell entries.
  readonly hyperlinks: Map<string, Map<string, string | null>>
  /// sheetId → accumulated Page Layout settings; only fields the user touched
  /// this session are present, everything else stays verbatim in the file.
  readonly pageSetup: Map<string, PageSetupJournalState>
  /// Sheets whose notes changed; the save snapshots the live note set (same
  /// recipe as filters).
  readonly noteDirty: Set<string>
  /// pivotCacheDefinition part paths whose pivots were recomputed this
  /// session; the save flags them refreshOnLoad so Excel rebuilds its cache.
  readonly pivotCacheRefresh: Set<string>
  /// Pivots whose layout grew during refresh: cachePath → new output area. On
  /// save, the pivotTableDefinition's location ref is widened to the new area
  /// (multiple refreshes of the same pivot keep only the last one). When relayout
  /// (an A3 layout edit) is present, the saver regenerates from the new layout
  /// definition/cacheDefinition.
  readonly pivotRefreshUpdates: Map<
    string,
    {
      sheetId: string
      newOutputRef: string
      relayout?: WorkbookPivotAdd
    }
  >
}

/// One printed header or footer: Excel's left/center/right sections. Text
/// carries Excel field codes verbatim (&P page, &N pages, &D date, &T time,
/// &F file name, &A sheet name, && literal ampersand).
export interface HeaderFooterParts {
  readonly left?: string
  readonly center?: string
  readonly right?: string
}

export interface PageSetupJournalState {
  orientation?: 'portrait' | 'landscape'
  /// OOXML paper-size code (1 = Letter, 9 = A4, …).
  paperSize?: number
  scale?: number
  /// 0 = automatic.
  fitToWidth?: number
  fitToHeight?: number
  /// Scale and fit are exclusive; whichever the user touched last wins.
  fitToPage?: boolean
  margins?: 'normal' | 'wide' | 'narrow'
  printGridlines?: boolean
  printHeadings?: boolean
  showGridlines?: boolean
  /// sheetView/@showFormulas: the sheet renders formulas instead of values.
  showFormulas?: boolean
  /// A1 range to print, or null to clear the print area.
  printArea?: string | null
  /// Rows repeated at the top of every page ("1:2"), or null to clear.
  printTitles?: string | null
  /// Frozen pane counts (0/0 = unfrozen); saved as sheetView <pane>.
  frozenRows?: number
  frozenColumns?: number
  /// Printed header/footer, or null to clear; saved as <headerFooter>.
  header?: HeaderFooterParts | null
  footer?: HeaderFooterParts | null
}

interface CellRange {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

export function createEditJournal(): EditJournal {
  return {
    cells: new Map(),
    structuralOps: new Map(),
    chartEdits: new Map(),
    visualAdds: [],
    visualEdits: new Map(),
    tableAdds: [],
    pivotAdds: [],
    sparklineAdds: [],
    sheets: {
      added: new Map(),
      removed: new Set(),
      renamed: new Map(),
      hidden: new Map(),
      orderDirty: false,
    },
    filterDirty: new Set(),
    cfDirty: new Set(),
    dvDirty: new Set(),
    sheetProtection: new Map(),
    definedNames: { dirty: false },
    hyperlinks: new Map(),
    pageSetup: new Map(),
    noteDirty: new Set(),
    pivotCacheRefresh: new Set(),
    pivotRefreshUpdates: new Map(),
  }
}

export function recordNoteChange(journal: EditJournal, sheetId: string): void {
  journal.noteDirty.add(sheetId)
}

export function recordPivotCacheRefresh(journal: EditJournal, cachePath: string): void {
  journal.pivotCacheRefresh.add(cachePath)
}

/// Records one pivot output-area expansion after layout growth (location ref
/// write-back); when relayout is present it is an A3 layout edit, and the saver
/// regenerates definition/cacheDefinition from it.
export function recordPivotRefreshUpdate(
  journal: EditJournal,
  cachePath: string,
  sheetId: string,
  newOutputRef: string,
  relayout?: WorkbookPivotAdd,
): void {
  journal.pivotRefreshUpdates.set(cachePath, {
    sheetId,
    newOutputRef,
    ...(relayout === undefined ? {} : { relayout }),
  })
}

export function recordPageSetup(
  journal: EditJournal,
  sheetId: string,
  patch: PageSetupJournalState,
): void {
  const state = journal.pageSetup.get(sheetId) ?? {}
  journal.pageSetup.set(sheetId, { ...state, ...patch })
}

export function toSavePageSetupStates(
  journal: EditJournal,
): ({ sheetId: string } & PageSetupJournalState)[] {
  return [...journal.pageSetup]
    .filter(([sheetId]) => !isSheetRemoved(journal, sheetId))
    .filter(([, state]) => Object.keys(state).length > 0)
    .map(([sheetId, state]) => ({ sheetId, ...state }))
}

export function recordDefinedNamesChange(journal: EditJournal): void {
  journal.definedNames.dirty = true
}

export function recordSheetProtection(
  journal: EditJournal,
  sheetId: string,
  desired: boolean,
  original: boolean,
): void {
  if (desired === original) journal.sheetProtection.delete(sheetId)
  else journal.sheetProtection.set(sheetId, desired)
}

export function recordCfChange(journal: EditJournal, sheetId: string): void {
  journal.cfDirty.add(sheetId)
}

export function recordDvChange(journal: EditJournal, sheetId: string): void {
  journal.dvDirty.add(sheetId)
}

export function recordHyperlinkEdit(
  journal: EditJournal,
  sheetId: string,
  row: number,
  column: number,
  target: string | null,
): void {
  const links = journal.hyperlinks.get(sheetId) ?? new Map<string, string | null>()
  links.set(`${row}:${column}`, target)
  journal.hyperlinks.set(sheetId, links)
}

export function hyperlinkEditAt(
  journal: EditJournal,
  sheetId: string,
  row: number,
  column: number,
): string | null | undefined {
  return journal.hyperlinks.get(sheetId)?.get(`${row}:${column}`)
}

export function toSaveHyperlinkEdits(journal: EditJournal): WorkbookHyperlinkEdit[] {
  const edits: WorkbookHyperlinkEdit[] = []
  for (const [sheetId, links] of journal.hyperlinks) {
    if (isSheetRemoved(journal, sheetId)) continue
    for (const [key, target] of links) {
      const [row, column] = key.split(':').map(Number)
      if (row === undefined || column === undefined) continue
      edits.push({ sheetId, row, column, target })
    }
  }
  return edits
}

export function recordFilterChange(journal: EditJournal, sheetId: string): void {
  journal.filterDirty.add(sheetId)
}

export function recordSheetInsert(journal: EditJournal, sheetId: string, name: string): void {
  // A re-insert of a previously removed sheet (undo of the removal) just
  // clears the mark; its journal entries were kept and stay valid.
  if (journal.sheets.removed.delete(sheetId)) return
  journal.sheets.added.set(sheetId, { name })
}

export function recordSheetRemove(journal: EditJournal, sheetId: string): void {
  journal.sheets.removed.add(sheetId)
}

/// Records a sheet copy. The save clones the source's worksheet part, so the
/// journal only needs the source's *pending* edits copied across — they bring
/// the clone up to the source's on-screen state at duplication time.
export function recordSheetDuplicate(
  journal: EditJournal,
  sheetId: string,
  name: string,
  sourceSheetId: string,
): void {
  if (journal.sheets.removed.delete(sheetId)) return
  journal.sheets.added.set(sheetId, { name, sourceSheetId })
  const sourceCells = journal.cells.get(sourceSheetId)
  if (sourceCells && sourceCells.size > 0) {
    journal.cells.set(sheetId, new Map(sourceCells))
  }
  const sourceLinks = journal.hyperlinks.get(sourceSheetId)
  if (sourceLinks && sourceLinks.size > 0) {
    journal.hyperlinks.set(sheetId, new Map(sourceLinks))
  }
  const sourceOps = journal.structuralOps.get(sourceSheetId)
  if (sourceOps && sourceOps.length > 0) {
    journal.structuralOps.set(sheetId, [...sourceOps])
  }
}

/// Walks duplicate-of-duplicate chains back to a sheet that exists in the
/// file. Returns undefined when the chain ends at a sheet created this
/// session — its content is entirely journal-owned, so a blank part works.
export function resolveDuplicateSource(journal: EditJournal, sheetId: string): string | undefined {
  let current = sheetId
  for (;;) {
    const added = journal.sheets.added.get(current)
    if (!added) return current
    if (added.sourceSheetId === undefined) return undefined
    current = added.sourceSheetId
  }
}

export function recordSheetRename(
  journal: EditJournal,
  sheetId: string,
  name: string,
  originalName: string | undefined,
): void {
  const added = journal.sheets.added.get(sheetId)
  if (added) {
    journal.sheets.added.set(sheetId, { name })
    return
  }
  if (name === originalName) journal.sheets.renamed.delete(sheetId)
  else journal.sheets.renamed.set(sheetId, name)
}

export function isSheetRemoved(journal: EditJournal, sheetId: string): boolean {
  return journal.sheets.removed.has(sheetId)
}

export function recordSheetOrderChange(journal: EditJournal): void {
  journal.sheets.orderDirty = true
}

export function recordSheetHidden(
  journal: EditJournal,
  sheetId: string,
  hidden: boolean,
  originallyHidden: boolean,
): void {
  if (hidden === originallyHidden) journal.sheets.hidden.delete(sheetId)
  else journal.sheets.hidden.set(sheetId, hidden)
}

export function sheetOpCount(journal: EditJournal): number {
  const { added, removed, renamed, hidden, orderDirty } = journal.sheets
  let total = orderDirty ? 1 : 0
  for (const sheetId of added.keys()) if (!removed.has(sheetId)) total += 1
  for (const sheetId of removed) if (!added.has(sheetId)) total += 1
  for (const sheetId of renamed.keys()) if (!removed.has(sheetId)) total += 1
  for (const sheetId of hidden.keys()) if (!removed.has(sheetId)) total += 1
  return total
}

export function toSaveSheetOps(journal: EditJournal): WorkbookSheetOp[] {
  const { added, removed, renamed, hidden, orderDirty } = journal.sheets
  const ops: WorkbookSheetOp[] = []
  for (const [sheetId, sheet] of added) {
    if (removed.has(sheetId)) continue
    const source =
      sheet.sourceSheetId === undefined
        ? undefined
        : resolveDuplicateSource(journal, sheet.sourceSheetId)
    ops.push(
      source === undefined
        ? { kind: 'add-sheet', sheetId, name: sheet.name }
        : { kind: 'duplicate-sheet', sheetId, name: sheet.name, sourceSheetId: source },
    )
  }
  for (const sheetId of removed) {
    if (!added.has(sheetId)) ops.push({ kind: 'remove-sheet', sheetId })
  }
  for (const [sheetId, newName] of renamed) {
    if (!removed.has(sheetId)) ops.push({ kind: 'rename-sheet', sheetId, newName })
  }
  for (const [sheetId, isHidden] of hidden) {
    if (!removed.has(sheetId)) ops.push({ kind: 'set-sheet-hidden', sheetId, hidden: isHidden })
  }
  if (orderDirty) ops.push({ kind: 'reorder-sheets' })
  return ops
}

export function recordChartEdit(
  journal: EditJournal,
  chartPath: string,
  edit: Omit<WorkbookChartEdit, 'chartPath'>,
): void {
  let previous = journal.chartEdits.get(chartPath)
  if (edit.seriesSet && previous) {
    // A full series replacement invalidates any earlier per-index series
    // edits and per-series colors — their indices refer to dropped series.
    const {
      series: _series,
      seriesColors: _colors,
      pointColors: _points,
      pointExplosions: _explosions,
      ...rest
    } = previous
    previous = rest
  }
  journal.chartEdits.set(chartPath, {
    ...previous,
    ...edit,
    ...(previous?.seriesColors || edit.seriesColors
      ? { seriesColors: { ...previous?.seriesColors, ...edit.seriesColors } }
      : {}),
    ...(previous?.pointColors || edit.pointColors
      ? { pointColors: mergePointColorEdits(previous?.pointColors, edit.pointColors) }
      : {}),
    ...(previous?.pointExplosions || edit.pointExplosions
      ? { pointExplosions: { ...previous?.pointExplosions, ...edit.pointExplosions } }
      : {}),
    ...(previous?.axisTitles || edit.axisTitles
      ? { axisTitles: { ...previous?.axisTitles, ...edit.axisTitles } }
      : {}),
    ...(previous?.valueAxis || edit.valueAxis
      ? { valueAxis: { ...previous?.valueAxis, ...edit.valueAxis } }
      : {}),
    ...(previous?.series || edit.series
      ? { series: mergeSeriesEdits(previous?.series, edit.series) }
      : {}),
  })
}

/// Point-color edits merge per series, then per point index.
function mergePointColorEdits(
  previous: WorkbookChartEdit['pointColors'],
  next: WorkbookChartEdit['pointColors'],
): NonNullable<WorkbookChartEdit['pointColors']> {
  const merged: Record<string, Record<string, string>> = { ...previous }
  for (const [seriesKey, points] of Object.entries(next ?? {})) {
    merged[seriesKey] = { ...merged[seriesKey], ...points }
  }
  return merged
}

/// Successive edits to one series merge field-by-field, keyed by index.
function mergeSeriesEdits(
  previous: WorkbookChartEdit['series'],
  next: WorkbookChartEdit['series'],
): NonNullable<WorkbookChartEdit['series']> {
  const byIndex = new Map<number, NonNullable<WorkbookChartEdit['series']>[number]>()
  for (const entry of [...(previous ?? []), ...(next ?? [])]) {
    byIndex.set(entry.index, { ...byIndex.get(entry.index), ...entry })
  }
  return [...byIndex.values()]
}

export function recordTableAdd(journal: EditJournal, table: WorkbookTableAdd): void {
  journal.tableAdds.push(table)
}

/// Updates the stored area and optionally columnNames of a session-added table.
/// Returns false when no matching table is found (fail-closed: caller should throw).
export function updateTableAdd(
  journal: EditJournal,
  sheetId: string,
  name: string,
  patch: {
    area?: { startRow: number; startColumn: number; endRow: number; endColumn: number }
    columnNames?: readonly string[]
  },
): boolean {
  const index = journal.tableAdds.findIndex(
    (t) => t.sheetId === sheetId && t.name.toLowerCase() === name.toLowerCase(),
  )
  if (index < 0) return false
  const existing = journal.tableAdds[index]!
  ;(journal.tableAdds as WorkbookTableAdd[])[index] = {
    ...existing,
    ...(patch.area !== undefined ? { area: patch.area } : {}),
    ...(patch.columnNames !== undefined ? { columnNames: [...patch.columnNames] } : {}),
  }
  return true
}

export function recordPivotAdd(journal: EditJournal, pivot: WorkbookPivotAdd): void {
  journal.pivotAdds.push(pivot)
}

export interface SparklineAddEntry {
  readonly id: string
  readonly sheetId: string
  readonly type: 'line' | 'column' | 'stacked'
  readonly color?: string | undefined
  readonly cells: readonly { cell: string; sourceRef: string }[]
}

export function recordSparklineAdd(journal: EditJournal, entry: SparklineAddEntry): void {
  journal.sparklineAdds.push(entry)
}

export function removeSparklineAdd(journal: EditJournal, id: string): boolean {
  const index = journal.sparklineAdds.findIndex((entry) => entry.id === id)
  if (index < 0) return false
  journal.sparklineAdds.splice(index, 1)
  return true
}

export function toSaveSparklineAdds(journal: EditJournal): {
  sheetId: string
  type: 'line' | 'column' | 'stacked'
  color?: string
  cells: { cell: string; sourceRef: string }[]
}[] {
  return journal.sparklineAdds
    .filter((entry) => !isSheetRemoved(journal, entry.sheetId))
    .map((entry) => ({
      sheetId: entry.sheetId,
      type: entry.type,
      ...(entry.color === undefined ? {} : { color: entry.color }),
      cells: entry.cells.map((cell) => ({ ...cell })),
    }))
}

/// Removes a session-added table (convert-to-range semantics: the baked
/// cells stay). False for file-native tables — they are view-only.
export function removeTableAdd(journal: EditJournal, sheetId: string, name: string): boolean {
  const index = journal.tableAdds.findIndex(
    (table) => table.sheetId === sheetId && table.name.toLowerCase() === name.toLowerCase(),
  )
  if (index < 0) return false
  journal.tableAdds.splice(index, 1)
  return true
}

/// Session pivot adds for the save request; pivots whose output or source
/// sheet was removed drop.
export function toSavePivotAdds(journal: EditJournal): WorkbookPivotAdd[] {
  return journal.pivotAdds.filter(
    (pivot) =>
      !isSheetRemoved(journal, pivot.sheetId) && !isSheetRemoved(journal, pivot.sourceSheetId),
  )
}

export function recordVisualAdd(journal: EditJournal, visual: WorkbookVisualObject): void {
  journal.visualAdds.push(visual)
}

export interface VisualEditEntry {
  readonly sheetId: string
  readonly drawingPath: string
  readonly drawingIndex: number
  readonly remove?: true
  readonly anchor?: WorkbookVisualObject['anchor']
}

/// Records a move/resize or removal of a visual that already lives in the
/// file. False when the visual carries no drawing locator (its part shape
/// is one the sidecar could not pin down) — the caller shows a message.
export function recordVisualEdit(
  journal: EditJournal,
  visual: WorkbookVisualObject,
  changes: { remove?: true; anchor?: WorkbookVisualObject['anchor'] },
): boolean {
  if (visual.drawingPath === undefined || visual.drawingIndex === undefined) return false
  const previous = journal.visualEdits.get(visual.id)
  // A removal wins over any earlier move; a later move revives nothing.
  const remove = changes.remove ?? previous?.remove
  const anchor = changes.anchor ?? previous?.anchor
  journal.visualEdits.set(visual.id, {
    sheetId: visual.sheetId,
    drawingPath: visual.drawingPath,
    drawingIndex: visual.drawingIndex,
    ...(remove ? { remove: true as const } : {}),
    ...(anchor ? { anchor } : {}),
  })
  return true
}

/// Drops a session visual outright (its object only exists in the journal).
export function removeVisualAdd(journal: EditJournal, visualId: string): boolean {
  const index = journal.visualAdds.findIndex((visual) => visual.id === visualId)
  if (index < 0) return false
  journal.visualAdds.splice(index, 1)
  return true
}

/// In-place edit (move/resize/retext) of a visual added this session; the
/// save writes the updated object, so no separate edit record is needed.
/// False when the id is not a session visual (file visuals are view-only).
export function updateVisualAdd(
  journal: EditJournal,
  visualId: string,
  changes: { anchor?: WorkbookVisualObject['anchor']; text?: string; fillColor?: string },
): boolean {
  const index = journal.visualAdds.findIndex((visual) => visual.id === visualId)
  const current = journal.visualAdds[index]
  if (!current) return false
  journal.visualAdds[index] = {
    ...current,
    ...(changes.anchor === undefined ? {} : { anchor: changes.anchor }),
    ...(changes.text === undefined ? {} : { text: changes.text }),
    ...(changes.fillColor === undefined ? {} : { fillColor: changes.fillColor }),
  }
  return true
}

/// Session table adds for the save request; tables on removed sheets drop.
export function toSaveTableAdds(journal: EditJournal): WorkbookTableAdd[] {
  return journal.tableAdds.filter((table) => !isSheetRemoved(journal, table.sheetId))
}

/// File-visual edits for the save request; edits on removed sheets drop
/// (the sheet removal takes the whole drawing part with it).
export function toSaveVisualEdits(journal: EditJournal): WorkbookVisualEdit[] {
  const edits: WorkbookVisualEdit[] = []
  for (const entry of journal.visualEdits.values()) {
    if (isSheetRemoved(journal, entry.sheetId)) continue
    edits.push({
      drawingPath: entry.drawingPath,
      drawingIndex: entry.drawingIndex,
      ...(entry.remove ? { remove: true as const } : {}),
      ...(entry.anchor ? { anchor: entry.anchor } : {}),
    })
  }
  return edits
}

/// Converts session visuals to the save-request wire shape. Skips visuals on
/// removed sheets and anything that is not a supported chart or shape.
export function toSaveVisualAdds(journal: EditJournal): WorkbookVisualAdd[] {
  const additions: WorkbookVisualAdd[] = []
  for (const visual of journal.visualAdds) {
    if (isSheetRemoved(journal, visual.sheetId)) continue
    if (visual.kind === 'image') {
      const mediaType = (['image/png', 'image/jpeg', 'image/gif'] as const).find(
        (type) => type === visual.mediaType,
      )
      const base64 = visual.mediaDataUrl?.split(',')[1]
      if (!mediaType || !base64) continue
      additions.push({
        sheetId: visual.sheetId,
        anchor: visual.anchor,
        image: { mediaType, base64 },
      })
      continue
    }
    if (visual.kind === 'shape') {
      const shapeType = ADDABLE_SHAPE_TYPES.find((type) => type === visual.shapeType)
      if (!shapeType) continue
      additions.push({
        sheetId: visual.sheetId,
        anchor: visual.anchor,
        shape: {
          shapeType,
          ...(visual.fillColor === undefined ? {} : { fillColor: visual.fillColor }),
          ...(visual.text === undefined ? {} : { text: visual.text }),
          ...(visual.name === 'TextBox' ? { isTextBox: true } : {}),
        },
      })
      continue
    }
    if (visual.kind !== 'chart' || !visual.chart) continue
    const chartType =
      visual.chart.chartTypes.includes('barChart') && visual.chart.chartTypes.includes('lineChart')
        ? 'combo'
        : visual.chart.chartTypes.includes('pieChart')
          ? 'pie'
          : visual.chart.chartTypes.includes('doughnutChart')
            ? 'doughnut'
            : visual.chart.chartTypes.includes('scatterChart')
              ? 'scatter'
              : visual.chart.chartTypes.includes('radarChart')
                ? 'radar'
                : visual.chart.chartTypes.includes('lineChart')
                  ? 'line'
                  : visual.chart.chartTypes.includes('areaChart')
                    ? 'area'
                    : visual.chart.chartTypes.includes('barChart')
                      ? visual.chart.barDirection === 'bar'
                        ? 'bar'
                        : 'column'
                      : null
    if (chartType === null) continue
    const axisTitles = {
      ...(typeof visual.chart.axisTitles?.category === 'string' && visual.chart.axisTitles.category
        ? { category: visual.chart.axisTitles.category }
        : {}),
      ...(typeof visual.chart.axisTitles?.value === 'string' && visual.chart.axisTitles.value
        ? { value: visual.chart.axisTitles.value }
        : {}),
    }
    additions.push({
      sheetId: visual.sheetId,
      anchor: visual.anchor,
      chart: {
        chartType,
        title: visual.chart.title,
        series: visual.chart.series.map((series) => ({
          name: series.name,
          categories: series.categories,
          values: series.values,
          ...(series.valuesRef === undefined ? {} : { valuesRef: series.valuesRef }),
          ...(series.categoriesRef === undefined ? {} : { categoriesRef: series.categoriesRef }),
          ...(series.color === undefined ? {} : { color: series.color }),
          ...(series.pointColors === undefined || series.pointColors.length === 0
            ? {}
            : {
                pointColors: Object.fromEntries(
                  series.pointColors.map((point) => [String(point.index), point.color]),
                ),
              }),
          ...(series.explosionPct === undefined
            ? {}
            : { explosionPct: Math.round(series.explosionPct) }),
          ...(series.pointExplosions === undefined || series.pointExplosions.length === 0
            ? {}
            : {
                pointExplosions: Object.fromEntries(
                  series.pointExplosions.map((point) => [
                    String(point.index),
                    Math.round(point.pct),
                  ]),
                ),
              }),
        })),
        ...(visual.chart.legend === undefined ? {} : { legend: visual.chart.legend }),
        ...(visual.chart.dataLabels === undefined ? {} : { dataLabels: visual.chart.dataLabels }),
        ...(visual.chart.dataLabelPosition === undefined
          ? {}
          : { dataLabelPosition: visual.chart.dataLabelPosition }),
        ...(visual.chart.dataLabelFormat === undefined
          ? {}
          : { dataLabelFormat: visual.chart.dataLabelFormat }),
        ...(Object.keys(axisTitles).length === 0 ? {} : { axisTitles }),
        ...(visual.chart.grouping === undefined ? {} : { grouping: visual.chart.grouping }),
        ...(visual.chart.gridlines === undefined ? {} : { gridlines: visual.chart.gridlines }),
        ...(visual.chart.valueAxis === undefined ? {} : { valueAxis: visual.chart.valueAxis }),
        ...(visual.chart.gapWidthPct === undefined
          ? {}
          : { gapWidthPct: Math.round(visual.chart.gapWidthPct) }),
        ...(visual.chart.holeSizePct === undefined
          ? {}
          : { holeSizePct: Math.round(visual.chart.holeSizePct) }),
      },
    })
  }
  return additions
}

export function toSaveChartEdits(journal: EditJournal): WorkbookChartEdit[] {
  return [...journal.chartEdits].map(([chartPath, edit]) => ({ chartPath, ...edit }))
}

interface RowColumnShift {
  readonly axis: 'row' | 'column'
  readonly removing: boolean
  readonly index: number
  readonly count: number
}

function toRowColumnShift(op: { kind: string; index: number; count: number }): RowColumnShift {
  return {
    axis: op.kind === 'insert-rows' || op.kind === 'remove-rows' ? 'row' : 'column',
    removing: op.kind === 'remove-rows' || op.kind === 'remove-cols',
    index: op.index,
    count: op.count,
  }
}

/// Same clamping the save-side drawing shift applies: marks inside a removed
/// span collapse onto its start (offset reset to 0).
function moveAnchorMark(
  position: number,
  shift: RowColumnShift,
): { position: number; clamped: boolean } {
  if (shift.removing) {
    if (position >= shift.index + shift.count) {
      return { position: position - shift.count, clamped: false }
    }
    if (position >= shift.index) return { position: shift.index, clamped: true }
    return { position, clamped: false }
  }
  return {
    position: position >= shift.index ? position + shift.count : position,
    clamped: false,
  }
}

function shiftVisualAnchor(
  anchor: WorkbookVisualObject['anchor'],
  shift: RowColumnShift,
): WorkbookVisualObject['anchor'] {
  const from = moveAnchorMark(shift.axis === 'row' ? anchor.fromRow : anchor.fromColumn, shift)
  const to = moveAnchorMark(shift.axis === 'row' ? anchor.toRow : anchor.toColumn, shift)
  return shift.axis === 'row'
    ? {
        ...anchor,
        fromRow: from.position,
        toRow: to.position,
        fromRowOffset: from.clamped ? 0 : anchor.fromRowOffset,
        toRowOffset: to.clamped ? 0 : anchor.toRowOffset,
      }
    : {
        ...anchor,
        fromColumn: from.position,
        toColumn: to.position,
        fromColumnOffset: from.clamped ? 0 : anchor.fromColumnOffset,
        toColumnOffset: to.clamped ? 0 : anchor.toColumnOffset,
      }
}

/// Shifts a chart `c:f`-style reference on the edited sheet; null when the
/// whole referenced range was removed, the input when it cannot be parsed or
/// points at another sheet.
function shiftSeriesRef(ref: string, sheetName: string, shift: RowColumnShift): string | null {
  const split = splitSheetRef(ref)
  if (!split || split.sheetName !== sheetName) return ref
  let area
  try {
    area = parseRange(split.range)
  } catch {
    return ref
  }
  let start = shift.axis === 'row' ? area.startRow : area.startColumn
  let end = shift.axis === 'row' ? area.endRow : area.endColumn
  const { index, count } = shift
  if (shift.removing) {
    if (start >= index && end < index + count) return null
    start = start >= index + count ? start - count : start >= index ? index : start
    end = end >= index + count ? end - count : end >= index ? index - 1 : end
  } else if (end >= index) {
    // An insert above shifts the range, an insert inside extends it.
    if (start >= index) start += count
    end += count
  }
  const moved =
    shift.axis === 'row'
      ? { ...area, startRow: start, endRow: end }
      : { ...area, startColumn: start, endColumn: end }
  const name = sheetName.replace(/'/g, "''")
  const cell = (row: number, column: number): string => `$${columnLabel(column)}$${row + 1}`
  const head = cell(moved.startRow, moved.startColumn)
  const tail = cell(moved.endRow, moved.endColumn)
  return `'${name}'!${head}${head === tail ? '' : `:${tail}`}`
}

/// A destroyed reference drops so the save falls back to literal values.
function shiftSeriesEntry<
  T extends { valuesRef?: string | undefined; categoriesRef?: string | undefined },
>(entry: T, sheetName: string, shift: RowColumnShift): T {
  const next = { ...entry }
  if (entry.valuesRef !== undefined) {
    const moved = shiftSeriesRef(entry.valuesRef, sheetName, shift)
    if (moved === null) delete next.valuesRef
    else next.valuesRef = moved
  }
  if (entry.categoriesRef !== undefined) {
    const moved = shiftSeriesRef(entry.categoriesRef, sheetName, shift)
    if (moved === null) delete next.categoriesRef
    else next.categoriesRef = moved
  }
  return next
}

/// Shifts one visual's anchor (own sheet) and chart references (matched by
/// sheet name) into the post-operation coordinate space. Also used on the
/// in-memory file visuals so the preview and live sync track the new space.
export function shiftVisualForStructuralOp(
  visual: WorkbookVisualObject,
  sheetId: string,
  sheetName: string | undefined,
  op: StructuralJournalOp,
): WorkbookVisualObject {
  if (!('index' in op)) return visual
  const shift = toRowColumnShift(op)
  const anchor =
    visual.sheetId === sheetId ? shiftVisualAnchor(visual.anchor, shift) : visual.anchor
  const chart =
    visual.chart === undefined || sheetName === undefined
      ? visual.chart
      : {
          ...visual.chart,
          series: visual.chart.series.map((series) => shiftSeriesEntry(series, sheetName, shift)),
        }
  return { ...visual, anchor, ...(chart === undefined ? {} : { chart }) }
}

/// Records a row/column insert/removal and shifts existing cell entries into
/// the new coordinate space (entries inside a removed range are dropped).
/// `sheetName` (the edited sheet's name) additionally shifts chart series
/// references held in the journal.
export function recordStructuralOp(
  journal: EditJournal,
  sheetId: string,
  op: StructuralJournalOp,
  sheetName?: string,
): void {
  const ops = journal.structuralOps.get(sheetId) ?? []
  // Undo of a just-recorded insert arrives as its exact inverse remove: cancel the
  // pair instead of stacking two shifting ops. Otherwise a refused insert (e.g. the
  // structural-shift guard) could never be cleared from the journal and every later
  // save would keep hitting the same guard.
  const last = ops[ops.length - 1]
  const cancels =
    'index' in op &&
    (op.kind === 'remove-rows' || op.kind === 'remove-cols') &&
    last !== undefined &&
    last.kind === (op.kind === 'remove-rows' ? 'insert-rows' : 'insert-cols') &&
    'index' in last &&
    last.index === op.index &&
    last.count === op.count
  if (cancels) {
    ops.pop()
    if (ops.length === 0) journal.structuralOps.delete(sheetId)
    else journal.structuralOps.set(sheetId, ops)
  } else {
    ops.push(op)
    journal.structuralOps.set(sheetId, ops)
  }

  // Merges don't move cells; covered-cell clears arrive as value mutations.
  // (`in` narrowing: TS doesn't compose `||` checks on multi-literal kinds.)
  if (!('index' in op)) return
  const rowColumnOp = op
  const axis = op.kind === 'insert-rows' || op.kind === 'remove-rows' ? 'row' : 'column'
  const removing = op.kind === 'remove-rows' || op.kind === 'remove-cols'
  const movePosition = (position: number): number | null => {
    const { index, count } = rowColumnOp
    if (removing) {
      if (position >= index && position < index + count) return null
      return position >= index + count ? position - count : position
    }
    return position >= index ? position + count : position
  }
  const sheetEntries = journal.cells.get(sheetId)
  if (sheetEntries && sheetEntries.size > 0) {
    const shifted = new Map<string, JournalEntry>()
    for (const entry of sheetEntries.values()) {
      const moved = movePosition(axis === 'row' ? entry.row : entry.column)
      if (moved === null) continue
      const next = axis === 'row' ? { ...entry, row: moved } : { ...entry, column: moved }
      shifted.set(`${next.row}:${next.column}`, next)
    }
    journal.cells.set(sheetId, shifted)
  }
  const links = journal.hyperlinks.get(sheetId)
  if (links && links.size > 0) {
    const shiftedLinks = new Map<string, string | null>()
    for (const [key, target] of links) {
      const [row, column] = key.split(':').map(Number)
      if (row === undefined || column === undefined) continue
      const moved = movePosition(axis === 'row' ? row : column)
      if (moved === null) continue
      const next = axis === 'row' ? `${moved}:${column}` : `${row}:${moved}`
      shiftedLinks.set(next, target)
    }
    journal.hyperlinks.set(sheetId, shiftedLinks)
  }
  // Session visuals and pending chart edits follow the same shift: the save
  // appends/applies them after the file's own structural pass, so they must
  // already live in post-operation coordinates.
  const shift = toRowColumnShift(rowColumnOp)
  journal.visualAdds.forEach((visual, at) => {
    journal.visualAdds[at] = shiftVisualForStructuralOp(visual, sheetId, sheetName, op)
  })
  for (const [id, entry] of journal.visualEdits) {
    if (entry.sheetId !== sheetId || entry.anchor === undefined) continue
    journal.visualEdits.set(id, { ...entry, anchor: shiftVisualAnchor(entry.anchor, shift) })
  }
  if (sheetName !== undefined) {
    for (const [chartPath, edit] of journal.chartEdits) {
      if (edit.series === undefined && edit.seriesSet === undefined) continue
      journal.chartEdits.set(chartPath, {
        ...edit,
        ...(edit.series === undefined
          ? {}
          : { series: edit.series.map((entry) => shiftSeriesEntry(entry, sheetName, shift)) }),
        ...(edit.seriesSet === undefined
          ? {}
          : {
              seriesSet: edit.seriesSet.map((entry) => shiftSeriesEntry(entry, sheetName, shift)),
            }),
      })
    }
  }
}

export function structuralOpCount(journal: EditJournal): number {
  let total = 0
  for (const ops of journal.structuralOps.values()) total += ops.length
  return total
}

export function toSaveStructuralOps(journal: EditJournal): WorkbookStructuralOp[] {
  const ops: WorkbookStructuralOp[] = []
  for (const [sheetId, sheetOps] of journal.structuralOps) {
    if (isSheetRemoved(journal, sheetId)) continue
    for (const op of sheetOps) ops.push({ sheetId, ...op })
  }
  return ops
}

/// Ingests a `sheet.mutation.set-range-values` payload. Returns the entries
/// that were recorded.
export function recordSetRangeValues(
  journal: EditJournal,
  sheetId: string,
  cellValue: unknown,
): JournalEntry[] {
  if (typeof cellValue !== 'object' || cellValue === null) return []
  const recorded: JournalEntry[] = []
  for (const [rowKey, rowValue] of Object.entries(cellValue)) {
    const row = Number(rowKey)
    if (!Number.isInteger(row) || row < 0) continue
    if (typeof rowValue !== 'object' || rowValue === null) continue
    for (const [columnKey, cell] of Object.entries(rowValue as Record<string, unknown>)) {
      const column = Number(columnKey)
      if (!Number.isInteger(column) || column < 0) continue
      const entry = mergeIntoJournal(journal, sheetId, row, column, cell)
      if (entry) recorded.push(entry)
    }
  }
  return recorded
}

/// Ingests a `sheet.mutation.set.numfmt` payload
/// (`{refMap: {key: {pattern}}, values: {key: {ranges}}}`).
export function recordSetNumfmt(
  journal: EditJournal,
  sheetId: string,
  params: unknown,
): JournalEntry[] {
  if (typeof params !== 'object' || params === null) return []
  const { refMap, values } = params as { refMap?: unknown; values?: unknown }
  if (
    typeof refMap !== 'object' ||
    refMap === null ||
    typeof values !== 'object' ||
    values === null
  ) {
    return []
  }
  const recorded: JournalEntry[] = []
  for (const [key, group] of Object.entries(values as Record<string, unknown>)) {
    const pattern = (refMap as Record<string, { pattern?: unknown } | undefined>)[key]?.pattern
    if (typeof pattern !== 'string' || pattern.length === 0) continue
    const ranges = (group as { ranges?: unknown }).ranges
    if (!Array.isArray(ranges)) continue
    for (const range of ranges) {
      if (typeof range !== 'object' || range === null) continue
      const { startRow, endRow, startColumn, endColumn } = range as Record<string, unknown>
      if (![startRow, endRow, startColumn, endColumn].every(isBoundedIndex)) continue
      for (let row = startRow as number; row <= (endRow as number); row += 1) {
        for (let column = startColumn as number; column <= (endColumn as number); column += 1) {
          const entry = mergeIntoJournal(journal, sheetId, row, column, {
            s: { n: { pattern } },
          })
          if (entry) recorded.push(entry)
        }
      }
    }
  }
  return recorded
}

/// Guards the numfmt range expansion: a malformed huge range must not spin.
function isBoundedIndex(value: unknown): boolean {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0 && value < 1_048_576
}

/// Journals a neutral style delta directly — for attributes Univer cannot
/// represent (cell protection). They save to the file but have no on-screen
/// rendering and no undo entry.
export function recordNeutralStyleEdit(
  journal: EditJournal,
  sheetId: string,
  row: number,
  column: number,
  delta: WorkbookStyleEdit,
): void {
  let sheetEntries = journal.cells.get(sheetId)
  const previous = sheetEntries?.get(`${row}:${column}`)
  const entry: JournalEntry = previous
    ? { ...previous, style: { ...previous.style, ...delta } }
    : { row, column, hasValue: false, value: null, style: delta }
  if (!sheetEntries) {
    sheetEntries = new Map()
    journal.cells.set(sheetId, sheetEntries)
  }
  sheetEntries.set(`${row}:${column}`, entry)
}

function mergeIntoJournal(
  journal: EditJournal,
  sheetId: string,
  row: number,
  column: number,
  cell: unknown,
): JournalEntry | null {
  let sheetEntries = journal.cells.get(sheetId)
  const previous = sheetEntries?.get(`${row}:${column}`)
  const entry = mergeCellIntoEntry(previous, row, column, cell)
  if (!entry) return null
  if (!sheetEntries) {
    sheetEntries = new Map()
    journal.cells.set(sheetId, sheetEntries)
  }
  sheetEntries.set(`${row}:${column}`, entry)
  return entry
}

function mergeCellIntoEntry(
  previous: JournalEntry | undefined,
  row: number,
  column: number,
  cell: unknown,
): JournalEntry | null {
  // A null cell is a full delete: content cleared, style reset to default.
  if (cell === null || cell === undefined) {
    return { row, column, hasValue: true, value: null, styleReset: true }
  }
  if (typeof cell !== 'object') return null
  const data = cell as Record<string, unknown>

  const styleDelta =
    typeof data.s === 'object' && data.s !== null
      ? toNeutralStyle(data.s as Record<string, unknown>)
      : undefined
  let style = styleDelta ? { ...previous?.style, ...styleDelta } : previous?.style
  let styleReset = previous?.styleReset === true
  // Clear Formats sends an explicit `s: null`: earlier deltas are pre-reset
  // and drop; later deltas stack on the default style.
  if ('s' in data && data.s === null) {
    styleReset = true
    style = undefined
  }

  let hasValue = previous?.hasValue ?? false
  let value: JournalEntry['value'] = previous?.value ?? null
  let formula = previous?.formula
  const formulaText =
    typeof data.f === 'string' && data.f.length > 0
      ? data.f.startsWith('=')
        ? data.f
        : `=${data.f}`
      : undefined
  let rich = previous?.rich
  const richText = extractRichText(data.p)
  if (formulaText) {
    hasValue = true
    value = null
    formula = formulaText
    rich = undefined
  } else if (richText !== undefined) {
    hasValue = true
    value = richText.text
    formula = undefined
    rich = richText.runs
  } else if ('v' in data) {
    // The formula engine writes calculation results as bare `{v}` mutations;
    // only an explicit `f: null` (editor overwrite) clears a journaled
    // formula — otherwise the value is just the formula's cached result.
    const isCalculationResult = previous?.formula !== undefined && !('f' in data)
    const raw = data.v
    if (isCalculationResult) {
      // keep the journaled formula; the file stores <f> and Excel recalcs
    } else if (raw === null || raw === undefined) {
      hasValue = true
      value = null
      formula = undefined
      rich = undefined
    } else if (
      typeof raw === 'string' ||
      typeof raw === 'boolean' ||
      (typeof raw === 'number' && Number.isFinite(raw))
    ) {
      hasValue = true
      value = raw
      formula = undefined
      rich = undefined
    }
  }

  if (!hasValue && !style && !styleReset) return null
  return {
    row,
    column,
    hasValue,
    value,
    ...(formula === undefined ? {} : { formula }),
    ...(style === undefined ? {} : { style }),
    ...(rich === undefined ? {} : { rich }),
    ...(styleReset ? { styleReset: true } : {}),
  }
}

function extractRichText(p: unknown): { text: string; runs?: WorkbookRichRun[] } | undefined {
  if (typeof p !== 'object' || p === null) return undefined
  const body = (p as Record<string, unknown>).body
  if (typeof body !== 'object' || body === null) return undefined
  const dataStream = (body as Record<string, unknown>).dataStream
  if (typeof dataStream !== 'string') return undefined
  // Univer streams paragraph breaks as \r; the file model (and xlsx) use \n.
  const text = dataStream
    .replace(/\r?\n$/, '')
    .replace(/\r$/, '')
    .replace(/\r/g, '\n')
  const runs = extractRichRuns(text, (body as Record<string, unknown>).textRuns)
  return runs === undefined ? { text } : { text, runs }
}

/// Univer text runs → neutral per-run styling covering the whole text
/// (gaps become unstyled runs). Undefined when nothing carries styling.
function extractRichRuns(text: string, textRuns: unknown): WorkbookRichRun[] | undefined {
  if (!Array.isArray(textRuns) || text.length === 0) return undefined
  const runs: WorkbookRichRun[] = []
  let cursor = 0
  let anyStyled = false
  const plain = (segment: string): WorkbookRichRun => ({
    text: segment,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
  })
  for (const raw of textRuns) {
    if (typeof raw !== 'object' || raw === null) continue
    const { st, ed, ts } = raw as { st?: unknown; ed?: unknown; ts?: unknown }
    if (typeof st !== 'number' || typeof ed !== 'number' || st < cursor || ed <= st) continue
    const start = Math.min(st, text.length)
    const end = Math.min(ed, text.length)
    if (start > cursor) runs.push(plain(text.slice(cursor, start)))
    const style = (typeof ts === 'object' && ts !== null ? ts : {}) as Record<string, unknown>
    const color = (style.cl as { rgb?: unknown } | undefined)?.rgb
    const run: WorkbookRichRun = {
      text: text.slice(start, end),
      bold: style.bl === 1,
      italic: style.it === 1,
      underline: (style.ul as { s?: unknown } | undefined)?.s === 1,
      strikethrough: (style.st as { s?: unknown } | undefined)?.s === 1,
      ...(typeof color === 'string' ? { color } : {}),
      ...(typeof style.fs === 'number' ? { size: style.fs } : {}),
      ...(typeof style.ff === 'string' ? { family: style.ff } : {}),
    }
    if (
      run.bold ||
      run.italic ||
      run.underline ||
      run.strikethrough ||
      run.color !== undefined ||
      run.size !== undefined ||
      run.family !== undefined
    ) {
      anyStyled = true
    }
    runs.push(run)
    cursor = end
  }
  if (!anyStyled) return undefined
  if (cursor < text.length) runs.push(plain(text.slice(cursor)))
  return runs.filter((run) => run.text.length > 0)
}

/// Univer BorderStyleTypes numeric values ↔ OOXML border style names.
const UNIVER_BORDER_TO_XLSX: Record<number, string> = {
  1: 'thin',
  2: 'hair',
  3: 'dotted',
  4: 'dashed',
  5: 'dashDot',
  6: 'dashDotDot',
  7: 'double',
  8: 'medium',
  9: 'mediumDashed',
  10: 'mediumDashDot',
  11: 'mediumDashDotDot',
  12: 'slantDashDot',
  13: 'thick',
}
const XLSX_BORDER_TO_UNIVER: Record<string, number> = Object.fromEntries(
  Object.entries(UNIVER_BORDER_TO_XLSX).map(([value, name]) => [name, Number(value)]),
)
const BORDER_EDGE_KEYS: readonly (readonly [
  't' | 'b' | 'l' | 'r',
  'borderTop' | 'borderBottom' | 'borderLeft' | 'borderRight',
])[] = [
  ['t', 'borderTop'],
  ['b', 'borderBottom'],
  ['l', 'borderLeft'],
  ['r', 'borderRight'],
]

type NeutralBorderEdge = NonNullable<WorkbookStyleEdit['borderTop']>

function univerBorderEdgeToNeutral(edge: unknown): NeutralBorderEdge | undefined {
  if (typeof edge !== 'object' || edge === null) return undefined
  const { s, cl } = edge as { s?: unknown; cl?: unknown }
  const style = typeof s === 'number' ? UNIVER_BORDER_TO_XLSX[s] : undefined
  if (style === undefined) return undefined
  const color = extractRgb(cl)
  return {
    style: style as NonNullable<NeutralBorderEdge>['style'],
    ...(color === undefined ? {} : { color }),
  }
}

const UNIVER_HORIZONTAL: Record<number, WorkbookStyleEdit['horizontalAlignment']> = {
  1: 'left',
  2: 'center',
  3: 'right',
  4: 'justify',
  6: 'distributed',
}
const UNIVER_VERTICAL: Record<number, WorkbookStyleEdit['verticalAlignment']> = {
  1: 'top',
  2: 'center',
  3: 'bottom',
}

/// Converts a Univer IStyleData delta (from a set-range-values mutation) to
/// the renderer-neutral wire format. Unknown keys are ignored; `null` values
/// mean "remove the attribute" and map to explicit `false`.
export function toNeutralStyle(s: Record<string, unknown>): WorkbookStyleEdit | undefined {
  const style: Record<string, unknown> = {}
  if ('bl' in s) style.bold = s.bl === 1
  if ('it' in s) style.italic = s.it === 1
  if ('ul' in s) {
    style.underline = isLineOn(s.ul)
    // Univer TextDecoration.DOUBLE = 10; anything else saves as single.
    if (style.underline) {
      style.underlineStyle = (s.ul as Record<string, unknown>).t === 10 ? 'double' : 'single'
    }
  }
  if ('st' in s) style.strikethrough = isLineOn(s.st)
  if (typeof s.ff === 'string' && s.ff.length > 0) style.fontFamily = s.ff
  if (typeof s.fs === 'number' && Number.isFinite(s.fs) && s.fs > 0) style.fontSize = s.fs
  if ('cl' in s && s.cl === null) {
    style.fontColor = null
  } else {
    const fontColor = extractRgb(s.cl)
    if (fontColor) style.fontColor = fontColor
  }
  if ('bg' in s) {
    if (s.bg === null) style.fillColor = null
    else {
      const fillColor = extractRgb(s.bg)
      if (fillColor) style.fillColor = fillColor
    }
  }
  if (typeof s.ht === 'number' && UNIVER_HORIZONTAL[s.ht]) {
    style.horizontalAlignment = UNIVER_HORIZONTAL[s.ht]
  }
  if (typeof s.vt === 'number' && UNIVER_VERTICAL[s.vt]) {
    style.verticalAlignment = UNIVER_VERTICAL[s.vt]
  }
  if ('tb' in s && typeof s.tb === 'number') style.wrapText = s.tb === 3
  if ('tr' in s) {
    // Univer angles are counterclockwise-positive; OOXML stores 0-90 up,
    // 91-180 down (value-90), 255 stacked.
    const tr = s.tr as { a?: unknown; v?: unknown } | null
    if (tr === null) style.textRotation = 0
    else if (tr.v === 1) style.textRotation = 255
    else if (typeof tr.a === 'number' && Number.isFinite(tr.a)) {
      const angle = Math.max(-90, Math.min(90, Math.round(tr.a)))
      style.textRotation = angle >= 0 ? angle : 90 - angle
    }
  }
  if ('pd' in s) {
    // Indent renders as left padding (INDENT_STEP_PX per step); the journal
    // stores the OOXML step count. pd: null clears the indent.
    if (s.pd === null) style.indent = 0
    else if (typeof s.pd === 'object') {
      const left = (s.pd as Record<string, unknown>).l
      if (typeof left === 'number' && Number.isFinite(left)) {
        style.indent = Math.min(250, Math.max(0, Math.round(left / INDENT_STEP_PX)))
      }
    }
  }
  const pattern =
    typeof s.n === 'object' && s.n !== null ? (s.n as Record<string, unknown>).pattern : undefined
  if (typeof pattern === 'string' && pattern.length > 0 && pattern.length <= 255) {
    style.numberFormat = pattern
  }
  if ('bd' in s && typeof s.bd === 'object' && s.bd !== null) {
    const bd = s.bd as Record<string, unknown>
    for (const [univerKey, neutralKey] of BORDER_EDGE_KEYS) {
      if (!(univerKey in bd)) continue
      if (bd[univerKey] === null) {
        style[neutralKey] = null
        continue
      }
      const edge = univerBorderEdgeToNeutral(bd[univerKey])
      if (edge) style[neutralKey] = edge
    }
  }
  return Object.keys(style).length === 0 ? undefined : (style as WorkbookStyleEdit)
}

function isLineOn(value: unknown): boolean {
  return typeof value === 'object' && value !== null && (value as Record<string, unknown>).s === 1
}

function extractRgb(value: unknown): string | undefined {
  if (typeof value !== 'object' || value === null) return undefined
  const rgb = (value as Record<string, unknown>).rgb
  if (typeof rgb !== 'string') return undefined
  const match = /^#?([0-9a-fA-F]{6})$/.exec(rgb)
  return match?.[1] ? `#${match[1].toUpperCase()}` : undefined
}

const XLSX_HORIZONTAL_TO_UNIVER: Record<string, number> = {
  left: 1,
  center: 2,
  right: 3,
  justify: 4,
  distributed: 6,
}
const XLSX_VERTICAL_TO_UNIVER: Record<string, number> = {
  top: 1,
  center: 2,
  bottom: 3,
}

/// Converts a neutral style delta back to a Univer IStyleData patch, used to
/// re-apply journaled style edits after a streamed viewport re-install.
export function fromNeutralStyle(style: WorkbookStyleEdit): Record<string, unknown> {
  const s: Record<string, unknown> = {}
  if (style.bold !== undefined) s.bl = style.bold ? 1 : null
  if (style.italic !== undefined) s.it = style.italic ? 1 : null
  if (style.underline !== undefined) {
    s.ul = style.underline
      ? { s: 1, ...(style.underlineStyle === 'double' ? { t: 10 } : {}) }
      : null
  }
  if (style.strikethrough !== undefined) s.st = style.strikethrough ? { s: 1 } : null
  if (style.fontFamily !== undefined) s.ff = style.fontFamily
  if (style.fontSize !== undefined) s.fs = style.fontSize
  if (style.fontColor !== undefined) {
    s.cl = style.fontColor === null ? null : { rgb: style.fontColor }
  }
  if (style.fillColor !== undefined) {
    s.bg = style.fillColor === null ? null : { rgb: style.fillColor }
  }
  if (style.horizontalAlignment !== undefined) {
    s.ht = XLSX_HORIZONTAL_TO_UNIVER[style.horizontalAlignment]
  }
  if (style.verticalAlignment !== undefined) {
    s.vt = XLSX_VERTICAL_TO_UNIVER[style.verticalAlignment]
  }
  if (style.wrapText !== undefined) s.tb = style.wrapText ? 3 : null
  if (style.textRotation !== undefined) {
    s.tr =
      style.textRotation === 0
        ? null
        : style.textRotation === 255
          ? { a: 0, v: 1 }
          : style.textRotation > 90
            ? { a: 90 - style.textRotation }
            : { a: style.textRotation }
  }
  if (style.numberFormat !== undefined) s.n = { pattern: style.numberFormat }
  if (style.indent !== undefined) {
    s.pd = style.indent === 0 ? null : { l: style.indent * INDENT_STEP_PX }
  }
  const bd: Record<string, unknown> = {}
  for (const [univerKey, neutralKey] of BORDER_EDGE_KEYS) {
    const edge = style[neutralKey]
    if (edge === undefined) continue
    bd[univerKey] =
      edge === null
        ? null
        : {
            s: XLSX_BORDER_TO_UNIVER[edge.style] ?? 1,
            cl: { rgb: edge.color ?? '#000000' },
          }
  }
  if (Object.keys(bd).length > 0) s.bd = bd
  return s
}

export function journalEntriesInRange(
  journal: EditJournal,
  sheetId: string,
  range: CellRange,
): JournalEntry[] {
  const sheetEntries = journal.cells.get(sheetId)
  if (!sheetEntries) return []
  const matches: JournalEntry[] = []
  for (const entry of sheetEntries.values()) {
    if (
      entry.row >= range.startRow &&
      entry.row <= range.endRow &&
      entry.column >= range.startColumn &&
      entry.column <= range.endColumn
    ) {
      matches.push(entry)
    }
  }
  return matches
}

/// Counts what a save would actually write, so entries on removed sheets
/// don't inflate the pending-edit badge.
export function journalSize(journal: EditJournal): number {
  let total = journal.chartEdits.size + sheetOpCount(journal)
  for (const visual of journal.visualAdds) {
    if (!isSheetRemoved(journal, visual.sheetId)) total += 1
  }
  for (const edit of journal.visualEdits.values()) {
    if (!isSheetRemoved(journal, edit.sheetId)) total += 1
  }
  for (const table of journal.tableAdds) {
    if (!isSheetRemoved(journal, table.sheetId)) total += 1
  }
  for (const pivot of journal.pivotAdds) {
    if (!isSheetRemoved(journal, pivot.sheetId) && !isSheetRemoved(journal, pivot.sourceSheetId)) {
      total += 1
    }
  }
  for (const sparkline of journal.sparklineAdds) {
    if (!isSheetRemoved(journal, sparkline.sheetId)) total += 1
  }
  for (const sheetId of journal.filterDirty) {
    if (!isSheetRemoved(journal, sheetId)) total += 1
  }
  for (const sheetId of journal.cfDirty) {
    if (!isSheetRemoved(journal, sheetId)) total += 1
  }
  for (const sheetId of journal.dvDirty) {
    if (!isSheetRemoved(journal, sheetId)) total += 1
  }
  for (const sheetId of journal.sheetProtection.keys()) {
    if (!isSheetRemoved(journal, sheetId)) total += 1
  }
  if (journal.definedNames.dirty) total += 1
  for (const [sheetId, state] of journal.pageSetup) {
    if (!isSheetRemoved(journal, sheetId) && Object.keys(state).length > 0) total += 1
  }
  for (const sheetId of journal.noteDirty) {
    if (!isSheetRemoved(journal, sheetId)) total += 1
  }
  total += journal.pivotCacheRefresh.size
  for (const [sheetId, links] of journal.hyperlinks) {
    if (!isSheetRemoved(journal, sheetId)) total += links.size
  }
  for (const [sheetId, ops] of journal.structuralOps) {
    if (!isSheetRemoved(journal, sheetId)) total += ops.length
  }
  for (const [sheetId, sheetEntries] of journal.cells) {
    if (!isSheetRemoved(journal, sheetId)) total += sheetEntries.size
  }
  return total
}

/// User input for the recalc engine: formulas verbatim, booleans
/// as TRUE/FALSE, and text guarded against reinterpretation as a formula.
export function toRecalcUserInput(entry: JournalEntry): string {
  if (entry.formula) return entry.formula
  const { value } = entry
  if (value === null) return ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (typeof value === 'number') return String(value)
  return value.startsWith('=') || value.startsWith("'") ? `'${value}` : value
}

export function toSaveEdits(journal: EditJournal): WorkbookCellEdit[] {
  const edits: WorkbookCellEdit[] = []
  for (const [sheetId, sheetEntries] of journal.cells) {
    if (isSheetRemoved(journal, sheetId)) continue
    for (const entry of sheetEntries.values()) {
      edits.push({
        sheetId,
        row: entry.row,
        column: entry.column,
        writeValue: entry.hasValue,
        value: entry.value,
        ...(entry.formula === undefined ? {} : { formula: entry.formula }),
        ...(entry.style === undefined ? {} : { style: entry.style }),
        ...(entry.rich === undefined ? {} : { rich: [...entry.rich] }),
        ...(entry.styleReset ? { styleReset: true } : {}),
      })
    }
  }
  return edits
}
