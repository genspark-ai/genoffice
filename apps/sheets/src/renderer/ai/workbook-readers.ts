/**
 * Read-only workbook views for the AI skill layer (read_cells, read_formats,
 * sheet feature reports, and the active-sheet context block). Extracted from
 * App.tsx; the App component wires these into SheetsSkillDeps with a
 * WorkbookReadContext closing over its refs.
 */
import type { IRange } from '@univerjs/core'
import { columnLabel, parseAddress } from '../../domain/cell-address'
import type { InMemoryWorkbookAdapter } from '../../domain/in-memory-workbook'
import type { CellFormatState, CellScalar } from '../../domain/workbook.types'
import { toSelectionFormat } from '../selection-format'
import { lazyCellReader } from '../univer-sync'
import type { LazyWorkbookState, UniverRuntime } from '../univer-state'
import type { ActiveSheetInfo } from './tools'

/** The App refs the readers need; passed per call so they never go stale. */
export interface WorkbookReadContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  adapterRef: { readonly current: InMemoryWorkbookAdapter }
}

/// Feature-state report for the AI: what already exists before it edits.
/// Each section is independent so one facade failure never hides the rest.
export function readSheetFeatures(ctx: WorkbookReadContext, sheetIdInput?: string): string {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  if (!workbook) return 'No workbook is currently open.'
  const worksheet =
    sheetIdInput === undefined
      ? workbook.getActiveSheet()
      : workbook.getSheetBySheetId(sheetIdInput)
  if (!worksheet) return `Unknown sheet: ${sheetIdInput}`
  const sheetId = worksheet.getSheetId()
  const state = ctx.lazyWorkbookRef.current
  const a1 = (range: IRange): string =>
    `${columnLabel(range.startColumn)}${range.startRow + 1}:${columnLabel(range.endColumn)}${range.endRow + 1}`
  const lines: string[] = [`Feature state of sheet ${worksheet.getSheetName()} (id=${sheetId}):`]

  const fileMeta = state?.file.sheets.find((sheet) => sheet.id === sheetId)
  const hidden = state?.editJournal.sheets.hidden.get(sheetId) ?? fileMeta?.hidden ?? false
  const isProtected =
    state?.editJournal.sheetProtection.get(sheetId) ??
    state?.sheetProtections.get(sheetId)?.protected ??
    false
  const status = [hidden ? 'hidden' : 'visible', isProtected ? 'protected' : 'unprotected']
  try {
    const freeze = worksheet.getFreeze()
    if (freeze.ySplit > 0 || freeze.xSplit > 0) {
      status.push(`frozen: first ${freeze.ySplit} row(s), ${freeze.xSplit} column(s)`)
    }
  } catch {
    // freeze state unavailable
  }
  lines.push(`Status: ${status.join(', ')}`)

  // AutoFilter/DV install from the file only once the sheet finishes
  // indexing (captureSheetFileState marks that by filling sheetProtections),
  // so until then the grid-backed sections below under-report.
  if (fileMeta && state && !state.sheetProtections.has(sheetId)) {
    lines.push(
      '⚠️ This sheet is still indexing: the AutoFilter / conditional formatting / data validation ' +
        'sections below may be missing rules that exist in the file. Do NOT modify or clear those ' +
        'features based on this read — retry after indexing completes.',
    )
  }

  if (fileMeta && fileMeta.pivotTables.length > 0) {
    lines.push(
      `Pivot tables: ${fileMeta.pivotTables.length} (output region read-only; recompute from source data with refresh_pivot):`,
    )
    for (const pivot of fileMeta.pivotTables) {
      const definition = state?.pivotDefinitions.get(pivot.path)
      if (!definition) {
        lines.push(`- ${pivot.outputRef}: definition not loaded`)
        continue
      }
      const rows = definition.rowFields
        .filter((field) => field >= 0)
        .map((field) => definition.fields[field]?.name)
        .filter(Boolean)
      const cols = definition.colFields
        .filter((field) => field >= 0)
        .map((field) => definition.fields[field]?.name)
        .filter(Boolean)
      const values = definition.dataFields.map((dataField) => dataField.name)
      lines.push(
        `- ${pivot.outputRef}: rows[${rows.join(',')}] cols[${cols.join(',')}] values[${values.join(',')}]` +
          ` source=${definition.sourceSheet}!${definition.sourceRef}` +
          (definition.unsupported.length > 0
            ? ` (recompute unsupported: ${definition.unsupported.join('; ')})`
            : ' (recomputable)'),
      )
    }
  } else if (fileMeta && fileMeta.pivotRanges.length > 0) {
    lines.push(
      `Pivot table output regions (read-only, writes are rejected): ${fileMeta.pivotRanges.map(a1).join(', ')}`,
    )
  }

  try {
    const filter = worksheet.getFilter()
    if (filter) {
      const bounds = filter.getRange().getRange()
      const criteria: string[] = []
      for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
        const columnCriteria = filter.getColumnFilterCriteria(column)
        if (!columnCriteria) continue
        const values = (columnCriteria.filters?.filters ?? []) as string[]
        criteria.push(
          values.length > 0
            ? `${columnLabel(column)} keeps [${values.slice(0, 10).join(', ')}${values.length > 10 ? ', …' : ''}]`
            : `${columnLabel(column)} has custom criteria`,
        )
      }
      lines.push(
        `AutoFilter: ${a1(bounds)}${criteria.length > 0 ? `;${criteria.join(';')}` : ' (no column criteria)'}`,
      )
    } else {
      lines.push('AutoFilter: none')
    }
  } catch {
    lines.push('AutoFilter: read failed')
  }

  try {
    const rules = worksheet.getConditionalFormattingRules()
    if (rules.length === 0) {
      lines.push('Conditional formatting: none')
    } else {
      lines.push(
        `Conditional formatting: ${rules.length} rule(s) (clear_conditional_formats removes them all):`,
      )
      for (const rule of rules.slice(0, 20)) {
        const detail = rule.rule as unknown as Record<string, unknown>
        const summary = ['type', 'subType', 'operator', 'value']
          .filter((key) => detail[key] !== undefined)
          .map((key) => `${key}=${JSON.stringify(detail[key])}`)
          .join(' ')
        lines.push(`- ${(rule.ranges ?? []).map(a1).join(',')}: ${summary}`)
      }
      if (rules.length > 20) lines.push(`- …${rules.length - 20} more`)
    }
  } catch {
    lines.push('Conditional formatting: read failed')
  }

  try {
    const validations = worksheet.getDataValidations()
    if (validations.length === 0) {
      lines.push('Data validation: none')
    } else {
      lines.push(`Data validation: ${validations.length} rule(s):`)
      for (const validation of validations.slice(0, 20)) {
        const ranges = validation
          .getRanges()
          .map((range) => range.getA1Notation())
          .join(',')
        const rule = validation.rule
        const parts = [rule.type, rule.operator, rule.formula1, rule.formula2].filter(
          (part) => part !== undefined && part !== '',
        )
        lines.push(`- ${ranges}: ${parts.join(' ')}`)
      }
    }
  } catch {
    lines.push('Data validation: read failed')
  }

  try {
    const names = workbook.getDefinedNames()
    lines.push(
      names.length === 0
        ? 'Defined names: none'
        : `Defined names: ${names
            .slice(0, 30)
            .map((name) => `${name.getName()} = ${name.getFormulaOrRefString()}`)
            .join('; ')}`,
    )
  } catch {
    lines.push('Defined names: read failed')
  }

  try {
    const notes = worksheet.getNotes()
    if (notes.length > 0) {
      lines.push(`Notes: ${notes.length}:`)
      for (const note of notes.slice(0, 20)) {
        const text = note.note.length > 80 ? `${note.note.slice(0, 80)}…` : note.note
        lines.push(`- ${columnLabel(note.col)}${note.row + 1}: ${text.replace(/\n/g, ' ')}`)
      }
    }
  } catch {
    // notes unavailable
  }

  if (state) {
    const visuals = [...state.file.visuals, ...state.editJournal.visualAdds].filter(
      (visual) => visual.sheetId === sheetId && visual.kind !== 'chart',
    )
    if (visuals.length > 0) {
      lines.push(`Shapes/images: ${visuals.length} (charts are listed in get_workbook_context):`)
      for (const visual of visuals.slice(0, 20)) {
        const origin = visual.id.startsWith('added-shape-')
          ? `id=${visual.id}, editable via edit_shape`
          : visual.id.startsWith('added-')
            ? 'added this session'
            : 'came with the file, not modifiable by the AI'
        lines.push(
          `- ${visual.kind === 'image' ? 'image' : `shape ${visual.shapeType ?? ''}`} @ ` +
            `${columnLabel(visual.anchor.fromColumn)}${visual.anchor.fromRow + 1}` +
            `${visual.text ? ` text "${visual.text}"` : ''} (${origin})`,
        )
      }
    }
    const pageSetup = state.editJournal.pageSetup.get(sheetId)
    if (pageSetup && Object.keys(pageSetup).length > 0) {
      lines.push(`Page setup pending save this session: ${JSON.stringify(pageSetup)}`)
    }
  }
  return lines.join('\n')
}

/// Explicit formatting only: cells whose merged style resolves to all
/// defaults are omitted, so the model sees a compact "what is styled" map.
export function readFormats(
  ctx: WorkbookReadContext,
  addresses: readonly string[],
  sheetId?: string,
): Record<string, CellFormatState> {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet =
    sheetId === undefined ? workbook?.getActiveSheet() : workbook?.getSheetBySheetId(sheetId)
  if (!worksheet) return {}
  const result: Record<string, CellFormatState> = {}
  for (const address of addresses) {
    const range = worksheet.getRange(address)
    let pattern = ''
    try {
      pattern = range.getNumberFormat()
    } catch {
      // number-format resolution can fail on never-touched cells; treat as General
    }
    const echo = toSelectionFormat(range.getCellStyleData() ?? {}, pattern)
    const format: Record<string, unknown> = {}
    if (echo.bold) format.bold = true
    if (echo.italic) format.italic = true
    if (echo.underline) format.underline = true
    if (echo.strike) format.strikethrough = true
    if (echo.fontFamily) format.fontFamily = echo.fontFamily
    if (echo.fontSize) format.fontSize = echo.fontSize
    if (echo.fontColor) format.fontColor = echo.fontColor
    if (echo.fillColor) format.fillColor = echo.fillColor
    if (echo.numberFormat && echo.numberFormat !== 'General')
      format.numberFormat = echo.numberFormat
    if (
      echo.horizontalAlignment === 'left' ||
      echo.horizontalAlignment === 'center' ||
      echo.horizontalAlignment === 'right'
    ) {
      format.horizontalAlign = echo.horizontalAlignment
    }
    if (
      echo.verticalAlignment === 'top' ||
      echo.verticalAlignment === 'center' ||
      echo.verticalAlignment === 'bottom'
    ) {
      format.verticalAlign = echo.verticalAlignment
    }
    if (echo.wrap) format.wrapText = true
    if (Object.keys(format).length > 0) result[address] = format as CellFormatState
  }
  return result
}

export function getActiveSheetInfo(ctx: WorkbookReadContext): ActiveSheetInfo {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const selection = workbook?.getActiveRange()?.getA1Notation() ?? undefined
  const state = ctx.lazyWorkbookRef.current
  if (state) {
    const worksheet = workbook?.getActiveSheet()
    if (!workbook || !worksheet)
      return { mode: 'none', sheetId: '', sheetName: '', knownAddresses: [], sheets: [] }
    const sheetId = worksheet.getSheetId()
    const loaded = state.loadedRanges.get(sheetId)
    const loadedRange = loaded
      ? `${columnLabel(loaded.startColumn)}${loaded.startRow + 1}:` +
        `${columnLabel(loaded.endColumn)}${loaded.endRow + 1}`
      : undefined
    return {
      mode: 'lazy',
      sheetId,
      sheetName: worksheet.getSheetName(),
      knownAddresses: [],
      loadedRange,
      sheets: workbook.getSheets().map((sheet) => {
        const meta = state.file.sheets.find((entry) => entry.id === sheet.getSheetId())
        return {
          id: sheet.getSheetId(),
          name: sheet.getSheetName(),
          ...(meta ? { rows: meta.rowCount, columns: meta.columnCount } : {}),
        }
      }),
      selection,
      merges: worksheet.getMergedRanges().map((range) => range.getA1Notation()),
      // Session-added charts have no chart part yet; their visual id
      // doubles as the edit_chart path.
      charts: [...state.file.visuals, ...state.editJournal.visualAdds]
        .filter((visual) => visual.kind === 'chart' && (visual.chartPath || visual.chart))
        .map((visual) => ({
          path: visual.chartPath ?? visual.id,
          title: visual.chart?.title ?? '',
          types: visual.chart?.chartTypes.join('+') ?? '',
          sheetId: visual.sheetId,
        })),
    }
  }
  const snapshot = ctx.adapterRef.current.getSnapshot()
  // The adapter has no active-sheet notion — resolve it from the grid (demo
  // Univer sheets reuse the snapshot ids), falling back to the first sheet.
  const activeId = workbook?.getActiveSheet()?.getSheetId()
  const sheet = snapshot.sheets.find((entry) => entry.id === activeId) ?? snapshot.sheets[0]
  if (!sheet) return { mode: 'none', sheetId: '', sheetName: '', knownAddresses: [], sheets: [] }
  return {
    mode: 'demo',
    sheetId: sheet.id,
    sheetName: sheet.name,
    revision: snapshot.revision,
    knownAddresses: Object.keys(sheet.cells),
    sheets: snapshot.sheets.map((entry) => {
      let maxRow = -1
      let maxColumn = -1
      for (const address of Object.keys(entry.cells)) {
        const cell = parseAddress(address)
        if (cell.row > maxRow) maxRow = cell.row
        if (cell.column > maxColumn) maxColumn = cell.column
      }
      return {
        id: entry.id,
        name: entry.name,
        ...(maxRow >= 0 ? { rows: maxRow + 1, columns: maxColumn + 1 } : {}),
      }
    }),
    selection,
    merges: sheet.merges ?? [],
    charts: snapshot.sheets.flatMap((entry) =>
      (entry.visuals ?? []).map((visual) => ({
        path: visual.id,
        title: visual.chart.title,
        types: visual.chart.chartTypes.join('+'),
        sheetId: visual.sheetId,
      })),
    ),
  }
}

export function readCells(
  ctx: WorkbookReadContext,
  addresses: readonly string[],
  sheetId?: string,
): Record<string, { value: CellScalar; formula?: string }> {
  const result: Record<string, { value: CellScalar; formula?: string }> = {}
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const state = ctx.lazyWorkbookRef.current
  if (state) {
    const worksheet =
      sheetId === undefined ? workbook?.getActiveSheet() : workbook?.getSheetBySheetId(sheetId)
    if (!worksheet) return result
    const reader = lazyCellReader(worksheet)
    for (const address of addresses) {
      const cell = reader(address)
      result[address] = cell.formula
        ? { value: cell.value, formula: cell.formula }
        : { value: cell.value }
    }
    return result
  }
  const sheets = ctx.adapterRef.current.getSnapshot().sheets
  const targetId = sheetId ?? workbook?.getActiveSheet()?.getSheetId()
  const sheet =
    sheets.find((s) => s.id === targetId) ?? (sheetId === undefined ? sheets[0] : undefined)
  if (!sheet) return result
  // The in-memory model stores only value:null for formula cells; computed
  // values live in Univer's formula engine, backfilled by reading the grid
  const worksheet =
    sheetId === undefined ? workbook?.getActiveSheet() : workbook?.getSheetBySheetId(sheetId)
  for (const address of addresses) {
    const cell = sheet.cells[address] ?? { value: null }
    if (cell.formula) {
      let computed = cell.value
      if (computed === null && worksheet) {
        try {
          computed = worksheet.getRange(address).getValue() ?? null
        } catch {
          /* fail-open */
        }
      }
      result[address] = { value: computed, formula: cell.formula }
    } else {
      result[address] = { value: cell.value }
    }
  }
  return result
}
