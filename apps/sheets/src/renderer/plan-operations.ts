/**
 * AI change-plan builders for the sheets renderer.
 *
 * proposeOperations validates agent-provided DSL operations and builds a
 * preview plan; runDeterministicPlan does the same for the regex planner.
 * Extracted from App.tsx; App-scope state comes in through PlanContext.
 */
import { planPrompt } from '../ai/deterministic-planner'
import {
  columnIndex,
  columnLabel,
  formatAddress,
  parseAddress,
  parseRange,
  rangeCellCount,
} from '../domain/cell-address'
import { CHART_EDIT_TYPES, chartDataFromValues } from '../domain/chart-visual'
import type { InMemoryWorkbookAdapter } from '../domain/in-memory-workbook'
import {
  convertToValuesBatchError,
  copyTargetBounds,
  expandToPrimitiveOps,
  isLayoutOp,
  isStructuralOp,
  MAX_EXPANDED_CELL_OPS,
  workbookCommandBatchSchema,
  type PrimitiveOperation,
  type WorkbookOperation,
} from '../domain/workbook-dsl'
import type { ApplyOutcome, ChangePlan } from '../domain/workbook.types'
import { isSheetRemoved } from './edit-journal'
import { fillFormulaCostError, quadraticFormulaError, type FormulaCostSheet } from './formula-cost'
import { t } from './i18n/locale'
import { buildLazyChangePlan } from './lazy-plan'
import {
  lazyCellEditable,
  lazyRangeEditable,
  lazyWorkbookCellReader,
  normalizeLinkTarget,
  protectSheetGuard,
} from './univer-sync'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'
import { convertibleType } from './WorkbookVisuals'

/** App-scope state the plan builders need, threaded explicitly. */
export interface PlanContext {
  readonly adapterRef: { readonly current: InMemoryWorkbookAdapter }
  readonly univerRef: { readonly current: UniverRuntime | null }
  readonly lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  readonly lazyPreviewRef: {
    current: { sessionId: string; sheetId: string; plan: ChangePlan } | null
  }
  readonly setPreview: (plan: ChangePlan | null) => void
  readonly autoApplySafePlan: (plan: ChangePlan) => Promise<ApplyOutcome>
}

/** shared by the agent's propose_operations tool; identical validation and
 * CAS/streaming-guard checks as handlePlan/handleLazyPlan, just fed
 * AI-provided operations instead of the regex planner's output. */
export function proposeOperations(
  ctx: PlanContext,
  operations: readonly WorkbookOperation[],
  summary: string,
): { ok: true; plan: ChangePlan; applied: Promise<ApplyOutcome> } | { ok: false; error: string } {
  const state = ctx.lazyWorkbookRef.current
  if (state) {
    const worksheet = ctx.univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!worksheet) return { ok: false, error: 'No workbook is open.' }
    const sheetId = worksheet.getSheetId()
    try {
      const batch = workbookCommandBatchSchema.parse({
        dslVersion: 1,
        transactionId: `agent-${crypto.randomUUID()}`,
        baseRevision: 0,
        summary,
        operations,
      })
      // find_replace plans against the grid; unloaded cells read as empty
      // and would silently miss matches — fail early instead. Ranges above
      // the per-cell expansion cap are exempt: they stay range-level and the
      // apply executor loads each chunk before scanning it.
      for (const operation of batch.operations) {
        if (operation.op !== 'find_replace') continue
        const bounds = parseRange(operation.range)
        if (rangeCellCount(bounds) > MAX_EXPANDED_CELL_OPS) continue
        const loaded = state.loadedRanges.get(operation.sheetId)
        const rangeLoaded =
          state.formulaMode ||
          (loaded !== undefined &&
            bounds.startRow >= loaded.startRow &&
            bounds.endRow <= loaded.endRow &&
            bounds.startColumn >= loaded.startColumn &&
            bounds.endColumn <= loaded.endColumn)
        if (!rangeLoaded) {
          return {
            ok: false,
            error:
              'The find_replace range is not fully loaded yet — narrow it to a loaded region, or read_range first and rewrite precisely with set_range.',
          }
        }
      }
      const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
      if (!workbook) return { ok: false, error: 'No workbook is open.' }
      const reader = lazyWorkbookCellReader(workbook)
      for (const operation of expandToPrimitiveOps(batch.operations, reader)) {
        // Every sheet-addressed op must reference an existing sheet BEFORE the
        // batch starts applying: apply routes through sheetById and a mid-batch
        // throw would leave earlier ops committed while the tool reports the
        // workbook unchanged.
        if (
          'sheetId' in operation &&
          typeof operation.sheetId === 'string' &&
          !workbook.getSheetBySheetId(operation.sheetId)
        ) {
          return {
            ok: false,
            error: `Unknown sheet: ${operation.sheetId} (use an id from get_workbook_context)`,
          }
        }
        const gateError = lazyGateError(state, operation)
        if (gateError) return { ok: false, error: gateError }
        if (operation.op === 'edit_chart') {
          const visual = [...state.file.visuals, ...state.editJournal.visualAdds].find(
            (candidate) =>
              candidate.chartPath === operation.chartPath || candidate.id === operation.chartPath,
          )
          if (!visual) {
            return { ok: false, error: `Unknown chart: ${operation.chartPath}` }
          }
          // Save-time chart patching fails closed on non-convertible plots;
          // reject here so the user never sees Apply succeed and ⌘S fail.
          if (
            operation.chartType !== undefined &&
            (!visual.chart || convertibleType(visual.chart) === null)
          ) {
            return {
              ok: false,
              error: `Chart ${operation.chartPath} cannot be converted to another type (only single-plot column/bar/line/area/pie/doughnut charts can).`,
            }
          }
          if (
            operation.axisTitles !== undefined &&
            visual.chart?.chartTypes.some((type) => /pie|doughnut/i.test(type))
          ) {
            return {
              ok: false,
              error: 'Pie/doughnut charts have no axes — axisTitles does not apply.',
            }
          }
          if (operation.grouping !== undefined) {
            const targetTypes =
              operation.chartType !== undefined
                ? (CHART_EDIT_TYPES[operation.chartType]?.chartTypes ?? [])
                : (visual.chart?.chartTypes ?? [])
            if (!targetTypes.some((type) => /^(barChart|lineChart|areaChart)$/.test(type))) {
              return { ok: false, error: 'Only bar, line, and area charts support grouping.' }
            }
          }
          for (const entry of operation.seriesData ?? []) {
            const seriesCount = visual.chart?.series.length ?? 0
            if (entry.index >= seriesCount) {
              return {
                ok: false,
                error: `Chart ${operation.chartPath} has no series #${entry.index} (it has ${seriesCount}).`,
              }
            }
            if (entry.sheetId !== undefined && !workbook?.getSheetBySheetId(entry.sheetId)) {
              return { ok: false, error: `Unknown sheet: ${entry.sheetId}` }
            }
            for (const vector of [entry.valuesRange, entry.categoriesRange]) {
              if (vector === undefined) continue
              const bounds = parseRange(vector)
              if (bounds.startRow !== bounds.endRow && bounds.startColumn !== bounds.endColumn) {
                return {
                  ok: false,
                  error: `${vector} must be a single row or a single column of cells.`,
                }
              }
              if (rangeCellCount(bounds) > 1000) {
                return { ok: false, error: `${vector} covers more than 1000 cells.` }
              }
            }
          }
          continue
        }
        if (operation.op === 'edit_shape') {
          const visual = state.editJournal.visualAdds.find(
            (candidate) => candidate.id === operation.visualId,
          )
          if (!visual || visual.kind !== 'shape') {
            return {
              ok: false,
              error:
                `No editable shape "${operation.visualId}" — only shapes added this session can be edited` +
                ' (ids come from read_sheet_features); shapes that came with the file cannot be modified.',
            }
          }
          continue
        }
        if (operation.op === 'delete_visual') {
          const exists = [...state.file.visuals, ...state.editJournal.visualAdds].some(
            (candidate) =>
              candidate.id === operation.visualId || candidate.chartPath === operation.visualId,
          )
          if (!exists) return { ok: false, error: `Unknown visual: ${operation.visualId}` }
          continue
        }
        if (operation.op === 'delete_table') {
          const exists = state.editJournal.tableAdds.some(
            (table) =>
              table.sheetId === operation.sheetId &&
              table.name.toLowerCase() === operation.tableName.toLowerCase(),
          )
          if (!exists) {
            return {
              ok: false,
              error: `Table "${operation.tableName}" does not exist or was not created this session — tables that came with the file cannot be deleted yet.`,
            }
          }
          continue
        }
        if (
          operation.op === 'add_chart' ||
          operation.op === 'add_shape' ||
          operation.op === 'add_image'
        ) {
          if (
            operation.op === 'add_image' &&
            !/^https?:\/\//i.test(operation.path) &&
            !/\.(png|jpe?g|gif)$/i.test(operation.path)
          ) {
            return {
              ok: false,
              error:
                'Only PNG/JPEG/GIF images are supported (judged by extension; URLs are validated on download).',
            }
          }
          const targetSheet = workbook?.getSheetBySheetId(operation.sheetId)
          if (!targetSheet || isSheetRemoved(state.editJournal, operation.sheetId)) {
            return { ok: false, error: `Unknown sheet: ${operation.sheetId}` }
          }
          if (operation.op === 'add_chart') {
            const bounds = parseRange(operation.dataRange)
            if (rangeCellCount(bounds) > 2000) {
              return { ok: false, error: 'add_chart dataRange covers more than 2000 cells.' }
            }
            // The grid only holds streamed-in cells, so an empty read is a
            // data problem only when the range is actually loaded; apply
            // reads the real values through the sidecar either way.
            const loaded = state.loadedRanges.get(operation.sheetId)
            const rangeLoaded =
              state.formulaMode ||
              (loaded !== undefined &&
                bounds.startRow >= loaded.startRow &&
                bounds.endRow <= loaded.endRow &&
                bounds.startColumn >= loaded.startColumn &&
                bounds.endColumn <= loaded.endColumn)
            const values = targetSheet.getRange(operation.dataRange).getRawValues() as (
              string | number | boolean | null | undefined
            )[][]
            if (rangeLoaded && !chartDataFromValues(values)) {
              return {
                ok: false,
                error: 'The chart dataRange needs at least one numeric column.',
              }
            }
          }
          continue
        }
        if (operation.op === 'add_pivot') {
          // Aggregation reads the on-screen grid; a partially streamed source
          // would silently produce wrong totals — fail closed like refresh_pivot.
          if (!state.formulaMode || !state.flags.preloadComplete) {
            return {
              ok: false,
              error:
                'add_pivot needs the fully-loaded mode — this workbook is streamed in partially. ' +
                'Build a formula aggregation table instead (SUMIFS fallback in the pivot guide), ' +
                'and tell the user the result is a formula summary, not a native pivot table.',
            }
          }
          continue
        }
        if (operation.op === 'refresh_pivot') {
          const sheetMeta = state.file.sheets.find((sheet) => sheet.id === operation.sheetId)
          if (!sheetMeta) return { ok: false, error: `Unknown sheet: ${operation.sheetId}` }
          if (sheetMeta.pivotTables.length === 0) {
            return { ok: false, error: 'This sheet has no pivot tables.' }
          }
          if (!state.formulaMode || !state.flags.preloadComplete) {
            return {
              ok: false,
              error:
                'Refreshing pivot tables needs the fully-loaded mode — this workbook is too large and was only streamed in.',
            }
          }
          for (const pivot of sheetMeta.pivotTables) {
            const definition = state.pivotDefinitions.get(pivot.path)
            if (definition && definition.unsupported.length > 0) {
              return {
                ok: false,
                error: `Pivot table ${pivot.path} does not support recompute: ${definition.unsupported.join('; ')}`,
              }
            }
          }
          continue
        }
        if (operation.op === 'set_hyperlink') {
          if (operation.target !== null && normalizeLinkTarget(operation.target) === null) {
            return {
              ok: false,
              error:
                'set_hyperlink target must be a URL (https://…) or a sheet reference like Sheet1!A1.',
            }
          }
          continue
        }
        if (operation.op === 'protect_sheet') {
          const guard = protectSheetGuard(state, operation.sheetId, operation.protected)
          if (guard) return { ok: false, error: guard }
          continue
        }
        if (operation.op === 'duplicate_sheet') {
          const isAdded = state.editJournal.sheets.added.has(operation.sheetId)
          if (!isAdded && (!state.formulaMode || !state.flags.preloadComplete)) {
            return {
              ok: false,
              error:
                'Duplicating a sheet needs the fully-loaded mode — this workbook is too large and streams partially.',
            }
          }
          const sheetMeta = state.file.sheets.find((sheet) => sheet.id === operation.sheetId)
          if (sheetMeta && sheetMeta.pivotRanges.length > 0) {
            return {
              ok: false,
              error: 'This sheet contains a PivotTable — duplicating it is not supported yet.',
            }
          }
          continue
        }
        if (
          operation.op === 'fill_range' ||
          operation.op === 'clear_range' ||
          operation.op === 'copy_range' ||
          operation.op === 'convert_to_values' ||
          operation.op === 'find_replace'
        ) {
          // Range-level bulk ops (fill / copy / convert / large clear and
          // find_replace): validate the whole target rectangle at once — the
          // per-cell checks below would never see these because they are not
          // expanded. (find_replace only arrives here above the expansion
          // cap.) Unloaded target regions are fine: the apply executor loads
          // them chunk by chunk.
          const sheetMeta = state.file.sheets.find((sheet) => sheet.id === operation.sheetId)
          const targetSheet = workbook?.getSheetBySheetId(operation.sheetId)
          if (!targetSheet || isSheetRemoved(state.editJournal, operation.sheetId)) {
            return { ok: false, error: `Unknown sheet: ${operation.sheetId}` }
          }
          const bounds =
            operation.op === 'fill_range'
              ? parseRange(operation.target)
              : operation.op === 'copy_range'
                ? copyTargetBounds(operation)
                : parseRange(operation.range)
          if (
            bounds.endRow >= targetSheet.getMaxRows() ||
            bounds.endColumn >= targetSheet.getMaxColumns()
          ) {
            return {
              ok: false,
              error:
                `The target range extends beyond the sheet grid (${targetSheet.getMaxRows()} rows × ` +
                `${targetSheet.getMaxColumns()} columns) — stay within it, or insert rows/columns first.`,
            }
          }
          if (
            sheetMeta?.pivotRanges.some(
              (range) =>
                bounds.startRow <= range.endRow &&
                bounds.endRow >= range.startRow &&
                bounds.startColumn <= range.endColumn &&
                bounds.endColumn >= range.startColumn,
            )
          ) {
            return {
              ok: false,
              error:
                'The target range overlaps a pivot table output region; those cells are read-only. ' +
                'If the source data changed, recompute with refresh_pivot instead.',
            }
          }
          if (operation.op === 'copy_range') {
            // Source loadedness is NOT required: the executor reads it chunk
            // by chunk. No fill-style cost guard either — every copied
            // formula evaluates exactly once, the same load the originals
            // already impose on the engine.
            const sourceSheetId = operation.sourceSheetId ?? operation.sheetId
            if (
              !workbook?.getSheetBySheetId(sourceSheetId) ||
              isSheetRemoved(state.editJournal, sourceSheetId)
            ) {
              return { ok: false, error: `Unknown sheet: ${sourceSheetId}` }
            }
          }
          if (operation.op === 'fill_range') {
            const sourceSheetId = operation.sourceSheetId ?? operation.sheetId
            const sourceSheet = workbook?.getSheetBySheetId(sourceSheetId)
            if (!sourceSheet || isSheetRemoved(state.editJournal, sourceSheetId)) {
              return { ok: false, error: `Unknown sheet: ${sourceSheetId}` }
            }
            const sourceBounds = parseRange(operation.source)
            // The source must already be in the grid: its real contents are
            // copied (and its formulas cost-checked) synchronously here.
            if (!lazyRangeEditable(state, sourceSheetId, sourceBounds)) {
              return {
                ok: false,
                error:
                  'The fill source is not loaded yet — read_range the source first (that loads it), then retry the fill.',
              }
            }
            // Each source formula evaluates once per filled copy; reject
            // fills whose total cost would freeze the formula engine.
            const copies = rangeCellCount(bounds) / rangeCellCount(sourceBounds)
            const hostName =
              state.file.sheets.find((sheet) => sheet.id === operation.sheetId)?.name ??
              workbook?.getSheetBySheetId(operation.sheetId)?.getSheetName() ??
              ''
            for (let row = sourceBounds.startRow; row <= sourceBounds.endRow; row += 1) {
              for (
                let column = sourceBounds.startColumn;
                column <= sourceBounds.endColumn;
                column += 1
              ) {
                const formula = sourceSheet.getRange(formatAddress(row, column)).getFormula()
                if (!formula) continue
                const costError = fillFormulaCostError(
                  formula,
                  copies,
                  bounds.endRow - sourceBounds.endRow,
                  bounds.endColumn - sourceBounds.endColumn,
                  hostName,
                  lazyFormulaCostSheets(state),
                )
                if (costError) return { ok: false, error: costError }
              }
            }
          }
          continue
        }
        if (
          operation.op === 'rename_sheet' ||
          operation.op === 'format_range' ||
          isLayoutOp(operation) ||
          isStructuralOp(operation)
        ) {
          // Row/column inserts and deletes must anchor on an existing grid
          // line: Univer rejects out-of-bounds anchors AFTER the plan is
          // accepted — the command silently no-ops and the grid pops a
          // misleading "range is protected" permission dialog. Fail here
          // with an actionable error instead.
          if (
            operation.op === 'insert_rows' ||
            operation.op === 'delete_rows' ||
            operation.op === 'insert_cols' ||
            operation.op === 'delete_cols'
          ) {
            const targetSheet = workbook?.getSheetBySheetId(operation.sheetId)
            if (!targetSheet || isSheetRemoved(state.editJournal, operation.sheetId)) {
              return { ok: false, error: `Unknown sheet: ${operation.sheetId}` }
            }
            const maxRows = targetSheet.getMaxRows()
            const maxColumns = targetSheet.getMaxColumns()
            const lastColumn = columnLabel(maxColumns - 1)
            if (operation.op === 'insert_rows' && operation.row > maxRows) {
              return {
                ok: false,
                error:
                  `insert_rows: row=${operation.row} is beyond the sheet grid (${maxRows} rows) — ` +
                  `rows can only be inserted before an existing row (1–${maxRows}). ` +
                  'To add data below the last row, write into the empty rows instead.',
              }
            }
            if (operation.op === 'delete_rows' && operation.row + operation.count - 1 > maxRows) {
              return {
                ok: false,
                error: `delete_rows: rows ${operation.row}–${operation.row + operation.count - 1} extend beyond the sheet grid (${maxRows} rows).`,
              }
            }
            if (operation.op === 'insert_cols' && columnIndex(operation.column) >= maxColumns) {
              return {
                ok: false,
                error:
                  `insert_cols: column ${operation.column} is beyond the sheet grid (last column ${lastColumn}) — ` +
                  `columns can only be inserted before an existing column (A–${lastColumn}). ` +
                  `To add a new column near the right edge, insert before the last column ${lastColumn} ` +
                  '(the current last column shifts right and stays last).',
              }
            }
            if (
              operation.op === 'delete_cols' &&
              columnIndex(operation.column) + operation.count > maxColumns
            ) {
              return {
                ok: false,
                error: `delete_cols: columns ${operation.column} + count=${operation.count} extend beyond the sheet grid (last column ${lastColumn}).`,
              }
            }
            continue
          }
          // add_table_row/col and delete_table_row/col are layout ops, but also
          // need table existence + range validity checks here (fail-closed).
          if (
            operation.op === 'add_table_row' ||
            operation.op === 'add_table_column' ||
            operation.op === 'delete_table_row' ||
            operation.op === 'delete_table_column'
          ) {
            const tableEntry = state.editJournal.tableAdds.find(
              (t) =>
                t.sheetId === operation.sheetId &&
                t.name.toLowerCase() === operation.tableName.toLowerCase(),
            )
            if (!tableEntry) {
              return {
                ok: false,
                error:
                  `Table "${operation.tableName}" does not exist in this session. Only tables created via add_table can be modified — ` +
                  'for tables that came with the file, save and reopen before modifying.',
              }
            }
            const dataRows = tableEntry.area.endRow - tableEntry.area.startRow
            if (operation.op === 'add_table_row') {
              const insertRow = operation.row ?? dataRows + 1
              if (insertRow < 1 || insertRow > dataRows + 1) {
                return {
                  ok: false,
                  error: `add_table_row: row=${insertRow} is out of range (the data area has ${dataRows} rows; valid insert positions are 1–${dataRows + 1}).`,
                }
              }
            }
            if (operation.op === 'delete_table_row') {
              const { row, count = 1 } = operation
              if (row < 1 || row + count - 1 > dataRows) {
                return {
                  ok: false,
                  error: `delete_table_row: row=${row} count=${count} is outside the table data area (${dataRows} rows total).`,
                }
              }
              if (dataRows - count < 1) {
                return {
                  ok: false,
                  error:
                    'delete_table_row: at least 1 data row must remain — cannot delete them all.',
                }
              }
            }
            const tableCols = tableEntry.area.endColumn - tableEntry.area.startColumn + 1
            if (operation.op === 'add_table_column') {
              const insertCol = operation.column ?? tableCols + 1
              if (insertCol < 1 || insertCol > tableCols + 1) {
                return {
                  ok: false,
                  error: `add_table_column: column=${insertCol} is out of range (${tableCols} columns total).`,
                }
              }
              if (
                tableEntry.columnNames.some(
                  (n) => n.toLowerCase() === operation.columnName.toLowerCase(),
                )
              ) {
                return {
                  ok: false,
                  error: `add_table_column: table "${operation.tableName}" already has a column "${operation.columnName}" — column names must be unique.`,
                }
              }
            }
            if (operation.op === 'delete_table_column') {
              const { column, count = 1 } = operation
              if (column < 1 || column + count - 1 > tableCols) {
                return {
                  ok: false,
                  error: `delete_table_column: column=${column} count=${count} is out of range (${tableCols} columns total).`,
                }
              }
              if (tableCols - count < 1) {
                return {
                  ok: false,
                  error:
                    'delete_table_column: at least 1 column must remain — cannot delete them all.',
                }
              }
            }
          }
          continue
        }
        // Quadratic array-criteria formulas would freeze Univer's main-thread
        // engine (distinct-count COUNTIF idioms over 80k+ rows take minutes).
        if (operation.op === 'set_formula') {
          const hostName =
            state.file.sheets.find((sheet) => sheet.id === operation.sheetId)?.name ??
            workbook?.getSheetBySheetId(operation.sheetId)?.getSheetName() ??
            ''
          const costError = quadraticFormulaError(
            operation.formula,
            hostName,
            lazyFormulaCostSheets(state),
          )
          if (costError) return { ok: false, error: costError }
        }
        const target = parseAddress(operation.address)
        // Pivot output is baked into the worksheet (same guard as the cell
        // editor); the AI apply path bypasses the editor so check here.
        const sheetMeta = state.file.sheets.find((sheet) => sheet.id === operation.sheetId)
        if (
          sheetMeta?.pivotRanges.some(
            (range) =>
              target.row >= range.startRow &&
              target.row <= range.endRow &&
              target.column >= range.startColumn &&
              target.column <= range.endColumn,
          )
        ) {
          return {
            ok: false,
            error:
              `${operation.address} is inside a pivot table output region; those cells are read-only. ` +
              'If the source data changed, recompute with refresh_pivot; for a new pivot analysis, build one in a blank area with add_pivot.',
          }
        }
        if (!lazyCellEditable(state, operation.sheetId, target.row, target.column)) {
          return { ok: false, error: 'That cell is still streaming in — try again in a moment.' }
        }
      }
      const plan = buildLazyChangePlan(batch, reader, (id) => {
        const sheet = workbook.getSheetBySheetId(id)
        if (!sheet) throw new Error(`Unknown sheet: ${id}`)
        return sheet.getSheetName()
      })
      ctx.lazyPreviewRef.current = { sessionId: state.file.sessionId, sheetId, plan }
      ctx.setPreview(plan)
      // All plans auto-apply (undo covers them); the caller awaits `applied`
      // so a failed apply is reported instead of silently claimed as done.
      return { ok: true, plan, applied: ctx.autoApplySafePlan(plan) }
    } catch (error: unknown) {
      return {
        ok: false,
        error: error instanceof Error ? error.message : 'Unable to create a preview.',
      }
    }
  }
  try {
    const snapshot = ctx.adapterRef.current.getSnapshot()
    const demoSheets = demoFormulaCostSheets(snapshot.sheets)
    // convert_to_values freezes each formula cell to its COMPUTED value,
    // which only the live grid knows (the demo snapshot stores what was
    // written, not what formulas evaluate to) — pre-expand it into plain
    // set_cell writes against the grid here. Same-batch formula writes into
    // the convert range would be invisible to that read; reject the mix
    // (the lazy path enforces this inside expandToPrimitiveOps).
    const convertBatchError = convertToValuesBatchError(operations)
    if (convertBatchError) return { ok: false, error: convertBatchError }
    const demoWorkbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
    const prepared: WorkbookOperation[] = []
    for (const operation of operations) {
      if (operation.op !== 'convert_to_values') {
        prepared.push(operation)
        continue
      }
      const bounds = parseRange(operation.range)
      if (rangeCellCount(bounds) > MAX_EXPANDED_CELL_OPS) {
        return {
          ok: false,
          error: `convert_to_values on an in-memory workbook is limited to ${MAX_EXPANDED_CELL_OPS} cells per operation (imported xlsx files allow up to 200,000).`,
        }
      }
      const snapshotSheet = snapshot.sheets.find((sheet) => sheet.id === operation.sheetId)
      if (!snapshotSheet) {
        return { ok: false, error: `Unknown sheet: ${operation.sheetId}` }
      }
      const gridSheet = demoWorkbook?.getSheetBySheetId(operation.sheetId)
      for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
        for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
          const address = formatAddress(row, column)
          if (!snapshotSheet.cells[address]?.formula) continue
          const raw = gridSheet?.getRange(address).getValue()
          const value =
            typeof raw === 'string' || typeof raw === 'number' || typeof raw === 'boolean'
              ? raw
              : null
          prepared.push({ op: 'set_cell', sheetId: operation.sheetId, address, value })
        }
      }
    }
    if (prepared.length === 0) {
      return {
        ok: false,
        error: 'The convert_to_values range contains no formula cells — nothing to convert.',
      }
    }
    for (const formulaOp of formulaOperations(prepared)) {
      const hostName = snapshot.sheets.find((sheet) => sheet.id === formulaOp.sheetId)?.name ?? ''
      const costError = quadraticFormulaError(formulaOp.formula, hostName, demoSheets)
      if (costError) return { ok: false, error: costError }
    }
    // fill_range sources live in cells, not in the batch, so the generator
    // above never sees them — cost their formulas here (the demo adapter
    // caps fills at 2000 cells, but 2000 copies of a big-range scan would
    // still freeze the engine).
    for (const operation of prepared) {
      if (operation.op !== 'fill_range') continue
      const sourceSheet = snapshot.sheets.find(
        (sheet) => sheet.id === (operation.sourceSheetId ?? operation.sheetId),
      )
      if (!sourceSheet) continue
      const sourceBounds = parseRange(operation.source)
      const targetBounds = parseRange(operation.target)
      const copies = rangeCellCount(targetBounds) / rangeCellCount(sourceBounds)
      const hostName = snapshot.sheets.find((sheet) => sheet.id === operation.sheetId)?.name ?? ''
      for (const [address, cell] of Object.entries(sourceSheet.cells)) {
        if (!cell.formula) continue
        const position = parseAddress(address)
        if (
          position.row < sourceBounds.startRow ||
          position.row > sourceBounds.endRow ||
          position.column < sourceBounds.startColumn ||
          position.column > sourceBounds.endColumn
        ) {
          continue
        }
        const costError = fillFormulaCostError(
          cell.formula,
          copies,
          targetBounds.endRow - sourceBounds.endRow,
          targetBounds.endColumn - sourceBounds.endColumn,
          hostName,
          demoSheets,
        )
        if (costError) return { ok: false, error: costError }
      }
    }
    const plan = ctx.adapterRef.current.plan({
      dslVersion: 1,
      transactionId: `agent-${crypto.randomUUID()}`,
      baseRevision: snapshot.revision,
      summary,
      operations: prepared,
    })
    ctx.setPreview(plan)
    // All plans auto-apply (undo covers them); on failure the preview
    // card stays up so the user can Apply manually.
    return { ok: true, plan, applied: ctx.autoApplySafePlan(plan) }
  } catch (error: unknown) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : 'Unable to create a preview.',
    }
  }
}

export function runDeterministicPlan(
  ctx: PlanContext,
  instruction: string,
): { text: string; isError?: boolean } {
  const state = ctx.lazyWorkbookRef.current
  if (state) {
    const runtime = ctx.univerRef.current
    const workbook = runtime?.univerAPI.getActiveWorkbook()
    const worksheet = workbook?.getActiveSheet()
    if (!runtime || !workbook || !worksheet) return { text: t('appNoWorkbookOpen'), isError: true }
    try {
      const sheetId = worksheet.getSheetId()
      // AI output stays untrusted input: it must pass the DSL schema.
      const command = workbookCommandBatchSchema.parse(
        planPrompt(instruction, { revision: 0, sheetId }),
      )
      const reader = lazyWorkbookCellReader(workbook)
      for (const operation of expandToPrimitiveOps(command.operations, reader)) {
        // Range-level ops pass through unexpanded; the apply executor loads
        // unloaded chunks itself, so no streaming check here.
        if (
          operation.op === 'fill_range' ||
          operation.op === 'copy_range' ||
          operation.op === 'convert_to_values' ||
          operation.op === 'clear_range' ||
          operation.op === 'find_replace' ||
          operation.op === 'rename_sheet' ||
          operation.op === 'format_range' ||
          isLayoutOp(operation) ||
          isStructuralOp(operation)
        )
          continue
        const target = parseAddress(operation.address)
        if (!lazyCellEditable(state, operation.sheetId, target.row, target.column)) {
          return { text: t('appCellStreaming'), isError: true }
        }
      }
      const plan = buildLazyChangePlan(command, reader, (id) => {
        const sheet = workbook.getSheetBySheetId(id)
        if (!sheet) throw new Error(`Unknown sheet: ${id}`)
        return sheet.getSheetName()
      })
      ctx.lazyPreviewRef.current = { sessionId: state.file.sessionId, sheetId, plan }
      ctx.setPreview(plan)
      void ctx.autoApplySafePlan(plan)
      return { text: t('appPreviewCreated') }
    } catch (error: unknown) {
      return {
        text: error instanceof Error ? error.message : t('appPreviewFailed'),
        isError: true,
      }
    }
  }
  try {
    const snapshot = ctx.adapterRef.current.getSnapshot()
    const activeId = ctx.univerRef.current?.univerAPI
      .getActiveWorkbook()
      ?.getActiveSheet()
      ?.getSheetId()
    const command = planPrompt(instruction, {
      revision: snapshot.revision,
      sheetId:
        snapshot.sheets.find((sheet) => sheet.id === activeId)?.id ?? snapshot.sheets[0]?.id ?? '',
    })
    const plan = ctx.adapterRef.current.plan(command)
    ctx.setPreview(plan)
    void ctx.autoApplySafePlan(plan)
    return { text: t('appPreviewCreatedDemo') }
  } catch (error: unknown) {
    return {
      text: error instanceof Error ? error.message : t('appPreviewFailed'),
      isError: true,
    }
  }
}

const PIVOT_GATED_OPS = new Set([
  'insert_rows',
  'delete_rows',
  'insert_cols',
  'delete_cols',
  'merge_cells',
  'unmerge_cells',
])
const FILTER_GATED_OPS = new Set(['set_filter', 'clear_filter', 'set_filter_criteria'])

/** Mirrors the BeforeCommandExecute gates (App.tsx) that silently cancel the
 * facade commands these ops dispatch: checked at propose time and re-checked
 * at apply time so a gated op fails with a model-facing error instead of the
 * tool reporting success on a cancelled command. */
export function lazyGateError(
  state: LazyWorkbookState,
  operation: WorkbookOperation | PrimitiveOperation,
): string | null {
  const sheetId =
    'sheetId' in operation && typeof operation.sheetId === 'string' ? operation.sheetId : undefined
  if (sheetId === undefined) return null
  if (PIVOT_GATED_OPS.has(operation.op)) {
    const sheetMeta = state.file.sheets.find((sheet) => sheet.id === sheetId)
    if (sheetMeta && sheetMeta.pivotRanges.length > 0) {
      return (
        `Sheet "${sheetMeta.name}" contains a PivotTable — row/column inserts, deletions, and merges ` +
        'are blocked there because a shift would desync the baked pivot output. ' +
        'Make the change on a sheet without pivot tables.'
      )
    }
    return null
  }
  const isAddedSheet = state.editJournal.sheets.added.has(sheetId)
  if (FILTER_GATED_OPS.has(operation.op)) {
    if (!isAddedSheet && (!state.formulaMode || !state.flags.preloadComplete)) {
      return state.formulaMode
        ? 'Filter changes need the workbook fully loaded — it is still loading; retry after loading completes.'
        : 'Filter changes need the fully-loaded mode — this workbook is too large and streams partially, so filters cannot be edited.'
    }
    if (state.filterOrigins.get(sheetId)?.origin === 'table') {
      return "This sheet's auto-filter belongs to a table — table filters cannot be edited yet."
    }
    return null
  }
  if (
    operation.op === 'set_data_validation' &&
    !isAddedSheet &&
    !state.appliedDvSheets.has(sheetId)
  ) {
    return "This sheet's data-validation rules are still being indexed — retry after workbook indexing completes."
  }
  return null
}

/** Every formula a batch would write: set_formula plus "="-strings in set_range. */
function* formulaOperations(
  operations: readonly WorkbookOperation[],
): Generator<{ sheetId: string; formula: string }> {
  for (const operation of operations) {
    if (operation.op === 'set_formula') {
      yield { sheetId: operation.sheetId, formula: operation.formula }
    } else if (operation.op === 'set_range') {
      for (const row of operation.values) {
        for (const value of row) {
          if (typeof value === 'string' && value.startsWith('=')) {
            yield { sheetId: operation.sheetId, formula: value }
          }
        }
      }
    }
  }
}

function lazyFormulaCostSheets(state: LazyWorkbookState): FormulaCostSheet[] {
  return state.file.sheets.map((sheet) => ({
    name: sheet.name,
    rows: sheet.rowCount,
    columns: sheet.columnCount,
  }))
}

/** Demo sheets carry no extent metadata; derive it from the populated cells. */
function demoFormulaCostSheets(
  sheets: readonly { name: string; cells: Readonly<Record<string, unknown>> }[],
): FormulaCostSheet[] {
  return sheets.map((sheet) => {
    let rows = 1
    let columns = 1
    for (const address of Object.keys(sheet.cells)) {
      const cell = parseAddress(address)
      rows = Math.max(rows, cell.row + 1)
      columns = Math.max(columns, cell.column + 1)
    }
    return { name: sheet.name, rows, columns }
  })
}
