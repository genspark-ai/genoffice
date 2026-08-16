/**
 * AI change-plan builders for the sheets renderer.
 *
 * proposeOperations validates agent-provided DSL operations and builds a
 * preview plan; runDeterministicPlan does the same for the regex planner.
 * Extracted from App.tsx; App-scope state comes in through PlanContext.
 */
import { planPrompt } from '../ai/deterministic-planner'
import { parseAddress, parseRange, rangeCellCount } from '../domain/cell-address'
import { CHART_EDIT_TYPES, chartDataFromValues } from '../domain/chart-visual'
import type { InMemoryWorkbookAdapter } from '../domain/in-memory-workbook'
import {
  expandToPrimitiveOps,
  isLayoutOp,
  isStructuralOp,
  workbookCommandBatchSchema,
  type WorkbookOperation,
} from '../domain/workbook-dsl'
import type { ApplyOutcome, ChangePlan } from '../domain/workbook.types'
import { isSheetRemoved } from './edit-journal'
import { t } from './i18n/locale'
import { buildLazyChangePlan } from './lazy-plan'
import {
  lazyCellEditable,
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
      // and would silently miss matches — fail early instead.
      for (const operation of batch.operations) {
        if (operation.op !== 'find_replace') continue
        const bounds = parseRange(operation.range)
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
          operation.op === 'rename_sheet' ||
          operation.op === 'format_range' ||
          isLayoutOp(operation) ||
          isStructuralOp(operation)
        ) {
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
    const plan = ctx.adapterRef.current.plan({
      dslVersion: 1,
      transactionId: `agent-${crypto.randomUUID()}`,
      baseRevision: snapshot.revision,
      summary,
      operations,
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
        if (
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
