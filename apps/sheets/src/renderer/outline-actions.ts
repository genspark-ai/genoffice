/**
 * Outline (Group) actions behind the gutter buttons and the Data tab:
 * fold/unfold rides Univer's hide/show commands (journaled, undoable); the
 * summary-line collapsed flags and sheetPr/outlinePr go straight to the
 * journal because Univer has no outline model.
 */
import type { IRange } from '@univerjs/core'

import { isSheetRemoved, journalSize, recordStructuralOp } from './edit-journal'
import { t } from './i18n/locale'
import {
  groupTogglePlan,
  levelButtonPlan,
  outlineGroups,
  type IndexSpan,
  type OutlineEntry,
  type OutlineGroup,
  type SheetOutlineState,
} from './outline-model'
import { sheetOutline } from './univer-sync'
import type { LazyWorkbookState, UniverRuntime, UniverWorksheet } from './univer-state'

export type OutlineAxis = 'rows' | 'cols'

export interface OutlineActionContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  setMessage: (message: string) => void
  setPendingEdits: (count: number) => void
}

export function outlineEntries(
  outline: SheetOutlineState,
  axis: OutlineAxis,
): Map<number, OutlineEntry> {
  return axis === 'rows' ? outline.rows : outline.cols
}

/// Hidden state per line: Univer knows every loaded row; past the streamed
/// coverage the file's own hidden flag (from the sidecar) is the truth.
export function outlineHiddenProbe(
  state: LazyWorkbookState,
  worksheet: UniverWorksheet,
  sheetId: string,
  axis: OutlineAxis,
): (index: number) => boolean {
  const sheet = worksheet.getSheet()
  const entries = outlineEntries(sheetOutline(state, sheetId), axis)
  if (axis === 'cols') return (index) => !sheet.getColVisible(index)
  const covered = state.hiddenRowsCoveredThrough.get(sheetId) ?? -1
  return (index) =>
    !sheet.getRowVisible(index) || (index > covered && entries.get(index)?.fileHidden === true)
}

export function outlineGroupsFor(
  state: LazyWorkbookState,
  worksheet: UniverWorksheet,
  sheetId: string,
  axis: OutlineAxis,
): OutlineGroup[] {
  const outline = sheetOutline(state, sheetId)
  return outlineGroups(
    outlineEntries(outline, axis),
    axis === 'rows' ? outline.summaryBelow : outline.summaryRight,
    outlineHiddenProbe(state, worksheet, sheetId, axis),
  )
}

function applySpans(
  worksheet: UniverWorksheet,
  entries: Map<number, OutlineEntry>,
  axis: OutlineAxis,
  spans: readonly IndexSpan[],
  hidden: boolean,
): void {
  const limit = axis === 'rows' ? worksheet.getMaxRows() : worksheet.getMaxColumns()
  for (const span of spans) {
    const end = Math.min(span.end, limit - 1)
    if (end < span.start) continue
    const count = end - span.start + 1
    if (axis === 'rows') {
      if (hidden) worksheet.hideRows(span.start, count)
      else worksheet.showRows(span.start, count)
    } else if (hidden) {
      worksheet.hideColumns(span.start, count)
    } else {
      worksheet.showColumns(span.start, count)
    }
    for (let index = span.start; index <= end; index += 1) {
      const entry = entries.get(index)
      if (entry) entry.fileHidden = hidden
    }
  }
}

function recordSummaries(
  state: LazyWorkbookState,
  sheetId: string,
  axis: OutlineAxis,
  summaries: readonly { index: number; collapsed: boolean }[],
): void {
  const outline = sheetOutline(state, sheetId)
  const entries = outlineEntries(outline, axis)
  for (const summary of summaries) {
    const level = entries.get(summary.index)?.level ?? 0
    const entry = entries.get(summary.index)
    if (entry) entry.collapsed = summary.collapsed
    else entries.set(summary.index, { level, collapsed: summary.collapsed })
    recordStructuralOp(state.editJournal, sheetId, {
      kind: axis === 'rows' ? 'set-rows-outline' : 'set-cols-outline',
      start: summary.index,
      end: summary.index,
      level,
      collapsed: summary.collapsed,
    })
  }
  outline.version += 1
}

function activeTarget(
  ctx: OutlineActionContext,
): { state: LazyWorkbookState; worksheet: UniverWorksheet; sheetId: string } | null {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  const worksheet = runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!runtime || !state || !worksheet) return null
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return null
  return { state, worksheet, sheetId }
}

/// Gutter [-] / [+] on one group.
export function toggleOutlineGroup(
  ctx: OutlineActionContext,
  axis: OutlineAxis,
  group: OutlineGroup,
  collapsed: boolean,
): void {
  const target = activeTarget(ctx)
  if (!target) return
  const { state, worksheet, sheetId } = target
  const groups = outlineGroupsFor(state, worksheet, sheetId, axis)
  const plan = groupTogglePlan(groups, group, collapsed)
  const entries = outlineEntries(sheetOutline(state, sheetId), axis)
  applySpans(worksheet, entries, axis, plan.show, false)
  applySpans(worksheet, entries, axis, plan.hide, true)
  recordSummaries(state, sheetId, axis, plan.summaries)
  ctx.setPendingEdits(journalSize(state.editJournal))
  ctx.setMessage(collapsed ? t('appDetailHidden') : t('appDetailShown'))
}

/// Corner level button N: show detail through level N-1.
export function showOutlineLevel(
  ctx: OutlineActionContext,
  axis: OutlineAxis,
  level: number,
): void {
  const target = activeTarget(ctx)
  if (!target) return
  const { state, worksheet, sheetId } = target
  const entries = outlineEntries(sheetOutline(state, sheetId), axis)
  const groups = outlineGroupsFor(state, worksheet, sheetId, axis)
  const plan = levelButtonPlan(entries, groups, level)
  applySpans(worksheet, entries, axis, plan.show, false)
  applySpans(worksheet, entries, axis, plan.hide, true)
  recordSummaries(state, sheetId, axis, plan.summaries)
  ctx.setPendingEdits(journalSize(state.editJournal))
  ctx.setMessage(t('appOutlineLevelShown', { level }))
}

/// Data → Ungroup → Clear Outline: a single selected cell clears the whole
/// sheet, otherwise the rows and columns the selection spans.
export function clearOutline(ctx: OutlineActionContext, selection: IRange | null): void {
  const target = activeTarget(ctx)
  if (!target) return
  const { state, worksheet, sheetId } = target
  const outline = sheetOutline(state, sheetId)
  const whole =
    !selection ||
    (selection.startRow === selection.endRow && selection.startColumn === selection.endColumn)
  let cleared = 0
  for (const axis of ['rows', 'cols'] as const) {
    const entries = outlineEntries(outline, axis)
    const groups = outlineGroupsFor(state, worksheet, sheetId, axis)
    const [from, to] = whole
      ? [0, Number.MAX_SAFE_INTEGER]
      : axis === 'rows'
        ? [selection.startRow, selection.endRow]
        : [selection.startColumn, selection.endColumn]
    const spans: IndexSpan[] = []
    for (const index of [...entries.keys()].sort((left, right) => left - right)) {
      if (index < from || index > to) continue
      const last = spans[spans.length - 1]
      if (last && last.end === index - 1)
        spans[spans.length - 1] = { start: last.start, end: index }
      else spans.push({ start: index, end: index })
    }
    if (spans.length === 0) continue
    // Lines losing their outline would otherwise stay hidden with no
    // [+] left to bring them back; unfold them first.
    applySpans(
      worksheet,
      entries,
      axis,
      spans.filter((span) =>
        groups.some(
          (group) => group.collapsed && group.start <= span.end && group.end >= span.start,
        ),
      ),
      false,
    )
    for (const span of spans) {
      recordStructuralOp(state.editJournal, sheetId, {
        kind: axis === 'rows' ? 'set-rows-outline' : 'set-cols-outline',
        start: span.start,
        end: span.end,
        level: 0,
        collapsed: false,
      })
      for (let index = span.start; index <= span.end; index += 1) entries.delete(index)
      cleared += 1
    }
  }
  if (cleared === 0) {
    ctx.setMessage(t('appNothingToUngroup'))
    return
  }
  outline.version += 1
  ctx.setPendingEdits(journalSize(state.editJournal))
  ctx.setMessage(t('appOutlineCleared'))
}

/// Outline settings: summary rows below detail / summary columns right of detail.
export function setOutlineSummary(
  ctx: OutlineActionContext,
  patch: { below?: boolean; right?: boolean },
): void {
  const target = activeTarget(ctx)
  if (!target) return
  const { state, sheetId } = target
  const outline = sheetOutline(state, sheetId)
  if (patch.below !== undefined) outline.summaryBelow = patch.below
  if (patch.right !== undefined) outline.summaryRight = patch.right
  outline.version += 1
  // One outlinePr op per sheet: the latest placement replaces earlier ones.
  const ops = state.editJournal.structuralOps.get(sheetId)
  if (ops) {
    const kept = ops.filter((op) => op.kind !== 'set-outline-pr')
    if (kept.length === 0) state.editJournal.structuralOps.delete(sheetId)
    else state.editJournal.structuralOps.set(sheetId, kept)
  }
  recordStructuralOp(state.editJournal, sheetId, {
    kind: 'set-outline-pr',
    summaryBelow: outline.summaryBelow,
    summaryRight: outline.summaryRight,
  })
  ctx.setPendingEdits(journalSize(state.editJournal))
}

export function outlineSummarySettings(
  state: LazyWorkbookState | null,
  sheetId: string | undefined,
): { below: boolean; right: boolean } {
  if (!state || !sheetId) return { below: true, right: true }
  const outline = sheetOutline(state, sheetId)
  return { below: outline.summaryBelow, right: outline.summaryRight }
}

/// Row outline levels for the whole sheet come from the sidecar index, not
/// from the rows that happen to have streamed in; polled until indexing ends.
export async function loadRowOutlines(
  state: LazyWorkbookState,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
): Promise<void> {
  const sheets = state.file.sheets.filter((sheet) => (sheet.outlineLevelRow ?? 0) > 0)
  if (sheets.length === 0) return
  const deadline = Date.now() + 180_000
  for (const sheet of sheets) {
    for (;;) {
      if (lazyWorkbookRef.current !== state) return
      let result
      try {
        result = await window.desktopApi.readWorkbookRowOutline({
          sessionId: state.file.sessionId,
          sheetId: sheet.id,
        })
      } catch {
        return
      }
      const outline = sheetOutline(state, sheet.id)
      // Once rows shifted this session, file coordinates no longer map.
      if (state.editJournal.structuralOps.get(sheet.id)?.some((op) => 'index' in op)) break
      let changed = false
      for (const row of result.rows) {
        // Each file row seeds once; session edits own the entry afterwards.
        if (outline.seededRows.has(row.row)) continue
        outline.seededRows.add(row.row)
        if (outline.rows.has(row.row)) continue
        outline.rows.set(row.row, {
          level: row.level,
          collapsed: row.collapsed ?? false,
          fileHidden: row.hidden ?? false,
        })
        changed = true
      }
      if (changed) outline.version += 1
      if (result.indexingComplete || Date.now() > deadline) break
      await new Promise((resolve) => setTimeout(resolve, 500))
    }
  }
}
