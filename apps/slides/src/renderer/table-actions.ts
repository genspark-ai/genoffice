/**
 * Table editing actions extracted from App.tsx: row/column
 * structure, merge/split, resize, cell text commit and Tab navigation.
 * Functions read the latest App state through ActionCtx.
 */
import type { EditParagraph } from '../shared/ipc'
import type { ActionCtx } from './action-context'
import { FIT_WIDTH } from './app-constants'
import { t } from './i18n/locale'

export async function tableStructureOp(
  ctx: ActionCtx,
  sourceId: string,
  kind: 'insert-row' | 'delete-row' | 'insert-col' | 'delete-col',
  index: number,
  before?: boolean,
): Promise<void> {
  const r = await window.slidesApi.tableStructure({
    slideIndex: ctx.current,
    sourceId,
    kind,
    index,
    ...(before ? { before: true } : {}),
  })
  if (r) {
    ctx.applySlide(ctx.current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    ctx.setEditingCell(null)
  } else {
    ctx.setStatus(
      kind.startsWith('delete')
        ? t('appStatusTableDeleteLimit')
        : t('appStatusTableMergedUnsupported'),
    )
  }
}

export async function tableMergeOp(
  ctx: ActionCtx,
  sourceId: string,
  kind: 'merge-right' | 'merge-down' | 'split',
  row: number,
  col: number,
): Promise<void> {
  const r = await window.slidesApi.tableMerge({
    slideIndex: ctx.current,
    sourceId,
    kind,
    row,
    col,
  })
  if (r) {
    ctx.applySlide(ctx.current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    ctx.setEditingCell(null)
  } else {
    ctx.setStatus(t('appStatusMergeFailed'))
  }
}

export async function onTableColResize(
  ctx: ActionCtx,
  sourceId: string,
  col: number,
  wPx: number,
): Promise<void> {
  const updated = await window.slidesApi.setTableColWidth({
    slideIndex: ctx.current,
    sourceId,
    col,
    wPx,
    fitWidthPx: FIT_WIDTH,
  })
  if (updated) ctx.applySlide(ctx.current, updated)
}

export async function onTableRowResize(
  ctx: ActionCtx,
  sourceId: string,
  row: number,
  hPx: number,
): Promise<void> {
  const updated = await window.slidesApi.setTableRowHeight({
    slideIndex: ctx.current,
    sourceId,
    row,
    hPx,
    fitWidthPx: FIT_WIDTH,
  })
  if (updated) ctx.applySlide(ctx.current, updated)
}

export async function commitCellEdit(ctx: ActionCtx, paragraphs: EditParagraph[]): Promise<void> {
  if (!ctx.editingCell) return
  const updated = await window.slidesApi.editTableCell({
    slideIndex: ctx.current,
    sourceId: ctx.editingCell.sourceId,
    row: ctx.editingCell.row,
    col: ctx.editingCell.col,
    paragraphs,
  })
  if (updated) ctx.applySlide(ctx.current, updated)
  ctx.setEditingCell(null)
}

/** Cell Tab/Shift+Tab: commit the current cell (if content changed) and jump to the next/previous cell;
 * Tab on the last cell auto-adds a row and enters its first cell. */
export async function navigateCell(
  ctx: ActionCtx,
  paragraphs: EditParagraph[] | null,
  dir: 1 | -1,
): Promise<void> {
  const { editingCell, slide, current } = ctx
  if (!editingCell || !slide) return
  const tbl = slide.nodes.find((n) => n.sourceId === editingCell.sourceId)
  if (!tbl || tbl.type !== 'table') return
  if (paragraphs) {
    const updated = await window.slidesApi.editTableCell({
      slideIndex: current,
      sourceId: editingCell.sourceId,
      row: editingCell.row,
      col: editingCell.col,
      paragraphs,
    })
    if (updated) ctx.applySlide(current, updated)
  }
  const cells = [...tbl.cells].sort((a, b) => a.row - b.row || a.col - b.col)
  const idx = cells.findIndex((c) => c.row === editingCell.row && c.col === editingCell.col)
  const next = cells[idx + dir]
  if (next) {
    ctx.setEditingCell({ sourceId: editingCell.sourceId, row: next.row, col: next.col })
    return
  }
  if (dir > 0) {
    // Tab on the last cell: add a row and enter its first cell
    const lastRow = Math.max(...cells.map((c) => c.row))
    const r = await window.slidesApi.tableStructure({
      slideIndex: current,
      sourceId: editingCell.sourceId,
      kind: 'insert-row',
      index: lastRow,
    })
    if (r) {
      ctx.applySlide(current, r.slide)
      ctx.setSelectedIds([r.sourceId])
      const firstCol = Math.min(...cells.filter((c) => c.row === lastRow).map((c) => c.col))
      ctx.setEditingCell({ sourceId: r.sourceId, row: lastRow + 1, col: firstCol })
    }
  }
  // Shift+Tab on the first cell: stay in place
}
