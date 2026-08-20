/**
 * Table and chart edit ops. Structure changes (merge/insert/delete) and style
 * edits reparse the slide, which regenerates element ids — those ops report
 * the surviving element's new id in the record's `after` so callers can keep
 * the selection.
 */
import {
  editChartElement,
  editTableCellText,
  editTableStructure,
  ensureTableStylePart,
  markChartEditable,
  mergeTableCells,
  setTableCellAnchor,
  setTableColWidth,
  setTableRowHeight,
  type Paragraph,
  type TableMergeOp,
  type TableStructureOp,
  type TableStyleEdit,
} from '@genoffice/pptx-engine'
import { editTableStyle } from '@genoffice/pptx-engine'
import { GuidedError, register, resolveElement, type OpRecord } from './registry'

register({
  name: 'setTableCell',
  validate(op, ctx) {
    resolveElement(ctx, op, { types: ['table'] })
    if (typeof op.row !== 'number' || typeof op.col !== 'number' || !Array.isArray(op.paragraphs)) {
      throw new GuidedError('op "setTableCell" needs "row", "col" and "paragraphs".')
    }
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op, { types: ['table'] })
    if (
      !editTableCellText(
        slide,
        el.id,
        op.row as number,
        op.col as number,
        op.paragraphs as Paragraph[],
      )
    ) {
      throw new GuidedError(
        `op "setTableCell": cell (${op.row}, ${op.col}) does not exist on table "${el.id}".`,
      )
    }
    return { op }
  },
})

register({
  name: 'tableMerge',
  validate(op, ctx) {
    resolveElement(ctx, op, { types: ['table'] })
    if (typeof op.kind !== 'string' || typeof op.row !== 'number' || typeof op.col !== 'number') {
      throw new GuidedError('op "tableMerge" needs "kind", "row" and "col".')
    }
  },
  apply(op, ctx): OpRecord {
    const { index, el } = resolveElement(ctx, op, { types: ['table'] })
    const r = mergeTableCells(ctx.opened, index, el.id, {
      kind: op.kind,
      row: op.row,
      col: op.col,
    } as TableMergeOp)
    if (!r) {
      throw new GuidedError(
        `op "tableMerge": ${op.kind} is not possible at (${op.row}, ${op.col}) — check merge boundaries.`,
      )
    }
    return { op, after: { elementId: r.elementId } }
  },
})

register({
  name: 'tableStructure',
  validate(op, ctx) {
    resolveElement(ctx, op, { types: ['table'] })
    if (typeof op.kind !== 'string' || typeof op.index !== 'number') {
      throw new GuidedError('op "tableStructure" needs "kind" and "index".')
    }
  },
  apply(op, ctx): OpRecord {
    const { index, el } = resolveElement(ctx, op, { types: ['table'] })
    const r = editTableStructure(ctx.opened, index, el.id, {
      kind: op.kind,
      index: op.index,
      ...(op.before ? { before: true } : {}),
    } as TableStructureOp)
    if (!r) {
      throw new GuidedError(
        `op "tableStructure": ${op.kind} at index ${op.index} failed (merges crossing the boundary?).`,
      )
    }
    return { op, after: { elementId: r.elementId } }
  },
})

register({
  name: 'setTableRowHeight',
  validate(op, ctx) {
    resolveElement(ctx, op, { types: ['table'] })
    if (typeof op.row !== 'number' || typeof op.hEmu !== 'number') {
      throw new GuidedError('op "setTableRowHeight" needs "row" and "hEmu".')
    }
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op, { types: ['table'] })
    if (!setTableRowHeight(slide, el.id, op.row as number, op.hEmu as number)) {
      throw new GuidedError(`op "setTableRowHeight": row ${op.row} does not exist on "${el.id}".`)
    }
    return { op }
  },
})

register({
  name: 'setTableCellAnchor',
  validate(op, ctx) {
    resolveElement(ctx, op, { types: ['table'] })
    if (typeof op.row !== 'number' || typeof op.col !== 'number') {
      throw new GuidedError('op "setTableCellAnchor" needs "row" and "col".')
    }
    if (!['top', 'middle', 'bottom'].includes(String(op.anchor))) {
      throw new GuidedError('op "setTableCellAnchor" needs "anchor": top/middle/bottom.')
    }
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op, { types: ['table'] })
    if (
      !setTableCellAnchor(
        slide,
        el.id,
        op.row as number,
        op.col as number,
        op.anchor as 'top' | 'middle' | 'bottom',
      )
    ) {
      throw new GuidedError(
        `op "setTableCellAnchor": cell (${op.row}, ${op.col}) does not exist on "${el.id}".`,
      )
    }
    return { op }
  },
})

register({
  name: 'setTableColWidth',
  validate(op, ctx) {
    resolveElement(ctx, op, { types: ['table'] })
    if (typeof op.col !== 'number' || typeof op.wEmu !== 'number') {
      throw new GuidedError('op "setTableColWidth" needs "col" and "wEmu".')
    }
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op, { types: ['table'] })
    if (!setTableColWidth(slide, el.id, op.col as number, op.wEmu as number)) {
      throw new GuidedError(`op "setTableColWidth": column ${op.col} does not exist on "${el.id}".`)
    }
    return { op }
  },
})

// ── setTableStyle ───────────────────────────────────────────────────────
// Preset-name resolution stays in the shim (the preset table is app data);
// the op takes the resolved TableStyleEdit plus an optional style part to
// inject (fixed-color presets pin their definition into tableStyles.xml).
register({
  name: 'setTableStyle',
  validate(op, ctx) {
    resolveElement(ctx, op, { types: ['table'] })
    if (typeof op.edit !== 'object' || op.edit === null) {
      throw new GuidedError('op "setTableStyle" needs "edit": a TableStyleEdit object.')
    }
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op, { types: ['table'] })
    const part = op.stylePart as { styleId: string; styleDefXml: string } | undefined
    if (part) ensureTableStylePart(ctx.opened, part.styleId, part.styleDefXml)
    if (!editTableStyle(slide, el.id, op.edit as TableStyleEdit)) {
      throw new GuidedError(`op "setTableStyle": table "${el.id}" rejected the style edit.`)
    }
    return { op, after: op.edit }
  },
})

// ── setChart ────────────────────────────────────────────────────────────
// The chart part XML (and embedded workbook) is rewritten wholesale; the
// import-confirmation dialog is UI and stays in the shim.
register({
  name: 'setChart',
  validate(op, ctx) {
    resolveElement(ctx, op, { types: ['chart'] })
    if (typeof op.patch !== 'object' || op.patch === null) {
      throw new GuidedError('op "setChart" needs "patch": a chart edit object.')
    }
  },
  apply(op, ctx): OpRecord {
    const { index, slide, el } = resolveElement(ctx, op, { types: ['chart'] })
    // Mark aislides-chart on first edit (the conversion itself is lossless; no re-prompt after one confirmation)
    markChartEditable(slide, el.id)
    if (
      !editChartElement(
        ctx.opened,
        index,
        el.id,
        op.patch as Parameters<typeof editChartElement>[3],
      )
    ) {
      throw new GuidedError(`op "setChart": chart "${el.id}" rejected the edit.`)
    }
    return { op, after: op.patch }
  },
})
