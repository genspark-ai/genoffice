/**
 * Table and chart edit ops. Structure changes (merge/insert/delete) and style
 * edits reparse the slide, which regenerates element ids — those ops report
 * the surviving element's new id in the record's `after` so callers can keep
 * the selection.
 */
import {
  applyParagraphFormat,
  editChartElement,
  editTableCellText,
  editTableStructure,
  ensureTableStylePart,
  markChartEditable,
  mergeTableCells,
  setTableCellAnchor,
  setTableColWidth,
  setTableRowHeight,
  TABLE_STYLE_PRESETS,
  EMU_PER_PT,
  type TableElement,
  type TableMergeOp,
  type TableStructureOp,
  type TableStyleEdit,
} from '@genoffice/pptx-engine'
import { editTableStyle } from '@genoffice/pptx-engine'
import type { EditParagraph } from '../../shared/ipc'
import { applyEditParagraphs, collectParagraphFormatPatches } from '../edit-text'
import { GuidedError, register, resolveElement, type Op, type OpRecord } from './registry'

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
    const row = op.row as number
    const col = op.col as number
    // Callers (the cell editor, the AI tools, scripts) send EditParagraphs:
    // rebuild them onto the cell's current paragraphs exactly like setText does
    // for shapes, so the run/paragraph properties the caller cannot express
    // (size, color, font, bold, bullets, theme links) stay with the cell instead
    // of falling back to the table-style defaults.
    const cell = (el as TableElement).rows[row]?.[col]
    const edited = op.paragraphs as EditParagraph[]
    const paragraphs = applyEditParagraphs(cell?.text?.paragraphs ?? [], edited)
    // Bullets, spacing and direction toggled during the edit session ride along
    // as per-paragraph patches (bullet/lineSpacingPct/spaceBeforePt/spaceAfterPt/
    // rtl); setText applies them through setElementParagraphFormat, which is
    // table-wide for a table, so patch the cell's rebuilt paragraphs directly.
    for (const { index: pi, patch } of collectParagraphFormatPatches(edited)) {
      const target = paragraphs[pi]
      if (!target) continue
      // the rebuild spread shares pPrExplicit with the cell's current paragraph
      if (target.pPrExplicit) target.pPrExplicit = { ...target.pPrExplicit }
      applyParagraphFormat([target], patch)
    }
    if (!editTableCellText(slide, el.id, row, col, paragraphs)) {
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
// Model-facing fields resolve here so the ribbon and the AI send the same op:
// a preset name pins its fixed-color style definition into tableStyles.xml
// (built-in GUIDs track theme colors, so colors would drift), and applying a
// preset clears direct cell formatting like PowerPoint's style gallery does.
const HEX_RE = /^#[0-9A-Fa-f]{6}$/
const STYLE_FIELDS = [
  'firstRow',
  'bandRow',
  'rtl',
  'shadingColor',
  'borderColor',
  'borderWidthPt',
  'borderPreset',
  'cells',
] as const

interface ResolvedTableStyle {
  edit: TableStyleEdit
  stylePart?: { styleId: string; styleDefXml: string }
}

function resolveTableStyle(op: Op): ResolvedTableStyle {
  if (op.styleName != null) {
    const preset = TABLE_STYLE_PRESETS[String(op.styleName)]
    if (!preset) {
      throw new GuidedError(
        `op "setTableStyle": unknown styleName "${String(op.styleName)}". Presets: ${Object.keys(TABLE_STYLE_PRESETS).join(', ')}.`,
      )
    }
    return {
      edit: {
        tblPrXml: preset.tblPrXml,
        clearDirectFormatting: true,
        // Grid presets draw the outer frame with direct borders (styles only have inner lines)
        ...(preset.border
          ? {
              borderPreset: 'all' as const,
              borderColor: preset.border.color,
              borderWidthEmu: preset.border.widthEmu,
            }
          : {}),
      },
      ...(preset.styleId
        ? { stylePart: { styleId: preset.styleId, styleDefXml: preset.styleDefXml! } }
        : {}),
    }
  }
  if (!STYLE_FIELDS.some((f) => op[f] != null)) {
    throw new GuidedError(
      'op "setTableStyle" needs "styleName" or at least one of firstRow, bandRow, shadingColor, borderColor, borderWidthPt, borderPreset.',
    )
  }
  if (
    op.shadingColor != null &&
    op.shadingColor !== 'none' &&
    !HEX_RE.test(String(op.shadingColor))
  ) {
    throw new GuidedError('op "setTableStyle": shadingColor must be #RRGGBB or "none".')
  }
  if (op.borderColor != null && !HEX_RE.test(String(op.borderColor))) {
    throw new GuidedError('op "setTableStyle": borderColor must be #RRGGBB.')
  }
  if (op.borderWidthPt != null && !(Number(op.borderWidthPt) > 0)) {
    throw new GuidedError('op "setTableStyle": borderWidthPt must be a positive number.')
  }
  if (op.borderPreset != null && op.borderPreset !== 'all' && op.borderPreset !== 'none') {
    throw new GuidedError('op "setTableStyle": borderPreset must be "all" or "none".')
  }
  return {
    edit: {
      ...(op.firstRow != null ? { firstRow: Boolean(op.firstRow) } : {}),
      ...(op.bandRow != null ? { bandRow: Boolean(op.bandRow) } : {}),
      ...(op.rtl != null ? { rtl: Boolean(op.rtl) } : {}),
      ...(op.shadingColor != null ? { shadingColor: String(op.shadingColor) } : {}),
      ...(op.borderColor != null ? { borderColor: String(op.borderColor) } : {}),
      ...(op.borderWidthPt != null
        ? { borderWidthEmu: Math.round(Number(op.borderWidthPt) * EMU_PER_PT) }
        : {}),
      ...(op.borderPreset != null ? { borderPreset: op.borderPreset as 'all' | 'none' } : {}),
      ...(op.cells != null ? { cells: op.cells as TableStyleEdit['cells'] } : {}),
    },
  }
}

register({
  name: 'setTableStyle',
  validate(op, ctx) {
    resolveElement(ctx, op, { types: ['table'] })
    resolveTableStyle(op)
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op, { types: ['table'] })
    const { edit, stylePart } = resolveTableStyle(op)
    if (stylePart) ensureTableStylePart(ctx.opened, stylePart.styleId, stylePart.styleDefXml)
    if (!editTableStyle(slide, el.id, edit)) {
      throw new GuidedError(`op "setTableStyle": table "${el.id}" rejected the style edit.`)
    }
    return { op, after: edit }
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
