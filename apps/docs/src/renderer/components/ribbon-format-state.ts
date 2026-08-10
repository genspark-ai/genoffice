import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { isInTable, mergeCells, selectedRect, splitCell } from '@tiptap/pm/tables'
import type { DocDefaults, StyleInfo } from '@genoffice/docx-engine'
import { getActiveSubEditor } from '../editor/active-editor'
import { effectiveSizeHalfPoints } from '../editor/text-style-resolve'
import { textHasCjk } from '../line-metrics'
import { cachedByDoc } from '../doc-cache'

/**
 * Snapshot of every editor-state read the ribbon (and its tabs) shows at render
 * time. App computes it each render and keeps the reference stable via
 * useShallowStable, so React.memo skips the ribbon whenever none of these
 * change (e.g. caret moves inside uniformly formatted text).
 */
export interface RibbonFormatState {
  /** current PM doc: any content change invalidates the ribbon (doc-derived tab displays stay fresh) */
  doc: PmNode | null
  docEmpty: boolean
  /** focused textbox sub-editor receiving ribbon commands (active-editor routing) */
  sub: Editor | null
  editable: boolean
  inTable: boolean
  canMergeCells: boolean
  canSplitCell: boolean
  imageSelected: boolean
  imageDataUrl: string | null
  imageWrap: string | null
  imageAlign: string | null
  imageWidthPx: number | null
  imageHeightPx: number | null
  imageHasDocxIndex: boolean
  textboxSelected: boolean
  shapeFill: string | null
  shapeBorderColor: string | null
  shapePrst: string | null
  cellKey: number | null
  cellHeightCm: number | null
  cellWidthCm: number | null
  cellVAlign: string | null
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  vertAlign: string | null
  highlight: string | null
  textColor: string | null
  charStyleId: string | null
  fontSizePt: number
  fontFamily: string
  headingLevel: number | null
  listBullet: boolean
  listOrdered: boolean
  align: string | null
  /** RTL paragraph (w:bidi) at the cursor */
  bidi: boolean
  lineSpacing: number | null
  shadingFill: string | null
  paraBorders: unknown
  indentLeft: number
  indentRight: number
  spaceBefore: number
  spaceAfter: number
}

export const EMPTY_FORMAT_STATE: RibbonFormatState = {
  doc: null,
  docEmpty: true,
  sub: null,
  editable: false,
  inTable: false,
  canMergeCells: false,
  canSplitCell: false,
  imageSelected: false,
  imageDataUrl: null,
  imageWrap: null,
  imageAlign: null,
  imageWidthPx: null,
  imageHeightPx: null,
  imageHasDocxIndex: false,
  textboxSelected: false,
  shapeFill: null,
  shapeBorderColor: null,
  shapePrst: null,
  cellKey: null,
  cellHeightCm: null,
  cellWidthCm: null,
  cellVAlign: null,
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  vertAlign: null,
  highlight: null,
  textColor: null,
  charStyleId: null,
  fontSizePt: 11,
  fontFamily: '',
  headingLevel: null,
  listBullet: false,
  listOrdered: false,
  align: null,
  bidi: false,
  lineSpacing: null,
  shadingFill: null,
  paraBorders: null,
  indentLeft: 0,
  indentRight: 0,
  spaceBefore: 0,
  spaceAfter: 0,
}

const docEmptyOf = cachedByDoc((doc: PmNode) => doc.textContent.trim() === '')

const str = (v: unknown): string | null => (typeof v === 'string' && v !== '' ? v : null)
const num = (v: unknown): number | null => {
  const n = Number(v)
  return Number.isFinite(n) && n !== 0 ? n : null
}

/** attrs of the paragraph-like node at the cursor (mirrors activeParaAttrs in ribbon-tabs) */
function paraAttrsOf(editor: Editor): Record<string, unknown> {
  if (editor.isActive('docHeading')) return editor.getAttributes('docHeading')
  if (editor.isActive('docListItem')) return editor.getAttributes('docListItem')
  return editor.getAttributes('docParagraph')
}

export function computeFormatState(
  editor: Editor | null,
  styles?: Map<string, StyleInfo>,
  docDefaults?: DocDefaults,
): RibbonFormatState {
  if (!editor || editor.isDestroyed) return EMPTY_FORMAT_STATE
  const sub = getActiveSubEditor()
  const ed = sub ?? editor

  const inTable = !sub && isInTable(editor.state)
  let cellKey: number | null = null
  let cellHeightCm: number | null = null
  let cellWidthCm: number | null = null
  let cellVAlign: string | null = null
  if (inTable) {
    try {
      const rect = selectedRect(editor.state)
      const rowNode = rect.table.maybeChild(rect.top)
      const cellPos = rect.map.map[rect.top * rect.map.width + rect.left]
      const cellNode = editor.state.doc.nodeAt(rect.tableStart + cellPos)
      const colwidth = (cellNode?.attrs.colwidth as number[] | null) ?? null
      cellKey = rect.tableStart * 100000 + cellPos
      cellHeightCm = rowNode?.attrs.heightTwips
        ? ((rowNode.attrs.heightTwips as number) / 1440) * 2.54
        : null
      cellWidthCm = colwidth?.[0] ? (colwidth[0] / 96) * 2.54 : null
      cellVAlign = (cellNode?.attrs.vAlign as string | null) ?? null
    } catch {
      /* degenerate selection: leave cell fields empty like the old activeCellInfo */
    }
  }

  const protAttrs = editor.getAttributes('docProtected')
  const imageSelected = protAttrs.blockType === 'image' && !!protAttrs.imageDataUrl

  const textAttrs = ed.getAttributes('docTextStyle')
  // Dual-slot runs: like Word's font box, show the slot matching the script at the caret
  const displayFont = (): string => {
    const font = str(textAttrs.font)
    const fontAscii = str(textAttrs.fontAscii)
    if (!font || !fontAscii || font === fontAscii) return font ?? fontAscii ?? ''
    const { from, to } = ed.state.selection
    const sample =
      from === to
        ? ed.state.doc.textBetween(
            Math.max(0, from - 1),
            Math.min(ed.state.doc.content.size, from + 1),
          )
        : ed.state.doc.textBetween(from, Math.min(to, from + 32), ' ')
    return textHasCjk(sample) ? font : fontAscii
  }
  const paraAttrs = sub ? ed.getAttributes('docParagraph') : paraAttrsOf(editor)
  const mainPara = paraAttrsOf(editor)

  return {
    doc: editor.state.doc,
    docEmpty: docEmptyOf(editor.state.doc),
    sub,
    editable: editor.isEditable,
    inTable,
    canMergeCells: inTable && !!mergeCells(editor.state),
    canSplitCell: inTable && !!splitCell(editor.state),
    imageSelected,
    imageDataUrl: imageSelected ? str(protAttrs.imageDataUrl) : null,
    imageWrap: str(protAttrs.imageWrap),
    imageAlign: str(protAttrs.imageAlign),
    imageWidthPx: num(protAttrs.imageWidthPx),
    imageHeightPx: num(protAttrs.imageHeightPx),
    imageHasDocxIndex: protAttrs.docxIndex != null,
    textboxSelected: Array.isArray(protAttrs.textboxes) && protAttrs.textboxes.length > 0,
    shapeFill: Array.isArray(protAttrs.textboxes)
      ? str((protAttrs.textboxes[0] as { fill?: string } | undefined)?.fill)
      : null,
    shapeBorderColor: Array.isArray(protAttrs.textboxes)
      ? str((protAttrs.textboxes[0] as { borderColor?: string } | undefined)?.borderColor)
      : null,
    shapePrst: Array.isArray(protAttrs.textboxes)
      ? str((protAttrs.textboxes[0] as { prst?: string } | undefined)?.prst)
      : null,
    cellKey,
    cellHeightCm,
    cellWidthCm,
    cellVAlign,
    bold: ed.isActive('bold'),
    italic: ed.isActive('italic'),
    underline: ed.isActive('underline'),
    strike: ed.isActive('strike'),
    vertAlign: str(textAttrs.vertAlign),
    highlight: str(textAttrs.highlight),
    textColor: str(textAttrs.color),
    charStyleId: str(textAttrs.styleId),
    fontSizePt: (effectiveSizeHalfPoints(ed, styles, docDefaults) ?? 22) / 2,
    fontFamily: displayFont(),
    headingLevel: editor.isActive('docHeading')
      ? Number(editor.getAttributes('docHeading').level ?? 1)
      : null,
    listBullet: editor.isActive('docListItem', { kind: 'bullet' }),
    listOrdered: editor.isActive('docListItem', { kind: 'ordered' }),
    align: str(paraAttrs.align),
    bidi: paraAttrs.bidi === true,
    lineSpacing: typeof paraAttrs.lineSpacing === 'number' ? paraAttrs.lineSpacing : null,
    shadingFill: str(paraAttrs.shadingFill),
    paraBorders: paraAttrs.borders ?? null,
    indentLeft: Number(mainPara.indentLeft) || 0,
    indentRight: Number(mainPara.indentRight) || 0,
    spaceBefore: Number(mainPara.spaceBefore) || 0,
    spaceAfter: Number(mainPara.spaceAfter) || 0,
  }
}
