import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { isInTable, mergeCells, selectedRect, splitCell } from '@tiptap/pm/tables'
import type { DocDefaults, Run, StyleInfo, TextboxDisplay } from '@genoffice/docx-engine'
import { getActiveSubEditor } from '../editor/active-editor'
import { tableCellsSelection } from '../editor/table-ops'
import { effectiveBidi, selectionHasBidi } from '../editor/direction'
import {
  effectiveSizeHalfPoints,
  selectedFonts,
  selectedSizeMixed,
} from '../editor/text-style-resolve'
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
  /** Split Table needs a row above the caret's */
  canSplitTable: boolean
  imageSelected: boolean
  imageDataUrl: string | null
  imageWrap: string | null
  imageAlign: string | null
  imageWidthPx: number | null
  imageHeightPx: number | null
  imageFlipH: boolean
  imageFlipV: boolean
  imageHasDocxIndex: boolean
  /** document position of the selected picture: per-picture UI state keys on it */
  imageKey: number | null
  textboxSelected: boolean
  shapeFill: string | null
  shapeBorderColor: string | null
  shapePrst: string | null
  /**
   * What the selected shape's own text agrees on, for the Shape Format Text
   * group. Selecting a shape as an object leaves no text selection for a mark to
   * hang off, so Word formats every run at once and lights a button only while
   * the shape agrees throughout: one bold word among plain ones leaves Bold off,
   * and pressing it bolds all of them. Null fields mean the runs disagree; a
   * shape with no text to format has shapeHasText false.
   */
  shapeHasText: boolean
  shapeTextBold: boolean
  shapeTextItalic: boolean
  shapeTextUnderline: boolean
  shapeTextColor: string | null
  shapeTextAlign: string | null
  cellKey: number | null
  cellHeightTwips: number | null
  cellWidthPx: number | null
  cellVAlign: string | null
  cellTextDirection: string | null
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  vertAlign: string | null
  highlight: string | null
  textColor: string | null
  charStyleId: string | null
  /** w:pStyle of the paragraph at the caret (null = the default paragraph style) */
  paraStyleId: string | null
  fontSizePt: number
  /** selection spans more than one size: the size box shows blank (Word) */
  fontSizeMixed: boolean
  fontEastAsia: string | null
  fontLatin: string | null
  fontFamily: string
  headingLevel: number | null
  listBullet: boolean
  listOrdered: boolean
  align: string | null
  /** RTL paragraph (w:bidi) at the cursor */
  bidi: boolean
  /** an RTL paragraph is at the cursor or inside the selection */
  selectionBidi: boolean
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
  canSplitTable: false,
  imageSelected: false,
  imageDataUrl: null,
  imageWrap: null,
  imageAlign: null,
  imageWidthPx: null,
  imageHeightPx: null,
  imageFlipH: false,
  imageFlipV: false,
  imageHasDocxIndex: false,
  imageKey: null,
  textboxSelected: false,
  shapeFill: null,
  shapeBorderColor: null,
  shapePrst: null,
  shapeHasText: false,
  shapeTextBold: false,
  shapeTextItalic: false,
  shapeTextUnderline: false,
  shapeTextColor: null,
  shapeTextAlign: null,
  cellKey: null,
  cellHeightTwips: null,
  cellWidthPx: null,
  cellVAlign: null,
  cellTextDirection: null,
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  vertAlign: null,
  highlight: null,
  textColor: null,
  charStyleId: null,
  paraStyleId: null,
  fontSizePt: 11,
  fontSizeMixed: false,
  fontFamily: '',
  fontEastAsia: '',
  fontLatin: '',
  headingLevel: null,
  listBullet: false,
  listOrdered: false,
  align: null,
  bidi: false,
  selectionBidi: false,
  lineSpacing: null,
  shadingFill: null,
  paraBorders: null,
  indentLeft: 0,
  indentRight: 0,
  spaceBefore: 0,
  spaceAfter: 0,
}

// stops at the first character instead of building the whole document's text
const docEmptyOf = cachedByDoc((doc: PmNode) => {
  let empty = true
  doc.descendants((node) => {
    if (!empty) return false
    if (node.isText && node.text?.trim()) empty = false
    else if (node.isLeaf && !node.isText) empty = false
    return empty
  })
  return empty
})

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

/**
 * Fold every run and paragraph of a shape into what they all share. A run with
 * no color of its own draws in the shape style's default (a:fontRef), so that is
 * what the swatch reports — otherwise a new white-on-blue shape would read as
 * having no color at all.
 */
function shapeTextStateOf(
  box: TextboxDisplay | undefined,
): Pick<
  RibbonFormatState,
  | 'shapeHasText'
  | 'shapeTextBold'
  | 'shapeTextItalic'
  | 'shapeTextUnderline'
  | 'shapeTextColor'
  | 'shapeTextAlign'
> {
  const runs = box && !box.readOnly ? box.paras.flatMap((p) => p.runs) : []
  if (!box || runs.length === 0) {
    return {
      shapeHasText: false,
      shapeTextBold: false,
      shapeTextItalic: false,
      shapeTextUnderline: false,
      shapeTextColor: null,
      shapeTextAlign: null,
    }
  }
  const every = (has: (run: Run) => boolean): boolean => runs.every(has)
  // auto (Word's automatic colour) shows as "no colour" in the swatch
  const colorOf = (r: Run): string | null =>
    r.color === 'auto' ? null : (r.color ?? box.textColor ?? null)
  const firstColor = colorOf(runs[0])
  const firstAlign = box.paras[0]?.align ?? null
  return {
    shapeHasText: true,
    shapeTextBold: every((r) => r.bold === true),
    shapeTextItalic: every((r) => r.italic === true),
    shapeTextUnderline: every((r) => r.underline === true),
    shapeTextColor: every((r) => colorOf(r) === firstColor) ? firstColor : null,
    shapeTextAlign: box.paras.every((p) => (p.align ?? null) === firstAlign) ? firstAlign : null,
  }
}

export function computeFormatState(
  editor: Editor | null,
  styles?: Map<string, StyleInfo>,
  docDefaults?: DocDefaults,
): RibbonFormatState {
  if (!editor || editor.isDestroyed) return EMPTY_FORMAT_STATE
  const sub = getActiveSubEditor()
  const ed = sub ?? editor

  // Select Table / the move handle: the ribbon reads a selected table as every cell selected (Word)
  const cells = sub ? null : tableCellsSelection(editor.state)
  const tableState = cells ? editor.state.apply(editor.state.tr.setSelection(cells)) : editor.state
  const inTable = !sub && isInTable(tableState)
  let cellKey: number | null = null
  let cellHeightTwips: number | null = null
  let cellWidthPx: number | null = null
  let cellVAlign: string | null = null
  let cellTextDirection: string | null = null
  let canSplitTable = false
  if (inTable) {
    try {
      const rect = selectedRect(tableState)
      const rowNode = rect.table.maybeChild(rect.top)
      const cellPos = rect.map.map[rect.top * rect.map.width + rect.left]
      const cellNode = tableState.doc.nodeAt(rect.tableStart + cellPos)
      const colwidth = (cellNode?.attrs.colwidth as number[] | null) ?? null
      cellKey = rect.tableStart * 100000 + cellPos
      cellHeightTwips = (rowNode?.attrs.heightTwips as number | null) || null
      cellWidthPx = colwidth?.[0] || null
      cellVAlign = (cellNode?.attrs.vAlign as string | null) ?? null
      cellTextDirection = (cellNode?.attrs.textDirection as string | null) ?? null
      canSplitTable = rect.top > 0
    } catch {
      /* degenerate selection: leave cell fields empty like the old activeCellInfo */
    }
  }

  const protAttrs = editor.getAttributes('docProtected')
  const imageSelected = protAttrs.blockType === 'image' && !!protAttrs.imageDataUrl

  const textAttrs = ed.getAttributes('docTextStyle')
  const fonts = selectedFonts(ed, styles, docDefaults)
  // Word's font box names the slot matching the script at the caret; null = mixed
  const displayFont = (): string => {
    const { from, to } = ed.state.selection
    const sample =
      from === to
        ? ed.state.doc.textBetween(
            Math.max(0, from - 1),
            Math.min(ed.state.doc.content.size, from + 1),
          )
        : ed.state.doc.textBetween(from, Math.min(to, from + 32), ' ')
    const slot = textHasCjk(sample)
      ? fonts.fontEastAsia === ''
        ? fonts.fontLatin
        : fonts.fontEastAsia
      : fonts.fontLatin
    return slot ?? ''
  }
  const paraAttrs = sub ? ed.getAttributes('docParagraph') : paraAttrsOf(editor)
  const mainPara = paraAttrsOf(editor)

  return {
    doc: editor.state.doc,
    docEmpty: docEmptyOf(editor.state.doc),
    sub,
    editable: editor.isEditable,
    inTable,
    canMergeCells: inTable && !!mergeCells(tableState),
    canSplitCell: inTable && !!splitCell(tableState),
    canSplitTable,
    imageSelected,
    imageDataUrl: imageSelected ? str(protAttrs.imageDataUrl) : null,
    imageWrap: str(protAttrs.imageWrap),
    imageAlign: str(protAttrs.imageAlign),
    imageWidthPx: num(protAttrs.imageWidthPx),
    imageHeightPx: num(protAttrs.imageHeightPx),
    imageFlipH: !!protAttrs.imageFlipH,
    imageFlipV: !!protAttrs.imageFlipV,
    imageHasDocxIndex: protAttrs.docxIndex != null,
    imageKey: imageSelected ? editor.state.selection.from : null,
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
    ...shapeTextStateOf(
      Array.isArray(protAttrs.textboxes)
        ? (protAttrs.textboxes[0] as TextboxDisplay | undefined)
        : undefined,
    ),
    cellKey,
    cellHeightTwips,
    cellWidthPx,
    cellVAlign,
    cellTextDirection,
    bold: ed.isActive('bold'),
    italic: ed.isActive('italic'),
    underline: ed.isActive('underline'),
    strike: ed.isActive('strike'),
    vertAlign: str(textAttrs.vertAlign),
    highlight: str(textAttrs.highlight),
    textColor: textAttrs.color === 'auto' ? null : str(textAttrs.color),
    charStyleId: str(textAttrs.styleId),
    paraStyleId: str(paraAttrs.styleId),
    fontSizePt: (effectiveSizeHalfPoints(ed, styles, docDefaults) ?? 20) / 2,
    fontSizeMixed: selectedSizeMixed(ed, styles, docDefaults),
    fontFamily: displayFont(),
    ...fonts,
    headingLevel: editor.isActive('docHeading')
      ? Number(editor.getAttributes('docHeading').level ?? 1)
      : null,
    listBullet: editor.isActive('docListItem', { kind: 'bullet' }),
    listOrdered: editor.isActive('docListItem', { kind: 'ordered' }),
    align: str(paraAttrs.align),
    bidi: paraAttrs.bidi === true || paraAttrs.bidiInferred === true,
    selectionBidi: effectiveBidi(paraAttrs) || (!sub && selectionHasBidi(editor)),
    lineSpacing: typeof paraAttrs.lineSpacing === 'number' ? paraAttrs.lineSpacing : null,
    shadingFill: str(paraAttrs.shadingFill),
    paraBorders: paraAttrs.borders ?? null,
    indentLeft: Number(mainPara.indentLeft) || 0,
    indentRight: Number(mainPara.indentRight) || 0,
    spaceBefore: Number(mainPara.spaceBefore) || 0,
    spaceAfter: Number(mainPara.spaceAfter) || 0,
  }
}
