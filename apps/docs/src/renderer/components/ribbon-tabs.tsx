import { useRef, useState } from 'react'
import { imageSizeOf } from '../editor/picture-bytes'
import type { Editor, JSONContent } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { TextSelection, type Transaction } from '@tiptap/pm/state'
import { closeHistory } from '@tiptap/pm/history'
import {
  SHAPE_GALLERY_GROUPS,
  useDismissablePopover,
  wordArtSolidColor,
  type WordArtPreset,
} from '@genoffice/ui'
import {
  buildLineParagraphXml,
  buildShapeParagraphXml,
  buildTextboxParagraphXml,
  buildWordArtParagraphXml,
  LINE_KINDS,
  type HeaderFooter,
  type StyleInfo,
  type TextboxDisplay,
  type TextboxVert,
} from '@genoffice/docx-engine'
import type { DocsTabInfo } from '../../shared/ipc'
import { runUiOps } from '../ai/ops'
import { defaultParagraphStyleId, headingStyleId, type StyleMap } from '../style-gallery'
import { tableModelToPmNode } from '../editor/convert'
import { wordTableCell, wordTableModel } from '../editor/new-table'
import { insertPageBreak } from '../editor/page-break'
import { insertPictureAtSelection, settleCaretAfterPicture } from '../editor/insert-picture'
import { isStraightLineKind } from '../editor/shape-svg'
import { HIGHLIGHTER_ALPHA, isPenKind, type InkPenKind, type InkTool } from '../editor/ink'
import { colorsFor, newInkPreset, widthsFor, type InkPreset } from '../editor/ink-presets'
import { autoCaptionInserted } from '../editor/auto-caption'
import { t, useI18n, type StringKey } from '../i18n/locale'
import iconEditor from '../assets/icon-editor.png'
import iconTranslate from '../assets/icon-translate.png'
import type { BalloonMode, RevisionDisplayMode } from '../editor/revision-view'
import type { RevisionFilter, RevisionScope } from '../editor/revision-filters'
import type { ActionCommandId } from '../../shared/commands'
import { RbCommand, useCommandGreyed, useRunCommand } from './ribbon-command'
import { TRANSLATE_LANGS, ribbonLangKey } from './translate-langs'
import {
  IconAccept,
  IconCaret,
  IconComment,
  IconCommentNext,
  IconCommentPrev,
  IconComments,
  IconCompare,
  IconEraser,
  IconHighlighterPen,
  IconPalette,
  IconPen,
  IconPencilInk,
  IconPlus,
  IconLock,
  IconMoon,
  IconNavPane,
  IconOutlineView,
  IconStoryView,
  IconPageWidth,
  IconGridlines,
  IconNewWindow,
  IconPrintLayout,
  IconEye,
  IconReadMode,
  IconRuler,
  IconSpellcheck,
  IconSplit,
  IconSwitchWindows,
  IconRedo,
  IconReject,
  IconTrackChanges,
  IconTrash,
  IconUndo,
  IconWebLayout,
  IconWholePage,
  IconMultiPage,
  IconZoom,
  IconZoom100,
  IconZoomIn,
  IconZoomOut,
  IconBook,
  IconWordCount,
  IconGlobe,
  IconResolveComment,
  IconReadOnlyDoc,
  IconDraftView,
  IconFocusView,
  IconArrangeAll,
  IconMergeWindows,
} from './icons'

export { ChartInsertModal, InsertTab, LinkInsertModal, TableInsertModal } from './ribbon-insert-tab'
export { DesignTab } from './ribbon-design-tab'
export { LayoutTab } from './ribbon-layout-tab'
export { MailingsTab } from './ribbon-mailings-tab'
export { ReferencesTab } from './ribbon-references-tab'

/** icon size for the big icon-over-label ribbon buttons (slides ribbon parity) */
export const BIG = 28

/* ---------- shared helpers ---------- */

export type SetDropdown = (updater: (prev: string | null) => string | null) => void

export const toggleDropdown = (setDropdown: SetDropdown, key: string) =>
  setDropdown((prev) => (prev === key ? null : key))

/**
 * Apply paragraph-level attrs to every paragraph in the selection (the
 * setParagraphAttrs op: headings, list items and table-cell paragraphs alike;
 * `align` also lands on selected images as their w:jc).
 */
/** Word's line-spacing menu: 12 pt before/after when the paragraph has none, none when it has some */
export function toggleParaSpace(
  editor: Editor,
  side: 'spaceBefore' | 'spaceAfter',
  currentTwips: number,
): void {
  setParaAttrs(editor, { [side]: currentTwips > 0 ? 0 : 240, [`${side}Auto`]: null })
}

export function setParaAttrs(
  editor: Editor,
  attrs: Record<string, unknown>,
  /// Explicit target range: blur-committed inputs capture the selection at
  /// focus time — by blur, a click may already have moved the live selection
  /// to another paragraph.
  range?: { from: number; to: number },
): void {
  const size = editor.state.doc.content.size
  const target = range
    ? { range: { from: Math.min(range.from, size), to: Math.min(range.to, size) } }
    : { scope: 'selection' as const }
  runUiOps(editor, [{ op: 'setParagraphAttrs', target, attrs }], { focus: !range })
}

/** direct paragraph formatting dropped by Word's Ctrl+Q (the style's own values then show through) */
const DIRECT_PARA_ATTRS: Record<string, unknown> = {
  align: null,
  lineSpacing: null,
  lineRule: null,
  lineRawTwips: null,
  indentLeft: null,
  indentRight: null,
  indentFirstLine: null,
  indentFirstLineChars: null,
  spaceBefore: null,
  spaceAfter: null,
  spaceBeforeAuto: null,
  spaceAfterAuto: null,
  contextualSpacing: null,
  suppressAutoHyphens: null,
  mirrorIndents: null,
  snapToGrid: null,
  keepNext: null,
  keepLines: null,
  widowControl: null,
  suppressLineNumbers: null,
  shadingFill: null,
  borders: null,
  borderLines: null,
  tabStops: null,
  dropCap: null,
  pageBreakBefore: false,
  overflowPunctDirect: null,
  wordWrapDirect: null,
  kinsoku: null,
  topLinePunct: null,
  textAlignment: null,
  wordWrap: null,
  autoSpaceDE: null,
  autoSpaceDN: null,
  // the page's switches fall back to Word's defaults; a style's own value shows again on reopen
  overflowPunct: null,
  autoSpace: null,
}

/** Word's Ctrl+Q: reset the paragraph to its style, keeping the text and its character formatting */
export function clearParagraphFormatting(editor: Editor): void {
  setParaAttrs(editor, { ...DIRECT_PARA_ATTRS })
}

/**
 * Word's Clear Formatting entry: direct character formatting off (a collapsed
 * caret takes its whole paragraph), then the paragraphs back to the default
 * style with their direct paragraph formatting dropped. A textbox sub-editor
 * only clears the runs.
 */
export function clearAllFormatting(editor: Editor, sub: Editor | null, styles?: StyleMap): void {
  const ed = sub ?? editor
  const { $from, empty } = ed.state.selection
  const chain = ed.chain().focus()
  if (empty && $from.parent.isTextblock)
    chain.setTextSelection({ from: $from.start(), to: $from.end() })
  chain.unsetAllMarks()
  if (empty) chain.setTextSelection($from.pos)
  chain.run()
  if (sub) return
  applyParagraphStyleId(editor, defaultParagraphStyleId(styles) ?? null, null)
  clearParagraphFormatting(editor)
}

/** Word's Quick Style gallery keys on the shortcuts (Opt+Cmd+0..3) */
export function applyParagraphStyle(
  editor: Editor,
  key: 'p' | 'h1' | 'h2' | 'h3',
  styles?: StyleMap,
): void {
  const level = key === 'p' ? null : Number(key.slice(1))
  const styleId = level ? headingStyleId(styles, level) : defaultParagraphStyleId(styles)
  applyParagraphStyleId(editor, styleId ?? null, level)
}

/** mark types Word counts as direct character formatting when a paragraph style is applied */
const DIRECT_MARKS = ['bold', 'italic', 'underline', 'strike'] as const
const DIRECT_TEXT_STYLE_ATTRS = ['color', 'sizeHalfPoints', 'font', 'fontAscii'] as const

/**
 * Word's rule for direct character formatting under a new paragraph style: a
 * property covering more than half of the paragraph's text is dropped from the
 * whole paragraph (it was standing in for the style), a smaller run keeps it.
 */
export function shedMajorityDirectFormatting(tr: Transaction, pos: number): void {
  const node = tr.doc.nodeAt(pos)
  if (!node?.isTextblock) return
  const total = node.textContent.length
  if (total === 0) return
  const covered = new Map<string, number>()
  node.forEach((child) => {
    if (!child.isText) return
    const len = child.text!.length
    for (const m of child.marks) {
      if ((DIRECT_MARKS as readonly string[]).includes(m.type.name)) {
        covered.set(m.type.name, (covered.get(m.type.name) ?? 0) + len)
      } else if (m.type.name === 'docTextStyle') {
        for (const a of DIRECT_TEXT_STYLE_ATTRS) {
          if (m.attrs[a] != null) covered.set(`ts:${a}`, (covered.get(`ts:${a}`) ?? 0) + len)
        }
      }
    }
  })
  const shed = [...covered].filter(([, n]) => n * 2 > total).map(([k]) => k)
  if (shed.length === 0) return
  const start = pos + 1
  const end = pos + node.nodeSize - 1
  const schema = tr.doc.type.schema
  for (const k of shed) if (!k.startsWith('ts:')) tr.removeMark(start, end, schema.marks[k])
  const attrsToClear = shed.filter((k) => k.startsWith('ts:')).map((k) => k.slice(3))
  if (attrsToClear.length === 0) return
  const type = schema.marks.docTextStyle
  const jobs: Array<{ from: number; to: number; attrs: Record<string, unknown> | null }> = []
  node.forEach((child, offset) => {
    const m = child.marks.find((mm) => mm.type === type)
    if (!m || !attrsToClear.some((a) => m.attrs[a] != null)) return
    const attrs = { ...m.attrs }
    for (const a of attrsToClear) attrs[a] = null
    const keep = Object.values(attrs).some((v) => v != null && v !== false)
    jobs.push({
      from: start + offset,
      to: start + offset + child.nodeSize,
      attrs: keep ? attrs : null,
    })
  })
  for (const job of jobs) {
    tr.removeMark(job.from, job.to, type)
    if (job.attrs) tr.addMark(job.from, job.to, type.create(job.attrs))
  }
}

/**
 * Put a paragraph style on every paragraph the selection touches: a heading
 * style makes docHeading nodes, another style turns headings back into
 * paragraphs; list items keep their numbering unless they become headings.
 * Not for textbox sub-editors (no docHeading in their schema).
 */
export function applyParagraphStyleId(
  editor: Editor,
  styleId: string | null,
  headingLevel: number | null,
): void {
  editor
    .chain()
    .focus()
    .command(({ tr, state }) => {
      const { from, to } = tr.selection
      const schema = state.schema
      const targets: Array<{ pos: number; node: PmNode }> = []
      tr.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isTextblock) return true
        if (
          node.type.name === 'docParagraph' ||
          node.type.name === 'docHeading' ||
          node.type.name === 'docListItem'
        )
          targets.push({ pos, node })
        return false
      })
      for (const { pos, node } of targets) {
        let type = headingLevel
          ? schema.nodes.docHeading
          : node.type.name === 'docHeading'
            ? schema.nodes.docParagraph
            : node.type
        const $p = tr.doc.resolve(pos)
        if (type !== node.type && !$p.parent.canReplaceWith($p.index(), $p.index() + 1, type))
          type = node.type
        const attrs: Record<string, unknown> = { ...node.attrs, styleId }
        if (type === schema.nodes.docHeading) {
          attrs.level = headingLevel ?? node.attrs.level
          attrs.outlineOnly = null
        }
        // Clicking the already-active gallery card must be a no-op:
        // setNodeMarkup always emits a ReplaceAroundStep even with identical
        // attrs, which dirties the doc and triggers a whole save/reparse/
        // pagination pass — the "freeze, then nothing changed" of issue #2211.
        const unchanged =
          type === node.type &&
          (node.attrs.styleId ?? null) === styleId &&
          (type !== schema.nodes.docHeading ||
            ((node.attrs.level ?? 1) === (attrs.level ?? 1) &&
              (node.attrs.outlineOnly ?? null) === null))
        if (unchanged) continue
        tr.setNodeMarkup(pos, type === node.type ? undefined : type, attrs)
        shedMajorityDirectFormatting(tr, pos)
      }
      return tr.steps.length > 0
    })
    .run()
}

/**
 * Gallery / Styles pane click: a character style toggles on the selection (the
 * textbox sub-editor when one is active), a paragraph style goes on the
 * paragraphs. A fallback entry the document does not define sets the node type only.
 */
export function applyGalleryStyle(
  editor: Editor,
  sub: Editor | null,
  info: StyleInfo,
  styles: StyleMap | undefined,
  activeCharStyleId: string | null,
): void {
  if (info.type === 'character') {
    const ed = sub ?? editor
    if (activeCharStyleId === info.styleId) ed.chain().focus().unsetMark('docTextStyle').run()
    else ed.chain().focus().setMark('docTextStyle', { styleId: info.styleId }).run()
    return
  }
  if (sub) return
  const styleId = styles?.has(info.styleId) ? info.styleId : null
  applyParagraphStyleId(editor, styleId, info.headingLevel ?? null)
}

/** attrs of the paragraph-like node at the cursor */
export function activeParaAttrs(editor: Editor): Record<string, unknown> {
  if (editor.isActive('docHeading')) return editor.getAttributes('docHeading')
  if (editor.isActive('docListItem')) return editor.getAttributes('docListItem')
  return editor.getAttributes('docParagraph')
}

export interface TabProps {
  editor: Editor
  hasDoc: boolean
  dropdown: string | null
  setDropdown: SetDropdown
}

/* ================= Insert ================= */

const MAX_IMAGE_WIDTH_PX = 620 // ~content width of a US Letter page at 96dpi

export { imageSizeOf }

/* insert commands shared by the ribbon and the native application menu */

/** Word's Insert Table dialog column limit */
export const MAX_TABLE_COLS = 63
/** row cap keeps a single insert from freezing layout (Word allows 32767) */
export const MAX_TABLE_ROWS = 200

export function insertTableAt(
  editor: Editor,
  rows: number,
  cols: number,
  /** the caret's paragraph is a typed +--+--+ line the table takes the place of, in one undo step */
  replaceLine = false,
): void {
  rows = Math.min(MAX_TABLE_ROWS, Math.max(1, Math.round(rows)))
  cols = Math.min(MAX_TABLE_COLS, Math.max(1, Math.round(cols)))
  const before = editor.state.doc
  const table = wordTableModel(
    Array.from({ length: rows }, () => Array.from({ length: cols }, () => wordTableCell(['']))),
    cols,
  )
  // inside a cell a top-level docTable insert would split the outer table
  // — Word semantics is a nested child table at the end of the cell
  const { $from } = editor.state.selection
  const typed = replaceLine ? $from.parent.content.size : 0
  const start = () => {
    const chain = editor.chain().focus()
    if (typed === 0) return chain
    // its own undo step: undo brings back the typed line, not the typing before it
    return chain
      .command(({ tr }) => !!closeHistory(tr))
      .deleteRange({ from: $from.start(), to: $from.end() })
  }
  for (let depth = $from.depth; depth > 0; depth--) {
    const name = $from.node(depth).type.name
    if (name === 'docTableCell' || name === 'docTableHeader') {
      start()
        .insertContentAt($from.end(depth) - typed, [
          { type: 'docNestedTable', attrs: { model: table } },
          { type: 'docParagraph' },
        ])
        .run()
      return
    }
  }
  const node = tableModelToPmNode(table)
  // Word: an empty paragraph stays below the new table (insertContent would
  // swallow it); the caret lands in the first cell either way
  const block = $from.depth > 0 ? $from.node(1) : null
  const at =
    block?.isTextblock && block.content.size === typed && (typed === 0 || $from.depth === 1)
      ? $from.before(1)
      : null
  const chain = start()
  if (at != null) chain.insertContentAt(at, node)
  else chain.insertContent(node)
  const from = at ?? ($from.depth > 0 ? $from.before() : 0)
  chain
    .command(({ tr }) => {
      // a table inserted mid-paragraph sits after the split-off head
      let table = -1
      tr.doc.nodesBetween(from, tr.doc.content.size, (n, pos) => {
        if (table < 0 && n.type.name === 'docTable' && pos >= from) table = pos
        return table < 0 && !n.isTextblock
      })
      if (table >= 0) tr.setSelection(TextSelection.near(tr.doc.resolve(table + 1)))
      return true
    })
    .run()
  autoCaptionInserted(editor, 'table', before)
}

/** Insert an inline image from a dataURL at the cursor (shared by paste/dialog; size scaled to content width) */
export async function insertImageFromDataUrl(
  editor: Editor,
  dataUrl: string,
  label = t('ribbonPicture'),
): Promise<boolean> {
  const m = /^data:([^;]+);base64,(.*)$/s.exec(dataUrl)
  if (!m) return false
  const mime = m[1]
  if (!/^image\/(png|jpeg|gif)$/.test(mime)) return false
  try {
    const natural = await imageSizeOf(dataUrl)
    const scale = Math.min(1, MAX_IMAGE_WIDTH_PX / natural.width)
    insertPictureAtSelection(editor, {
      dataUrl,
      base64: m[2],
      mime,
      widthPx: Math.round(natural.width * scale),
      heightPx: Math.round(natural.height * scale),
      label,
    })
    settleCaretAfterPicture(editor)
    return true
  } catch {
    return false
  }
}

export async function insertImageViaDialog(editor: Editor): Promise<void> {
  const picked = await window.desktop.pickImage()
  if (!picked) return
  await insertImageFromDataUrl(
    editor,
    `data:${picked.mime};base64,${picked.base64}`,
    t('ribbonPictureLabel', { name: picked.name }),
  )
}

/** 5 cm × 3 cm default textbox size in EMU (1 cm = 360000 EMU) */
export const TEXTBOX_WIDTH_EMU = 1800000
export const TEXTBOX_HEIGHT_EMU = 1080000

/** Default TextboxDisplay model for a freshly inserted empty textbox */
function emptyTextboxDisplay(
  widthEmu: number,
  heightEmu: number,
  vert?: TextboxVert,
): TextboxDisplay {
  return {
    fill: 'FFFFFF',
    borderColor: '000000',
    widthPx: Math.round(widthEmu / 9525),
    heightPx: Math.round(heightEmu / 9525),
    ...(vert ? { vert } : {}),
    paras: [{ runs: [{ text: '' }] }],
  }
}

/**
 * Insert a floating text box (wp:anchor + wps:wsp) at the current cursor, or at
 * an explicit top-level position with an explicit size (Draw Text Box).
 * Returns the position the block was inserted at (null if the insert failed).
 */
export function insertTextboxAt(
  editor: Editor,
  opts?: { widthEmu?: number; heightEmu?: number; atPos?: number; vert?: TextboxVert },
): number | null {
  const widthEmu = opts?.widthEmu ?? TEXTBOX_WIDTH_EMU
  const heightEmu = opts?.heightEmu ?? TEXTBOX_HEIGHT_EMU
  const xml = buildTextboxParagraphXml({
    widthEmu,
    heightEmu,
    id: Math.floor(Math.random() * 900000) + 100000,
    ...(opts?.vert ? { vert: opts.vert } : {}),
  })
  const content = {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('ribbonTextBox'),
      genXml: xml,
      textboxes: [emptyTextboxDisplay(widthEmu, heightEmu, opts?.vert)],
    },
  }
  // top-level insert: a plain insertContent would replace a selected floating
  // node and fails silently from inside a table cell
  const { $from } = editor.state.selection
  const position = opts?.atPos ?? ($from.depth > 0 ? $from.after(1) : editor.state.selection.to)
  return editor.chain().focus().insertContentAt(position, content).run() ? position : null
}

/** Put the caret into the (first) text box of the floating node at `pos`. */
export function focusTextboxEditorAt(editor: Editor, pos: number): void {
  const wrapper = editor.view.nodeDOM(pos) as HTMLElement | null
  const box = wrapper?.querySelector('.doc-textbox') as HTMLElement | null
  const sub = box?.querySelector('.doc-textbox-editor') as HTMLElement | null
  if (!box || !sub) return
  // a fresh box sits in object mode; the double-click path is what turns its editor on
  box.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))
  sub.focus()
}

/**
 * Shape gallery for the picker dropdown: the full cross-app shared groups
 * (slides parity). Line/connector kinds insert stroke-only wps shapes
 * (buildLineParagraphXml); filled presets insert prstGeom shapes.
 */
export const DOC_SHAPE_GROUPS = SHAPE_GALLERY_GROUPS

const DOC_SHAPES = DOC_SHAPE_GROUPS.flatMap((g) => g.shapes)

/** Display label for a prst inserted from the gallery. */
export function shapeLabel(prst: string): string {
  const def = DOC_SHAPES.find((s) => s.prst === prst)
  return def ? t(def.labelKey as StringKey) : prst
}

/**
 * Insert a block beside the current top-level block. Ribbon actions can be
 * invoked while the caret is nested in a table cell, where inserting a block
 * directly at the selection is invalid and TipTap otherwise fails silently.
 */
export function insertTopLevelBlockAtSelection(editor: Editor, content: JSONContent): boolean {
  const { $from } = editor.state.selection
  const position = $from.depth > 0 ? $from.after(1) : editor.state.selection.to
  return editor.chain().focus().insertContentAt(position, content).run()
}

/**
 * Insert a floating preset shape (wps:wsp with prstGeom) at the cursor, or at
 * an explicit top-level doc position with an explicit size (shape draw mode).
 * Returns the position the block was inserted at (null if the insert failed).
 */
export function insertShapeAt(
  editor: Editor,
  prst: string,
  opts?: { widthEmu?: number; heightEmu?: number; atPos?: number },
): number | null {
  if (prst in LINE_KINDS) return insertLineAt(editor, prst, opts)
  const widthEmu = opts?.widthEmu ?? 1800000
  const heightEmu = opts?.heightEmu ?? 1080000
  const id = Math.floor(Math.random() * 900000) + 100000
  const xml = buildShapeParagraphXml({
    prst,
    widthEmu,
    heightEmu,
    id,
    // default Office blue fill + slightly darker border
    fillHex: '4472C4',
    borderHex: '2F5496',
    withTextbox: true,
  })
  // mirrors what buildShapeParagraphXml just wrote: centered both ways, and the
  // light text the shape style's a:fontRef resolves to. Without this the shape
  // reads top-left and black until the file is saved and reopened.
  const textbox: TextboxDisplay = {
    fill: '4472C4',
    borderColor: '2F5496',
    widthPx: Math.round(widthEmu / 9525),
    heightPx: Math.round(heightEmu / 9525),
    prst,
    shapeId: String(id),
    vAlign: 'center',
    textColor: 'FFFFFF',
    paras: [{ runs: [{ text: '' }], align: 'center' }],
  }
  const content = {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('ribbonShapeLabel', { name: shapeLabel(prst) }),
      genXml: xml,
      textboxes: [textbox],
    },
  }
  const { $from } = editor.state.selection
  const position = opts?.atPos ?? ($from.depth > 0 ? $from.after(1) : editor.state.selection.to)
  return editor.chain().focus().insertContentAt(position, content).run() ? position : null
}

/** Word's horizontal-line extent (12 px grab band); straight lines always save this cy. */
const LINE_HEIGHT_EMU = 114300

/**
 * Insert a floating stroke-only line/connector (noFill wps:wsp) at the cursor
 * or an explicit position. Straight kinds ignore the drawn height and land as
 * a level line (the docx model stores Word's zero-ish-height extent); bent and
 * curved connectors keep the drawn box.
 */
function insertLineAt(
  editor: Editor,
  kind: string,
  opts?: { widthEmu?: number; heightEmu?: number; atPos?: number },
): number | null {
  const widthEmu = opts?.widthEmu ?? 1800000
  const heightEmu = isStraightLineKind(kind) ? LINE_HEIGHT_EMU : (opts?.heightEmu ?? 1080000)
  const xml = buildLineParagraphXml({
    kind,
    widthEmu,
    heightEmu,
    id: Math.floor(Math.random() * 900000) + 100000,
    colorHex: '000000',
  })
  // Mirror what parse.ts' lineBoxOf yields on reopen: read-only display box,
  // stroke color on borderColor, zero insets.
  const textbox: TextboxDisplay = {
    borderColor: '000000',
    widthPx: Math.round(widthEmu / 9525),
    heightPx: Math.round(heightEmu / 9525),
    prst: kind,
    paras: [],
    readOnly: true,
    insetTopPx: 0,
    insetRightPx: 0,
    insetBottomPx: 0,
    insetLeftPx: 0,
  }
  const content = {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('ribbonShapeLabel', { name: shapeLabel(kind) }),
      genXml: xml,
      textboxes: [textbox],
    },
  }
  const { $from } = editor.state.selection
  const position = opts?.atPos ?? ($from.depth > 0 ? $from.after(1) : editor.state.selection.to)
  return editor.chain().focus().insertContentAt(position, content).run() ? position : null
}

/** ~7.5 cm × 2 cm default size for WordArt in EMU */
const WORDART_WIDTH_EMU = 2700000
const WORDART_HEIGHT_EMU = 720000

/** Insert a floating WordArt text box at the cursor. */
export function insertWordArtAt(editor: Editor, preset: WordArtPreset): void {
  // The saved run can only carry a solid color; light fills fall back to the
  // outline color so the text stays readable when the file is reopened.
  const solidHex = wordArtSolidColor(preset).replace('#', '')
  const xml = buildWordArtParagraphXml({
    colorHex: solidHex,
    italic: preset.italic,
    widthEmu: WORDART_WIDTH_EMU,
    heightEmu: WORDART_HEIGHT_EMU,
    id: Math.floor(Math.random() * 900000) + 100000,
  })
  const textbox: TextboxDisplay = {
    // no background fill; shape border is also absent (noFill)
    widthPx: Math.round(WORDART_WIDTH_EMU / 9525),
    heightPx: Math.round(WORDART_HEIGHT_EMU / 9525),
    wordArtId: preset.id,
    paras: [
      {
        runs: [
          {
            text: t('ribbonWordArtDefaultText'),
            color: solidHex,
            bold: true,
            italic: preset.italic,
            sizeHalfPoints: 72,
          },
        ],
        align: 'center',
      },
    ],
  }
  // top-level insert: a plain insertContent would replace a selected floating
  // node and fails silently from inside a table cell
  insertTopLevelBlockAtSelection(editor, {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('ribbonWordArtLabel', { name: t(preset.nameKey as StringKey) }),
      genXml: xml,
      textboxes: [textbox],
    },
  })
}

export function insertPageBreakAt(editor: Editor): void {
  insertPageBreak(editor)
}

/**
 * Word's "Blank Page": one empty paragraph that starts its own page, and —
 * only when content follows it — a break pushed onto that following block so
 * it starts the page after. Unconditionally inserting two break paragraphs
 * turned a 1-page document into 3 pages.
 */
export function insertBlankPageAt(editor: Editor): void {
  editor
    .chain()
    .focus()
    .insertContent({ type: 'docParagraph', attrs: { pageBreakBefore: true } })
    .run()
  // locate the paragraph just inserted: the caret block, or the block before
  // it when insertContent left the caret in the split-off remainder
  const { doc, selection } = editor.state
  const isBlankBreak = (index: number): boolean => {
    if (index < 0 || index >= doc.childCount) return false
    const n = doc.child(index)
    return n.type.name === 'docParagraph' && n.childCount === 0 && n.attrs.pageBreakBefore === true
  }
  const caretIndex = selection.$to.index(0)
  const blankIndex = isBlankBreak(caretIndex)
    ? caretIndex
    : isBlankBreak(caretIndex - 1)
      ? caretIndex - 1
      : -1
  if (blankIndex < 0) return
  let blankPos = 0
  for (let i = 0; i < blankIndex; i++) blankPos += doc.child(i).nodeSize
  // Word leaves the caret on the new blank page
  editor.commands.setTextSelection(blankPos + 1)
  const followIndex = blankIndex + 1
  if (followIndex >= doc.childCount) return // blank page is the last page
  const follow = doc.child(followIndex)
  const followPos = blankPos + doc.child(blankIndex).nodeSize
  if ('pageBreakBefore' in follow.attrs) {
    if (!follow.attrs.pageBreakBefore) {
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(followPos, undefined, {
          ...follow.attrs,
          pageBreakBefore: true,
        }),
      )
    }
    return
  }
  // following block can't carry the property (e.g. a table): spacer fallback
  const spacer = editor.state.schema.nodes.docParagraph!.createAndFill({ pageBreakBefore: true })
  if (spacer) editor.view.dispatch(editor.state.tr.insert(followPos, spacer))
}

export interface InsertTabProps extends TabProps {
  header: HeaderFooter | null
  footer: HeaderFooter | null
  /** the Review-tab gate pair: commenting survives the comments-only restriction */
  isProtected: boolean
  commentsAllowed: boolean
}

/** One-time "AI rewrites the whole document" acknowledgement */
export const AI_REWRITE_ACK_KEY = 'docs-ai-rewrite-ack'

export type { BalloonMode, RevisionDisplayMode }

interface ReviewTabProps extends TabProps {
  onAiPreset: (instruction: string) => void
  commentCount: number
  /** unresolved root comments; 0 disables the AI resolve-comments action */
  openCommentCount: number
  resolvedCommentCount: number
  /** the caret sits in a comment anchor: Delete ▾ › Delete acts on that thread */
  commentAtCaret: boolean
  trackChanges: boolean
  /** native check-as-you-type spellcheck (red squiggle) */
  spellcheck: boolean
  revisionDisplay: RevisionDisplayMode
  /** Show Markup → Balloons */
  balloonMode: BalloonMode
  revisionCount: number
  /** Show Markup: unticked kinds / reviewers render as accepted */
  revisionFilter: RevisionFilter
  /** reviewers in first-appearance order (Specific People, Accept All by …) */
  revisionAuthors: string[]
  showReviewPane: boolean
  /** an editing restriction or write lock makes the body read-only */
  isProtected: boolean
  /** comments restriction: adding comments stays allowed although the body is read-only */
  commentsAllowed: boolean
  /** trackedChanges restriction: the recorder is forced on (toggle and accept/reject disabled) */
  trackChangesForced: boolean
  /** any protection is configured (highlights the Protect Document button) */
  protectActive: boolean
}

export function ReviewTab({
  editor,
  hasDoc,
  dropdown,
  setDropdown,
  onAiPreset,
  commentCount,
  openCommentCount,
  resolvedCommentCount,
  commentAtCaret,
  trackChanges,
  spellcheck,
  revisionDisplay,
  balloonMode,
  revisionCount,
  revisionFilter,
  revisionAuthors,
  showReviewPane,
  isProtected,
  commentsAllowed,
  trackChangesForced,
  protectActive,
}: ReviewTabProps) {
  const { t } = useI18n()
  const unknownAuthor = t('editorUnknownAuthor')
  const hasSelection = !!editor && !editor.state.selection.empty
  const runCommand = useRunCommand()
  // commenting outlives the comments-only restriction, which the protection terms below answer
  const commentOff = useCommandGreyed('insert-comment', { editable: true })
  const showCommentsOff = useCommandGreyed('show-comments')
  /** a pick from a menu: run the command, close the menu */
  const revisionDisplayChoices: Array<[RevisionDisplayMode, string]> = [
    ['simple', t('ribbonRevDisplaySimple')],
    ['all', t('ribbonRevDisplayAll')],
    ['none', t('ribbonRevDisplayNone')],
    ['original', t('ribbonRevDisplayOriginal')],
  ]
  const pickCommand = (id: ActionCommandId, payload?: string) => {
    runCommand(id, payload)
    setDropdown(() => null)
  }
  const revisionMenu = (action: 'accept' | 'reject') => {
    const key = (suffix: string) =>
      `${action === 'accept' ? 'ribbonAccept' : 'ribbonReject'}${suffix}` as StringKey
    const pick = (scope: RevisionScope) => () =>
      pickCommand(
        action === 'accept' ? 'accept-revision' : 'reject-revision',
        typeof scope === 'string' ? scope : JSON.stringify(scope),
      )
    return (
      <div data-rb-panel="" className="layout-menu">
        <button onClick={pick('next')}>{t(key('Next'))}</button>
        <button disabled={revisionCount === 0} onClick={pick('shown')}>
          {t(key('Shown'))}
        </button>
        <button onClick={pick('all')}>{t(key('All'))}</button>
        <button disabled={!hasSelection} onClick={pick('selection')}>
          {t(key('Selection'))}
        </button>
        {revisionAuthors.length > 1 && <div className="rb-menu-sep" />}
        {revisionAuthors.length > 1 &&
          revisionAuthors.map((author) => (
            <button key={author} onClick={pick({ author })}>
              {t(key('By'), { author: author || unknownAuthor })}
            </button>
          ))}
      </div>
    )
  }
  // One-time acknowledgement before whole-document AI rewrites:
  // Editor / Translate send the full document to the agent, consume credits and
  // may rewrite everything — say so once before the first run.
  const confirmAiRewrite = () => {
    if (localStorage.getItem(AI_REWRITE_ACK_KEY) === '1') return true
    if (!window.confirm(t('ribbonAiRewriteConfirm'))) return false
    localStorage.setItem(AI_REWRITE_ACK_KEY, '1')
    return true
  }
  // With a range selection the rewrite scopes to the selection (no whole-document ack needed)
  const hasRangeSelection = () => !editor.state.selection.empty
  return (
    <>
      {/* Word: Proofing (Editor) sits leftmost */}
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={`${t('ribbonEditorTip')} — ${t('ribbonAiCreditNote')}`}
            onClick={() => {
              if (hasRangeSelection()) onAiPreset(t('ribbonEditorSelectionPrompt'))
              else if (confirmAiRewrite()) onAiPreset(t('ribbonEditorPrompt'))
            }}
          >
            <span className="rb-big-icon">
              <span className="ai-feature-icon" aria-hidden="true">
                <img src={iconEditor} width={22} height={22} alt="" />
              </span>
            </span>
            <span>{t('ribbonEditorBtn')}</span>
          </button>
          <div className="rb-col" data-wide="1" data-collapse="2">
            <button
              className={`rb-small ${spellcheck ? 'active' : ''}`}
              disabled={!hasDoc}
              data-tip={t('ribbonSpellcheckTip')}
              onClick={() => runCommand('toggle-spellcheck')}
            >
              <IconSpellcheck size={14} />
              <span className="rb-small-label">{t('ribbonSpellcheckBtn')}</span>
            </button>
            <RbCommand cmd="tools-thesaurus" look="small">
              <IconBook size={14} />
            </RbCommand>
            <RbCommand cmd="word-count" look="small">
              <IconWordCount size={14} />
            </RbCommand>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupProofing')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={`${t('ribbonTranslateTip')} — ${t('ribbonAiCreditNote')}`}
              onClick={() => toggleDropdown(setDropdown, 'translate')}
            >
              <span className="rb-big-icon">
                <span className="ai-feature-icon" aria-hidden="true">
                  <img src={iconTranslate} width={22} height={22} alt="" />
                </span>
                <IconCaret />
              </span>
              <span>{t('ribbonTranslate')}</span>
            </button>
            {dropdown === 'translate' && (
              <div data-rb-panel="" className="layout-menu">
                {TRANSLATE_LANGS.map(({ code }) => (
                  <button
                    key={code}
                    onClick={() => {
                      setDropdown(() => null)
                      if (hasRangeSelection()) {
                        onAiPreset(
                          t('ribbonTranslateSelectionPrompt', { lang: t(ribbonLangKey(code)) }),
                        )
                      } else if (confirmAiRewrite()) {
                        onAiPreset(t('ribbonTranslatePrompt', { lang: t(ribbonLangKey(code)) }))
                      }
                    }}
                  >
                    {t('ribbonTranslateTo', { lang: t(ribbonLangKey(code)) })}
                  </button>
                ))}
              </div>
            )}
          </div>
          <RbCommand cmd="tools-language" look="big">
            <IconGlobe size={BIG} />
          </RbCommand>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupLanguage')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={commentOff || (isProtected && !commentsAllowed)}
            data-tip={commentOff ? t('ribbonNewCommentSelectTip') : t('ribbonNewCommentTip')}
            onClick={() => runCommand('insert-comment')}
          >
            <span className="rb-big-icon">
              <IconComment size={BIG} />
            </span>
            <span>{t('ribbonNewComment')}</span>
          </button>
          <div className="rb-col" data-wide="2">
            <div className="rb-split-wrap">
              <button
                className="rb-small"
                disabled={!hasDoc || commentCount === 0 || (isProtected && !commentsAllowed)}
                data-tip={t('ribbonDeleteCommentTip')}
                onClick={() => toggleDropdown(setDropdown, 'deleteComment')}
              >
                <IconTrash size={14} />
                <span className="rb-small-label">{t('ribbonDeleteComment')}</span>
                <IconCaret />
              </button>
              {dropdown === 'deleteComment' && (
                <div data-rb-panel="" className="layout-menu">
                  <button disabled={!commentAtCaret} onClick={() => pickCommand('delete-comment')}>
                    {t('ribbonDeleteComment')}
                  </button>
                  <button
                    disabled={resolvedCommentCount === 0}
                    onClick={() => pickCommand('delete-comments', 'resolved')}
                  >
                    {t('ribbonDeleteResolvedComments')}
                  </button>
                  <button onClick={() => pickCommand('delete-comments', 'all')}>
                    {t('ribbonDeleteAllComments')}
                  </button>
                </div>
              )}
            </div>
            <button
              className="rb-small"
              disabled={!commentAtCaret || (isProtected && !commentsAllowed)}
              data-tip={t('ribbonResolve')}
              onClick={() => runCommand('comment-resolve')}
            >
              <IconResolveComment size={14} />
              <span className="rb-small-label">{t('ribbonResolve')}</span>
            </button>
            <button
              className="rb-small"
              disabled={!hasDoc || openCommentCount === 0}
              data-tip={t('ribbonPrevCommentTip')}
              onClick={() => runCommand('comment-previous')}
            >
              <IconCommentPrev size={14} />
              <span className="rb-small-label">{t('ribbonPrevComment')}</span>
            </button>
          </div>
          <div className="rb-col rb-pair" data-wide="2">
            <button
              className="rb-small"
              disabled={!hasDoc || openCommentCount === 0}
              data-tip={t('ribbonNextCommentTip')}
              onClick={() => runCommand('comment-next')}
            >
              <IconCommentNext size={14} />
              <span className="rb-small-label">{t('ribbonNextComment')}</span>
            </button>
            <button
              className="rb-small"
              disabled={showCommentsOff}
              data-tip={t('ribbonShowCommentsTip', { count: commentCount })}
              onClick={() => runCommand('show-comments')}
            >
              <IconComments size={14} />
              <span className="rb-small-label">{t('ribbonShowComments')}</span>
            </button>
          </div>
          <button
            className="rb-big"
            disabled={!hasDoc || openCommentCount === 0}
            data-tip={`${t('ribbonAiCommentsTip', { count: openCommentCount })} — ${t('ribbonAiCreditNote')}`}
            onClick={() => onAiPreset(t('ribbonAiCommentsPrompt'))}
          >
            <span className="rb-big-icon">
              <span className="ai-feature-icon" aria-hidden="true">
                <svg
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="1.5"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <path d="M20 11.5a7.5 7.5 0 0 1-7.5 7.5H6l-3 3V11.5a7.5 7.5 0 0 1 7.5-7.5h2A7.5 7.5 0 0 1 20 11.5z" />
                  <path
                    d="M17 14l.26.7c.34.91.5 1.37.84 1.7.33.33.79.5 1.7.84l.7.26-.7.26c-.91.34-1.37.5-1.7.84-.34.33-.5.79-.84 1.7L17 21l-.26-.7c-.34-.91-.5-1.37-.84-1.7-.33-.34-.79-.5-1.7-.84l-.7-.26.7-.26c.91-.34 1.37-.5 1.7-.84.34-.33.5-.79.84-1.7L17 14z"
                    strokeLinejoin="round"
                  />
                </svg>
              </span>
            </span>
            <span>{t('ribbonAiComments')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupComments')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className={`rb-big ${trackChanges ? 'active' : ''}`}
            disabled={!hasDoc || isProtected || trackChangesForced}
            data-tip={t('ribbonTrackChangesTip')}
            onClick={() => runCommand('track-changes')}
          >
            <span className="rb-big-icon">
              <IconTrackChanges size={BIG} />
            </span>
            <span>{t('ribbonTrackChanges')}</span>
          </button>
          <div className="rb-col rb-pair">
            <div className="rb-split-wrap">
              <button
                className={`rb-small ${revisionDisplay === 'none' || revisionDisplay === 'original' ? 'active' : ''}`}
                disabled={!hasDoc}
                aria-label={t('ribbonDisplayForReview')}
                data-tip={t('ribbonRevDisplayTip')}
                onClick={() => toggleDropdown(setDropdown, 'displayForReview')}
              >
                <IconEye size={14} />
                <span className="rb-small-label">
                  {revisionDisplayChoices.find(([mode]) => mode === revisionDisplay)?.[1]}
                </span>
                <IconCaret />
              </button>
              {dropdown === 'displayForReview' && (
                <div data-rb-panel="" className="layout-menu">
                  {revisionDisplayChoices.map(([mode, label]) => (
                    <button
                      key={mode}
                      role="menuitemradio"
                      aria-checked={revisionDisplay === mode}
                      onClick={() => pickCommand('revision-display', mode)}
                    >
                      {label}
                      {revisionDisplay === mode ? ' ✓' : ''}
                    </button>
                  ))}
                </div>
              )}
            </div>
            <div className="rb-split-wrap">
              <button
                className="rb-small"
                disabled={!hasDoc}
                data-tip={t('ribbonRevDisplay')}
                onClick={() => toggleDropdown(setDropdown, 'revDisplay')}
              >
                <IconReadMode size={14} />
                <span className="rb-small-label">{t('ribbonRevDisplay')}</span>
                <IconCaret />
              </button>
              {dropdown === 'revDisplay' && (
                <div data-rb-panel="" className="layout-menu">
                  <div className="layout-menu-head">{t('ribbonMarkupShow')}</div>
                  <button onClick={() => runCommand('markup-filter', 'insDel')}>
                    {t('ribbonMarkupInsDel')}
                    {revisionFilter.insDel ? ' ✓' : ''}
                  </button>
                  <button onClick={() => runCommand('markup-filter', 'format')}>
                    {t('ribbonMarkupFormatting')}
                    {revisionFilter.format ? ' ✓' : ''}
                  </button>
                  <div className="layout-menu-head">{t('ribbonMarkupBalloons')}</div>
                  {(
                    [
                      ['revisions', t('ribbonBalloonsRevisions')],
                      ['inline', t('ribbonBalloonsInline')],
                      ['comments', t('ribbonBalloonsComments')],
                    ] as Array<[BalloonMode, string]>
                  ).map(([mode, label]) => (
                    <button
                      key={mode}
                      role="menuitemradio"
                      aria-checked={balloonMode === mode}
                      onClick={() => pickCommand('balloon-mode', mode)}
                    >
                      {label}
                      {balloonMode === mode ? ' ✓' : ''}
                    </button>
                  ))}
                  {revisionAuthors.length > 0 && (
                    <div className="layout-menu-head">{t('ribbonMarkupPeople')}</div>
                  )}
                  {revisionAuthors.map((author) => (
                    <button
                      key={author}
                      onClick={() => runCommand('markup-filter', `author:${author}`)}
                    >
                      {author || unknownAuthor}
                      {revisionFilter.hiddenAuthors.includes(author) ? '' : ' ✓'}
                    </button>
                  ))}
                  <div className="rb-menu-sep" />
                  <button
                    data-testid="tc-options"
                    onClick={() => pickCommand('track-changes-options')}
                  >
                    {t('ribbonTcOptions')}
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupTracking')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className={`rb-big ${showReviewPane ? 'active' : ''}`}
            disabled={!hasDoc}
            data-tip={t('ribbonReviewingPaneTip')}
            onClick={() => runCommand('review-pane')}
          >
            <span className="rb-big-icon">
              <IconNavPane size={BIG} />
            </span>
            <span>{t('ribbonReviewingPane')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonReviewingPane')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc || isProtected || trackChangesForced}
              data-tip={t('ribbonAcceptTip', { count: revisionCount })}
              onClick={() => toggleDropdown(setDropdown, 'acceptRev')}
            >
              <span className="rb-big-icon">
                <IconAccept size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonAccept')}</span>
            </button>
            {dropdown === 'acceptRev' && revisionMenu('accept')}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc || isProtected || trackChangesForced}
              data-tip={t('ribbonRejectTip', { count: revisionCount })}
              onClick={() => toggleDropdown(setDropdown, 'rejectRev')}
            >
              <span className="rb-big-icon">
                <IconReject size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonReject')}</span>
            </button>
            {dropdown === 'rejectRev' && revisionMenu('reject')}
          </div>
          <div className="rb-col rb-pair" data-collapse="1">
            <button
              className="rb-small"
              disabled={!hasDoc}
              data-tip={t('ribbonPrevChangeTip')}
              onClick={() => runCommand('revision-previous')}
            >
              <IconUndo size={14} />
              <span className="rb-small-label">{t('ribbonPrevChange')}</span>
            </button>
            <button
              className="rb-small"
              disabled={!hasDoc}
              data-tip={t('ribbonNextChangeTip')}
              onClick={() => runCommand('revision-next')}
            >
              <IconRedo size={14} />
              <span className="rb-small-label">{t('ribbonNextChange')}</span>
            </button>
          </div>
          <button
            className="rb-big"
            disabled={!hasDoc || revisionCount === 0}
            data-tip={`${t('ribbonAiRevisionsTip', { count: revisionCount })} — ${t('ribbonAiCreditNote')}`}
            onClick={() => onAiPreset(t('ribbonAiRevisionsPrompt'))}
          >
            <span className="rb-big-icon">
              <span className="ai-feature-icon" aria-hidden="true">
                <svg
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="1.5"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <path d="M4 5h16M4 9h12M4 13h9M4 17h7" />
                  <path
                    d="M17 14l.26.7c.34.91.5 1.37.84 1.7.33.33.79.5 1.7.84l.7.26-.7.26c-.91.34-1.37.5-1.7.84-.34.33-.5.79-.84 1.7L17 21l-.26-.7c-.34-.91-.5-1.37-.84-1.7-.33-.34-.79-.5-1.7-.84l-.7-.26.7-.26c.91-.34 1.37-.5 1.7-.84.34-.33.5-.79.84-1.7L17 14z"
                    strokeLinejoin="round"
                  />
                  <path d="M19.5 4.5l-7 7-2 .5.5-2 7-7z" />
                </svg>
              </span>
            </span>
            <span>{t('ribbonAiRevisions')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupChanges')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonCompareTip')}
              onClick={() => toggleDropdown(setDropdown, 'compare')}
            >
              <span className="rb-big-icon">
                <IconCompare size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonCompare')}</span>
            </button>
            {dropdown === 'compare' && (
              <div data-rb-panel="" className="layout-menu">
                <button onClick={() => pickCommand('compare-documents')}>
                  {t('ribbonCompareDocuments')}
                </button>
                <button onClick={() => pickCommand('tools-combine-documents')}>
                  {t('ribbonCombineDocuments')}
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonCompare')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className={`rb-big ${protectActive ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonProtectDocTip')}
            onClick={() => runCommand('protect-document')}
          >
            <span className="rb-big-icon">
              <IconLock size={BIG} />
            </span>
            <span>{t('ribbonProtectDoc')}</span>
          </button>
          <RbCommand cmd="file-always-read-only" look="big">
            <IconReadOnlyDoc size={BIG} />
          </RbCommand>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupProtect')}</div>
      </div>
    </>
  )
}

/* ================= View ================= */

/** document rendering modes: Print Layout / Web Layout / Outline */
export type ViewMode = 'print' | 'web' | 'outline' | 'draft' | 'story'

interface ViewTabProps {
  hasDoc: boolean
  zoom: number
  darkPage: boolean
  showRuler: boolean
  showNav: boolean
  viewMode: ViewMode
  readMode: boolean
  showGrid: boolean
  splitView: boolean
}

export function ViewTab({
  hasDoc,
  zoom,
  darkPage,
  showRuler,
  showNav,
  viewMode,
  readMode,
  showGrid,
  splitView,
}: ViewTabProps) {
  const { t } = useI18n()
  const runCommand = useRunCommand()
  const [winMenuOpen, setWinMenuOpen] = useState(false)
  const [windows, setWindows] = useState<DocsTabInfo[]>([])
  /** wrap holding both the switch-tabs trigger and its menu */
  const winMenuRef = useRef<HTMLDivElement>(null)

  const toggleWinMenu = async () => {
    if (!winMenuOpen) setWindows(await window.desktop.listDocsTabs())
    setWinMenuOpen((v) => !v)
  }

  // presses inside the wrap (trigger + menu) are handled by their own onClick;
  // anything else — including window blur / shell chrome presses — closes it
  useDismissablePopover(winMenuOpen, () => setWinMenuOpen(false), {
    inside: () => [winMenuRef.current],
  })

  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className={`rb-big ${viewMode === 'print' && !readMode ? 'active' : ''}`}
            disabled={!hasDoc}
            data-tip={t('ribbonPrintLayoutTip')}
            onClick={() => runCommand('view-mode', 'print')}
          >
            <span className="rb-big-icon">
              <IconPrintLayout size={BIG} />
            </span>
            <span>{t('ribbonPrintLayout')}</span>
          </button>
          <button
            className={`rb-big ${viewMode === 'web' ? 'active' : ''}`}
            disabled={!hasDoc}
            data-tip={t('ribbonWebLayoutTip')}
            onClick={() => runCommand('view-mode', 'web')}
          >
            <span className="rb-big-icon">
              <IconWebLayout size={BIG} />
            </span>
            <span>{t('ribbonWebLayout')}</span>
          </button>
          <div className="rb-col rb-pair">
            <button
              className={`rb-small ${viewMode === 'outline' ? 'active' : ''}`}
              disabled={!hasDoc}
              data-tip={t('ribbonOutlineViewTip')}
              onClick={() => runCommand('view-mode', 'outline')}
            >
              <IconOutlineView size={14} />
              <span className="rb-small-label">{t('ribbonOutlineView')}</span>
            </button>
            <button
              className={`rb-small ${viewMode === 'draft' ? 'active' : ''}`}
              disabled={!hasDoc}
              data-tip={t('ribbonDraft')}
              onClick={() => runCommand('view-mode', 'draft')}
            >
              <IconDraftView size={14} />
              <span className="rb-small-label">{t('ribbonDraft')}</span>
            </button>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupViews')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <RbCommand cmd="view-focus" look="big">
            <IconFocusView size={BIG} />
          </RbCommand>
          <button
            className={`rb-big ${viewMode === 'story' ? 'active' : ''}`}
            disabled={!hasDoc}
            data-tip={t('ribbonStoryViewTip')}
            onClick={() => runCommand('view-mode', 'story')}
          >
            <span className="rb-big-icon">
              <IconStoryView size={BIG} />
            </span>
            <span>{t('ribbonStoryView')}</span>
          </button>
          <button
            className={`rb-big ${readMode ? 'active' : ''}`}
            disabled={!hasDoc}
            data-tip={t('ribbonReadModeTip')}
            onClick={() => runCommand('read-mode')}
          >
            <span className="rb-big-icon">
              <IconReadMode size={BIG} />
            </span>
            <span>{t('ribbonReadMode')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc || viewMode !== 'print' || readMode}
            data-tip={t('ribbonPagePreviewTip')}
            onClick={() => runCommand('page-preview')}
          >
            <span className="rb-big-icon">
              <IconWholePage size={BIG} />
            </span>
            <span>{t('ribbonPagePreview')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupImmersive')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-col">
            <button
              className={`rb-small ${showRuler ? 'active' : ''}`}
              disabled={!hasDoc}
              data-tip={t('ribbonRulerTip')}
              onClick={() => runCommand('toggle-ruler')}
            >
              <IconRuler size={14} />
              <span className="rb-small-label">{t('ribbonRuler')}</span>
            </button>
            <button
              className={`rb-small ${showGrid ? 'active' : ''}`}
              disabled={!hasDoc}
              data-tip={t('ribbonGridlinesTip')}
              onClick={() => runCommand('toggle-grid')}
            >
              <IconGridlines size={14} />
              <span className="rb-small-label">{t('ribbonGridlines')}</span>
            </button>
            <button
              className={`rb-small ${showNav ? 'active' : ''}`}
              disabled={!hasDoc}
              data-tip={t('ribbonNavPaneTip')}
              onClick={() => runCommand('toggle-navigation')}
            >
              <IconNavPane size={14} />
              <span className="rb-small-label">{t('ribbonNavPane')}</span>
            </button>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupShow')}</div>
      </div>

      <div className="ribbon-sep" />
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonZoomDialogTip')}
            onClick={() => runCommand('zoom-dialog')}
          >
            <span className="rb-big-icon">
              <IconZoom size={BIG} />
            </span>
            <span>{t('ribbonZoomDialog')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonZoomOut')}
            onClick={() => runCommand('zoom-out')}
          >
            <span className="rb-big-icon">
              <IconZoomOut size={BIG} />
            </span>
            <span>{t('ribbonZoomOut')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonZoomIn')}
            onClick={() => runCommand('zoom-in')}
          >
            <span className="rb-big-icon">
              <IconZoomIn size={BIG} />
            </span>
            <span>{t('ribbonZoomIn')}</span>
          </button>
          <button
            className={`rb-big ${zoom === 100 ? 'active' : ''}`}
            disabled={!hasDoc}
            data-tip={t('ribbonZoom100Tip')}
            aria-label={t('ribbonZoom100Tip')}
            onClick={() => runCommand('zoom-100')}
          >
            <span className="rb-big-icon">
              <IconZoom100 size={BIG} />
            </span>
            <span>100%</span>
          </button>
          <div className="rb-col">
            <button
              className="rb-small"
              disabled={!hasDoc}
              data-tip={t('ribbonWholePageTip')}
              onClick={() => runCommand('zoom-whole-page')}
            >
              <IconWholePage size={14} />
              <span className="rb-small-label">{t('ribbonWholePage')}</span>
            </button>
            <button
              className="rb-small"
              disabled={!hasDoc}
              data-tip={t('ribbonMultiPageTip')}
              onClick={() => runCommand('zoom-multi-page')}
            >
              <IconMultiPage size={14} />
              <span className="rb-small-label">{t('ribbonMultiPage')}</span>
            </button>
            <button
              className="rb-small"
              disabled={!hasDoc}
              data-tip={t('ribbonPageWidthTip')}
              onClick={() => runCommand('zoom-page-width')}
            >
              <IconPageWidth size={14} />
              <span className="rb-small-label">{t('ribbonPageWidth')}</span>
            </button>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupZoom')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            data-tip={t('ribbonNewTabTip')}
            onClick={() => runCommand('new-window')}
          >
            <span className="rb-big-icon">
              <IconNewWindow size={BIG} />
            </span>
            <span>{t('ribbonNewTab')}</span>
          </button>
          <RbCommand cmd="view-arrange-all" look="big">
            <IconArrangeAll size={BIG} />
          </RbCommand>
          <RbCommand cmd="view-merge-windows" look="big">
            <IconMergeWindows size={BIG} />
          </RbCommand>
          <button
            className={`rb-big ${splitView ? 'active' : ''}`}
            disabled={!hasDoc}
            data-tip={t('ribbonSplitTip')}
            onClick={() => runCommand('split-view')}
          >
            <span className="rb-big-icon">
              <IconSplit size={BIG} />
            </span>
            <span>{t('ribbonSplit')}</span>
          </button>
          <div className="rb-split-wrap" ref={winMenuRef}>
            <button
              className="rb-big"
              data-tip={t('ribbonSwitchTabsTip')}
              onClick={() => void toggleWinMenu()}
            >
              <span className="rb-big-icon">
                <IconSwitchWindows size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonSwitchTabs')}</span>
            </button>
            {winMenuOpen && (
              <div className="layout-menu align-right">
                {windows.map((w) => (
                  <button
                    key={w.id}
                    onClick={() => {
                      void window.desktop.focusDocsTab(w.id)
                      setWinMenuOpen(false)
                    }}
                  >
                    {w.focused ? '✓ ' : ''}
                    {w.title || 'GenOffice Docs'}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupWindow')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className={`rb-big ${darkPage ? 'active' : ''}`}
            data-tip={t('ribbonDarkModeTip')}
            onClick={() => runCommand('toggle-dark')}
          >
            <span className="rb-big-icon">
              <IconMoon size={BIG} />
            </span>
            <span>{t('ribbonDarkMode')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupAppearance')}</div>
      </div>
    </>
  )
}
/* ================= Draw ================= */

interface DrawTabProps {
  hasDoc: boolean
  dropdown: string | null
  setDropdown: SetDropdown
  tool: InkTool
  onTool: (tool: InkTool) => void
  presets: InkPreset[]
  activeId: string
  onPresets: (presets: InkPreset[]) => void
  onActivate: (id: string) => void
  annotationCount: number
  onClearAll: () => void
}

function PenGlyph({ kind, size }: { kind: InkPenKind; size: number }) {
  if (kind === 'highlighter') return <IconHighlighterPen size={size} />
  if (kind === 'pencil') return <IconPencilInk size={size} />
  return <IconPen size={size} />
}

export function DrawTab({
  hasDoc,
  dropdown,
  setDropdown,
  tool,
  onTool,
  presets,
  activeId,
  onPresets,
  onActivate,
  annotationCount,
  onClearAll,
}: DrawTabProps) {
  const { t } = useI18n()
  const active = presets.find((p) => p.id === activeId) ?? presets[0]
  const drawing = isPenKind(tool)
  const kindLabel = (kind: InkPenKind) =>
    kind === 'highlighter'
      ? t('ribbonHighlighter')
      : kind === 'pencil'
        ? t('ribbonPencil')
        : t('ribbonPen')
  const update = (id: string, patch: Partial<InkPreset>) =>
    onPresets(presets.map((p) => (p.id === id ? { ...p, ...patch } : p)))
  const remove = (id: string) => {
    const rest = presets.filter((p) => p.id !== id)
    onPresets(rest)
    if (id === activeId) {
      onActivate(rest[0].id)
      if (drawing) onTool(rest[0].kind)
    }
    setDropdown(() => null)
  }
  const add = (kind: InkPenKind) => {
    const preset = newInkPreset(kind)
    onPresets([...presets, preset])
    onActivate(preset.id)
    onTool(kind)
    setDropdown(() => null)
  }
  // a slot click arms that pen; a second click on the armed pen opens its properties
  const pickSlot = (preset: InkPreset) => {
    if (drawing && preset.id === activeId) {
      toggleDropdown(setDropdown, `inkPen:${preset.id}`)
      return
    }
    onActivate(preset.id)
    onTool(preset.kind)
  }

  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className={`rb-big ${drawing ? 'active' : ''}`}
            disabled={!hasDoc}
            data-tip={t('ribbonInkDrawTip')}
            onClick={() => onTool(drawing ? 'select' : active.kind)}
          >
            <span
              className="rb-big-icon"
              style={drawing ? { color: `#${active.color}` } : undefined}
            >
              <PenGlyph kind={active.kind} size={BIG} />
            </span>
            <span>{t('ribbonInkDraw')}</span>
          </button>
          <button
            className={`rb-big ${tool === 'eraser' ? 'active' : ''}`}
            disabled={!hasDoc}
            data-tip={t('ribbonEraserTip')}
            onClick={() => onTool(tool === 'eraser' ? 'select' : 'eraser')}
          >
            <span className="rb-big-icon">
              <IconEraser size={BIG} />
            </span>
            <span>{t('ribbonEraser')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupDrawingTools')}</div>
      </div>
      <div className="ribbon-sep" />
      <div className="ribbon-group">
        <div className="ribbon-group-items ink-pens">
          {presets.map((preset) => {
            const armed = drawing && preset.id === activeId
            const open = dropdown === `inkPen:${preset.id}`
            const widths = widthsFor(preset.kind)
            const colors = colorsFor(preset.kind)
            return (
              <div key={preset.id} className="rb-split-wrap">
                <button
                  className={`ink-pen-slot ${armed ? 'active' : ''}`}
                  disabled={!hasDoc}
                  data-tip={`${kindLabel(preset.kind)} · #${preset.color}`}
                  aria-label={kindLabel(preset.kind)}
                  onClick={() => pickSlot(preset)}
                >
                  <span className="ink-pen-glyph" style={{ color: `#${preset.color}` }}>
                    <PenGlyph kind={preset.kind} size={28} />
                  </span>
                  <span
                    className="ink-pen-bar"
                    style={{
                      background: `#${preset.color}`,
                      height: Math.max(
                        2,
                        Math.min(8, preset.width / (preset.kind === 'highlighter' ? 4 : 1)),
                      ),
                      opacity: preset.kind === 'highlighter' ? HIGHLIGHTER_ALPHA : 1,
                    }}
                  />
                  {armed && <IconCaret />}
                </button>
                {open && (
                  <div data-rb-panel="" className="layout-menu ink-pen-pop">
                    <div className="ink-pen-preview">
                      <span
                        style={{
                          background: `#${preset.color}`,
                          height: preset.width,
                          opacity: preset.kind === 'highlighter' ? HIGHLIGHTER_ALPHA : 1,
                        }}
                      />
                    </div>
                    <div className="ink-pen-section">{t('ribbonInkThickness')}</div>
                    <div className="ink-widths">
                      {widths.map((w, i) => (
                        <button
                          key={w}
                          className={`ink-width ${preset.width === w ? 'active' : ''}`}
                          data-tip={t('ribbonPixels', { w })}
                          aria-label={t('ribbonPixels', { w })}
                          onClick={() => update(preset.id, { width: w })}
                        >
                          <span
                            className="ink-width-dot"
                            style={{
                              width: 4 + i * 3,
                              height: 4 + i * 3,
                              background: `#${preset.color}`,
                            }}
                          />
                        </button>
                      ))}
                    </div>
                    <div className="ink-pen-section">{t('ribbonInkColors')}</div>
                    <div className="ink-pen-grid">
                      {colors.map((hex) => (
                        <button
                          key={hex}
                          className={`ink-swatch ${preset.color === hex ? 'active' : ''}`}
                          style={{ background: `#${hex}` }}
                          data-tip={`#${hex}`}
                          aria-label={`#${hex}`}
                          onClick={() => update(preset.id, { color: hex })}
                        />
                      ))}
                    </div>
                    <label className="ink-pen-item ink-pen-more">
                      <IconPalette size={14} />
                      {t('ribbonMoreColors')}
                      <input
                        type="color"
                        value={`#${preset.color}`}
                        onChange={(e) =>
                          update(preset.id, { color: e.target.value.slice(1).toUpperCase() })
                        }
                      />
                    </label>
                    <button
                      className="ink-pen-item ink-pen-delete"
                      disabled={presets.length <= 1}
                      onClick={() => remove(preset.id)}
                    >
                      <IconTrash size={14} />
                      {t('ribbonInkDelete')}
                    </button>
                  </div>
                )}
              </div>
            )
          })}
          <div className="rb-split-wrap">
            <button
              className="rb-big ink-pen-add"
              disabled={!hasDoc}
              data-tip={t('ribbonInkAddTip')}
              onClick={() => toggleDropdown(setDropdown, 'inkAdd')}
            >
              <span className="rb-big-icon">
                <IconPlus size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonInkAdd')}</span>
            </button>
            {dropdown === 'inkAdd' && (
              <div data-rb-panel="" className="layout-menu ink-pen-pop">
                {(['pen', 'highlighter', 'pencil'] as const).map((kind) => (
                  <button key={kind} className="ink-pen-item" onClick={() => add(kind)}>
                    <PenGlyph kind={kind} size={16} />
                    {kindLabel(kind)}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupPens')}</div>
      </div>
      <div className="ribbon-sep" />
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc || annotationCount === 0}
            data-tip={t('ribbonClearAllTip')}
            onClick={onClearAll}
          >
            <span className="rb-big-icon">
              <IconEraser size={BIG} />
            </span>
            <span>{t('ribbonClearAll')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupClear')}</div>
      </div>
    </>
  )
}
