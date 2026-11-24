/**
 * AutoFormat As You Type, the two rules that hand a paragraph a style on
 * Enter: "Date style" (a line that is a date takes the built-in Date style)
 * and "Define styles based on your formatting" (a short line formatted by
 * hand as a heading takes a built-in heading style that looks the same, the
 * style defined from the formatting when none does yet).
 */
import { getSplittedAttributes, type Editor } from '@tiptap/core'
import type { Mark, Node as PmNode, ResolvedPos } from '@tiptap/pm/model'
import { TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state'
import { closeHistory } from '@tiptap/pm/history'
import {
  builtInHeadingStyleUpsert,
  type DocDefaults,
  type StyleInfo,
  type StyleUpsert,
} from '@genoffice/docx-engine'
import { defaultParagraphStyleId, headingStyleId } from '../style-gallery'

type Styles = Map<string, StyleInfo>

const MONTH =
  '(?:January|February|March|April|May|June|July|August|September|October|November|December)'
const DATE_FORMS = [
  // era and Gregorian year, month, day in kanji
  /^(?:\u660e\u6cbb|\u5927\u6b63|\u662d\u548c|\u5e73\u6210|\u4ee4\u548c)(?:\u5143|\d{1,2})\u5e74\d{1,2}\u6708\d{1,2}\u65e5$/,
  /^\d{4}\u5e74\d{1,2}\u6708\d{1,2}\u65e5$/,
  /^\d{4}\ub144 ?\d{1,2}\uc6d4 ?\d{1,2}\uc77c$/,
  new RegExp(`^(?:(?:Mon|Tues|Wednes|Thurs|Fri|Satur|Sun)day, )?${MONTH} \\d{1,2}, \\d{4}$`),
  new RegExp(`^\\d{1,2} ${MONTH} \\d{4}$`),
  /^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$/,
  /^\d{1,2}\/\d{1,2}\/(?:\d{2}|\d{4})$/,
]

/** the line is a date and nothing else */
export function isDateLine(text: string): boolean {
  const plain = text
    .trim()
    .replace(/[\uff10-\uff19]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0xfee0))
  return plain.length > 0 && DATE_FORMS.some((form) => form.test(plain))
}

/** the caret ends a body paragraph of the default style that holds text only */
function typedLine(state: EditorState, styles: Styles | undefined): ResolvedPos | null {
  const { $from, empty } = state.selection
  const para = $from.parent
  if (!empty || $from.depth !== 1 || para.type.name !== 'docParagraph') return null
  if (para.content.size === 0 || $from.parentOffset !== para.content.size) return null
  const styleId = para.attrs.styleId as string | null
  if (styleId && styleId !== defaultParagraphStyleId(styles)) return null
  let textOnly = true
  para.forEach((child) => {
    if (!child.isText) textOnly = false
  })
  return textOnly ? $from : null
}

/** Enter after the line: the paragraph below is the one Enter would have made, in the style the line had */
function splitBelow(editor: Editor, tr: Transaction, $from: ResolvedPos): Transaction {
  const para = $from.parent
  const attrs = getSplittedAttributes(
    editor.extensionManager.attributes,
    para.type.name,
    para.attrs,
  )
  tr.split($from.pos, 1, [{ type: para.type, attrs }])
  return tr.setSelection(TextSelection.create(tr.doc, $from.pos + 2)).setStoredMarks([])
}

/** the document's Date style, defined as Word's built-in one when it has none; null when the id is taken */
function dateStyleId(
  styles: Styles | undefined,
  defineStyle: (up: StyleUpsert) => void,
): string | null {
  for (const info of styles?.values() ?? [])
    if (info.type === 'paragraph' && /^date$/i.test(info.name)) return info.styleId
  if (styles?.has('Date')) return null
  const base = defaultParagraphStyleId(styles)
  defineStyle({
    styleId: 'Date',
    type: 'paragraph',
    name: 'Date',
    ...(base ? { basedOn: base, next: base } : {}),
    quickFormat: false,
    builtIn: true,
  })
  return 'Date'
}

/** Date style: Enter after a line that is a date gives the line the Date style */
export function dateStyleOnEnter(
  editor: Editor,
  styles: Styles | undefined,
  defineStyle: ((up: StyleUpsert) => void) | undefined,
): boolean {
  const { state } = editor
  const $from = typedLine(state, styles)
  if (!$from || !defineStyle || !isDateLine($from.parent.textContent)) return false
  const styleId = dateStyleId(styles, defineStyle)
  if (!styleId) return false
  editor.view.dispatch(
    closeHistory(state.tr).setNodeMarkup($from.before(), null, { ...$from.parent.attrs, styleId }),
  )
  editor.view.dispatch(closeHistory(splitBelow(editor, editor.state.tr, $from)))
  return true
}

/** what was applied by hand to the whole of a line */
export interface ManualFormat {
  bold: boolean
  italic: boolean
  underline: boolean
  sizeHalfPoints: number | null
  font: string | null
  color: string | null
}

const TEXT_STYLE_KEPT = new Set([
  'color',
  'sizeHalfPoints',
  'font',
  'fontAscii',
  'eastAsiaFont',
  'eaSlotEmpty',
  'csFont',
])

// what Track Changes puts on text typed while it records: not formatting of the line
const RECORD_MARKS = new Set(['ins', 'rprChange'])

const formatMarksOf = (node: PmNode): Mark[] =>
  node.marks.filter((m) => !RECORD_MARKS.has(m.type.name))

/** the one formatting the whole line carries, when it is formatting a style can hold */
export function manualFormatOf(para: PmNode): ManualFormat | null {
  const first = para.firstChild
  if (!first) return null
  const marks = formatMarksOf(first)
  let uniform = true
  para.forEach((child) => {
    const own = formatMarksOf(child)
    if (!child.isText || !marks.every((m) => m.isInSet(own))) uniform = false
    if (own.length !== marks.length) uniform = false
  })
  if (!uniform) return null
  const format: ManualFormat = {
    bold: false,
    italic: false,
    underline: false,
    sizeHalfPoints: null,
    font: null,
    color: null,
  }
  for (const mark of marks) {
    const name = mark.type.name
    if (name === 'bold') format.bold = true
    else if (name === 'italic') format.italic = true
    else if (name === 'underline') {
      const { line, color } = mark.attrs as { line?: string | null; color?: string | null }
      if ((line && line !== 'single') || (color && color !== 'auto')) return null
      format.underline = true
    } else if (name === 'docTextStyle') {
      for (const [key, value] of Object.entries(mark.attrs))
        if (value !== null && value !== undefined && !TEXT_STYLE_KEPT.has(key)) return null
      format.sizeHalfPoints = (mark.attrs.sizeHalfPoints as number | null) ?? null
      format.font = (mark.attrs.font as string | null) ?? null
      format.color = (mark.attrs.color as string | null) ?? null
    } else return null
  }
  return format
}

const definedHere = new WeakMap<
  Editor,
  { styles: Styles | undefined; made: Map<string, { level: number; styleId: string }> }
>()

const HEADING_LINE = /^\S.{0,79}$/u
const SENTENCE_END = /[.,;:!?\u3002\uff0c\uff1b\uff1a\uff01\uff1f]\s*$/

function looksTheSame(info: StyleInfo, format: ManualFormat, baseSize: number): boolean {
  const shown = info.display
  if (!shown) return false
  return (
    !!shown.bold === format.bold &&
    !!shown.italic === format.italic &&
    !!shown.underline === format.underline &&
    (shown.sizeHalfPoints ?? baseSize) === (format.sizeHalfPoints ?? baseSize) &&
    (format.font === null || shown.font === format.font) &&
    (format.color === null || shown.color?.toLowerCase() === format.color.toLowerCase())
  )
}

/**
 * Define styles based on your formatting: Enter after a short line that is
 * bold or larger than body text throughout. The line becomes a heading in the
 * heading style that already looks like it; when none does, the first heading
 * level the document does not use is defined from the formatting.
 */
export function definedStyleOnEnter(
  editor: Editor,
  styles: Styles | undefined,
  docDefaults: DocDefaults | undefined,
  defineStyle: ((up: StyleUpsert) => void) | undefined,
): boolean {
  const { state } = editor
  const $from = typedLine(state, styles)
  const heading = state.schema.nodes.docHeading
  if (!$from || !defineStyle || !heading) return false
  const para = $from.parent
  const text = para.textContent
  if (!HEADING_LINE.test(text) || SENTENCE_END.test(text)) return false
  const format = manualFormatOf(para)
  if (!format) return false
  const base = defaultParagraphStyleId(styles)
  const baseSize =
    (base ? styles?.get(base)?.display?.sizeHalfPoints : undefined) ??
    docDefaults?.sizeHalfPoints ??
    20
  if (!format.bold && !((format.sizeHalfPoints ?? 0) > baseSize)) return false

  const used = new Set<number>()
  state.doc.descendants((node) => {
    if (node.type === heading) used.add(Number(node.attrs.level) || 1)
    return !node.isTextblock
  })
  // the styles this rule defined since the style table was last made: once the table is made
  // again (it has caught up, or another document is open) it is the table that answers
  let memory = definedHere.get(editor)
  if (!memory || memory.styles !== styles)
    definedHere.set(editor, (memory = { styles, made: new Map() }))
  const { made } = memory
  const look = JSON.stringify([format, para.attrs.align ?? null])
  const madeLevels = new Set([...made.values()].map((m) => m.level))
  const before = made.get(look)
  let level = before?.level ?? 0
  let styleId = before?.styleId
  let defined = !!before
  for (let l = 1; l <= 9 && !level; l++) {
    const id = headingStyleId(styles, l)
    const info = id ? styles?.get(id) : undefined
    if (id && info && !madeLevels.has(l) && looksTheSame(info, format, baseSize)) {
      level = l
      styleId = id
    }
  }
  for (let l = 1; l <= 9 && !level; l++) {
    if (used.has(l) || madeLevels.has(l)) continue
    level = l
    styleId = headingStyleId(styles, l)
    const fresh = builtInHeadingStyleUpsert(l, base)
    defineStyle({
      ...(styleId ? { styleId } : fresh),
      rPr: {
        bold: format.bold,
        italic: format.italic,
        underline: format.underline,
        sizeHalfPoints: format.sizeHalfPoints,
        font: format.font,
        color: format.color,
      },
      pPr: {
        ...(styleId ? {} : fresh.pPr),
        align: (para.attrs.align as NonNullable<StyleUpsert['pPr']>['align']) ?? null,
      },
    })
    styleId ??= fresh.styleId
    defined = true
    made.set(look, { level, styleId })
  }
  if (!level || !styleId) return false

  const at = $from.before()
  const tr = closeHistory(state.tr)
  for (const mark of formatMarksOf(para.firstChild!))
    tr.removeMark($from.start(), $from.end(), mark.type)
  tr.setNodeMarkup(at, heading, {
    ...para.attrs,
    styleId,
    level,
    outlineOnly: null,
    ...(defined ? { align: null } : {}),
  })
  editor.view.dispatch(tr)
  // the paragraph below is body text again, as the heading's following style says
  const below = editor.state.tr
  const plain = state.schema.nodes.docParagraph!
  const attrs = getSplittedAttributes(editor.extensionManager.attributes, plain.name, {
    ...para.attrs,
    ...(defined ? { align: null } : {}),
  })
  const end = below.doc.resolve(at + 1).end()
  below.split(end, 1, [{ type: plain, attrs }])
  below.setSelection(TextSelection.create(below.doc, end + 2)).setStoredMarks([])
  editor.view.dispatch(closeHistory(below))
  return true
}
