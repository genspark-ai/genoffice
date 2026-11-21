import type { Editor } from '@tiptap/core'
import type { EditorState } from '@tiptap/pm/state'
import { fontFamilyMarkAttrs } from '../font-list'

export const DROP_CAP_DEFAULT_LINES = 3
export const DROP_CAP_MAX_LINES = 10

export function clampDropCapLines(lines: unknown): number {
  const n = Math.round(Number(lines))
  if (!Number.isFinite(n)) return DROP_CAP_DEFAULT_LINES
  return Math.min(DROP_CAP_MAX_LINES, Math.max(1, n))
}

/** the lines a paragraph's dropCap attribute (JSON of type and lines) asks the letter to span */
export function dropCapLinesOf(attr: unknown): number {
  if (typeof attr !== 'string') return DROP_CAP_DEFAULT_LINES
  try {
    const lines = (JSON.parse(attr) as { lines?: unknown }).lines
    return lines === undefined ? DROP_CAP_DEFAULT_LINES : clampDropCapLines(lines)
  } catch {
    return DROP_CAP_DEFAULT_LINES
  }
}

/** Word's limit on a drop cap's distance from the text: 22 inches */
export const DROP_CAP_MAX_GAP = 31680

/** the distance from the text (w:hSpace, twips) a paragraph's dropCap attribute asks for */
export function dropCapGapOf(attr: unknown): number {
  if (typeof attr !== 'string') return 0
  try {
    const gap = Math.round(Number((JSON.parse(attr) as { hSpace?: unknown }).hSpace))
    return gap > 0 ? Math.min(DROP_CAP_MAX_GAP, gap) : 0
  } catch {
    return 0
  }
}

/** the first letter of each paragraph the selection touches: where a drop cap's font lives */
function firstLetters(state: EditorState): Array<{ from: number; to: number }> {
  const out: Array<{ from: number; to: number }> = []
  const { from, to } = state.selection
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isTextblock) return true
    const first = node.firstChild
    if (first?.isText && first.text) {
      const letter = String.fromCodePoint(first.text.codePointAt(0)!)
      out.push({ from: pos + 1, to: pos + 1 + letter.length })
    }
    return false
  })
  return out
}

const EAST_ASIAN_LETTER = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u

/**
 * The font the first letter of the caret's paragraph says it has ('' when it names none),
 * read from the slot that letter is drawn from.
 */
export function dropCapFontOf(state: EditorState): { font: string; eastAsian: boolean } {
  const at = firstLetters(state)[0]
  const node = at && state.doc.nodeAt(at.from)
  const attrs = node?.marks.find((m) => m.type.name === 'docTextStyle')?.attrs
  const eastAsian = EAST_ASIAN_LETTER.test(node?.text?.slice(0, at.to - at.from) ?? '')
  const own = eastAsian ? (attrs?.eastAsiaFont ?? attrs?.font) : (attrs?.fontAscii ?? attrs?.font)
  return { font: String(own ?? ''), eastAsian }
}

/** Word keeps a drop cap's font on the letter itself: the first letter takes `font` */
export function setDropCapFont(editor: Editor, font: string): boolean {
  const { state } = editor
  const type = state.schema.marks.docTextStyle
  const tr = state.tr
  for (const { from, to } of firstLetters(state)) {
    const mark = type.isInSet(state.doc.nodeAt(from)?.marks ?? [])
    tr.addMark(from, to, type.create({ ...mark?.attrs, ...fontFamilyMarkAttrs(font) }))
  }
  if (!tr.docChanged) return false
  editor.view.dispatch(tr)
  return true
}
