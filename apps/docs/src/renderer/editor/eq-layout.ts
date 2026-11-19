import type { Editor } from '@tiptap/core'
import type { Mark, Schema } from '@tiptap/pm/model'
import { NodeSelection, type EditorState } from '@tiptap/pm/state'
import type { EncloseStyle } from '@genoffice/docx-engine'
import { fontFamilyMarkAttrs } from '../font-list'

export type EqLayoutKind = 'enclose' | 'combine'

/** Word's limits: one full-width or two half-width characters in an enclosure, six combined */
export const ENCLOSE_MAX = 2
export const COMBINE_MAX = 6

/** the plain characters of enclosed or combined characters: what a search, a count or a copy sees */
export function eqLayoutText(attrs: { kind?: unknown; a?: unknown; b?: unknown }): string {
  return attrs.kind === 'combine'
    ? `${String(attrs.a ?? '')}${String(attrs.b ?? '')}`
    : String(attrs.b ?? '')
}

const MARK_CLASSES: Record<string, string> = {
  '\u25cb': 'doc-enclose-circle',
  '\u25a1': 'doc-enclose-square',
  '\u25b3': 'doc-enclose-triangle',
  '\u25c7': 'doc-enclose-diamond',
}

/**
 * The class that draws an enclosure. The shapes are drawn, not set as glyphs:
 * a Latin font's circle or square is far smaller than the full-width one
 * Word encloses with.
 */
export const encloseMarkClass = (enclosure: string): string => MARK_CLASSES[enclosure] ?? ''

/** Combined characters fill two lines: the upper takes the first half, rounded up. */
export function combineLines(text: string): { a: string; b: string } {
  const chars = Array.from(text)
  const upper = Math.ceil(chars.length / 2)
  return { a: chars.slice(0, upper).join(''), b: chars.slice(upper).join('') }
}

/** what the dialog works on: selected text, the selected enclosure or combination, or the caret */
export interface EqLayoutTarget {
  from: number
  to: number
  text: string
  /** the node already there, when one is selected */
  existing: { kind: EqLayoutKind; a: string; b: string; style: EncloseStyle | null } | null
}

/**
 * The target for Enclose Characters or Combine Characters. Null when the
 * selection is something else: several paragraphs, another kind of object, or
 * more text than the layout holds.
 */
export function eqLayoutTargetOf(state: EditorState, kind: EqLayoutKind): EqLayoutTarget | null {
  const { selection } = state
  if (selection instanceof NodeSelection) {
    const { node } = selection
    if (node.type.name !== 'docEqLayout' || node.attrs.kind !== kind) return null
    return {
      from: selection.from,
      to: selection.to,
      text: eqLayoutText(node.attrs),
      existing: {
        kind,
        a: String(node.attrs.a ?? ''),
        b: String(node.attrs.b ?? ''),
        style: (node.attrs.style as EncloseStyle | null) ?? null,
      },
    }
  }
  const { from, to, $from, $to } = selection
  if (!$from.sameParent($to) || !$from.parent.isTextblock) return null
  let plain = true
  state.doc.nodesBetween(from, to, (node) => {
    if (node.isInline && !node.isText) plain = false
  })
  const text = state.doc.textBetween(from, to)
  const max = kind === 'enclose' ? ENCLOSE_MAX : COMBINE_MAX
  if (!plain || Array.from(text).length > max) return null
  return { from, to, text, existing: null }
}

/** the font and size (pt) the target's text names itself; empty and 0 when it inherits them */
export function eqLayoutFontOf(
  state: EditorState,
  target: EqLayoutTarget,
): { font: string; sizePt: number } {
  const marks = state.doc.nodeAt(target.from)?.marks ?? state.selection.$from.marks()
  const attrs = marks.find((m) => m.type.name === 'docTextStyle')?.attrs
  return {
    font: String(attrs?.font ?? attrs?.fontAscii ?? ''),
    sizePt: Number(attrs?.sizeHalfPoints ?? 0) / 2,
  }
}

/** a font or size picked in the dialog; an absent one stays what the text has */
export interface EqLayoutFont {
  font?: string
  sizePt?: number
}

function marksWithFont(
  schema: Schema,
  marks: readonly Mark[],
  picked: EqLayoutFont,
): readonly Mark[] {
  const patch = {
    ...(picked.font ? fontFamilyMarkAttrs(picked.font) : {}),
    ...(picked.sizePt && picked.sizePt > 0
      ? { sizeHalfPoints: Math.round(picked.sizePt * 2) }
      : {}),
  }
  if (Object.keys(patch).length === 0) return marks
  const type = schema.marks.docTextStyle
  const own = type.isInSet(marks)
  const rest = own ? own.removeFromSet(marks) : marks
  return type.create({ ...own?.attrs, ...patch }).addToSet(rest)
}

/**
 * Puts enclosed or combined characters on the target, in the formatting of
 * the text they replace unless a font or size is picked. `null` takes the
 * layout off and leaves the plain text.
 */
export function applyEqLayout(
  editor: Editor,
  target: EqLayoutTarget,
  layout:
    | { kind: 'enclose'; text: string; enclosure: string; style: Exclude<EncloseStyle, 'none'> }
    | { kind: 'combine'; text: string }
    | null,
  picked: EqLayoutFont = {},
): boolean {
  const { state } = editor
  const { schema } = state
  const at = state.doc.nodeAt(target.from)
  const marks = marksWithFont(
    schema,
    at?.marks ?? state.storedMarks ?? state.selection.$from.marks(),
    layout ? picked : {},
  )
  const text = layout?.text.trim() ?? ''
  if (!layout || !text) {
    if (!target.existing) return false
    return editor
      .chain()
      .focus()
      .command(({ tr }) => {
        tr.replaceWith(target.from, target.to, schema.text(target.text, marks))
        return true
      })
      .run()
  }
  const attrs =
    layout.kind === 'enclose'
      ? { kind: 'enclose', a: layout.enclosure, b: text, style: layout.style, xml: '' }
      : { kind: 'combine', ...combineLines(text), style: null, xml: '' }
  const node = schema.nodes.docEqLayout.create(attrs, null, marks)
  return editor
    .chain()
    .focus()
    .command(({ tr }) => {
      tr.replaceWith(target.from, target.to, node)
      // it stays selected, so the dialog can be opened on it again
      tr.setSelection(NodeSelection.create(tr.doc, target.from))
      return true
    })
    .run()
}
