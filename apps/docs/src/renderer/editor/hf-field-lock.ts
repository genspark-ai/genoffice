/**
 * Lock Fields, Unlock Fields, Unlink Fields and the field stops of a header or
 * footer: its PAGE / NUMPAGES / SECTIONPAGES fields are live atoms with no
 * result of their own, so locking gives one its number and unlocking takes it back.
 */
import type { Editor } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'
import type { Node as PmDocNode } from '@tiptap/pm/model'
import { PluginKey, type EditorState } from '@tiptap/pm/state'
import { collectFields } from './field-codes'

export type HfFieldKind = 'PAGE' | 'NUMPAGES' | 'SECTIONPAGES'

export interface HfFieldNumbers {
  page: string
  total: string
  /** page count of the edited page's section (SECTIONPAGES) */
  section: string
}

export const hfFieldNumbersKey = new PluginKey<HfFieldNumbers>('hfFieldNumbers')

const HF_FIELD_KINDS: readonly HfFieldKind[] = ['PAGE', 'NUMPAGES', 'SECTIONPAGES']

/** the live header/footer field a bare instruction names; one with switches keeps its cached form */
export function liveHfFieldOf(instr: string): HfFieldKind | null {
  const bare = instr.trim().toUpperCase()
  return HF_FIELD_KINDS.find((kind) => kind === bare) ?? null
}

/** the live page fields the selection touches, a caret at either edge of one included */
function touchedHfFields(state: EditorState): Array<{ pos: number; node: PmDocNode }> {
  const { from, to } = state.selection
  const out: Array<{ pos: number; node: PmDocNode }> = []
  state.doc.descendants((node, pos) => {
    if (node.type.name === 'hfField' && pos <= to && pos + node.nodeSize >= from)
      out.push({ pos, node })
  })
  return out
}

const shownNumber = (state: EditorState, kind: unknown): string => {
  const nums = hfFieldNumbersKey.getState(state)
  return (
    (kind === 'NUMPAGES' ? nums?.total : kind === 'SECTIONPAGES' ? nums?.section : nums?.page) ||
    '1'
  )
}

/** where the live page fields of a header or footer stand, for Next Field / Previous Field */
export function hfFieldStops(doc: PmDocNode): number[] {
  const stops: number[] = []
  doc.descendants((node, pos) => {
    if (node.type.name === 'hfField') stops.push(pos)
  })
  return stops
}

/**
 * Lock Fields in a header or footer: a live page field the selection touches
 * becomes a locked field holding the number it shows, so every page shows that
 * number. Unlock Fields makes a locked page field live again. Returns how many changed.
 */
export function setHfPageFieldsLocked(editor: Editor, locked: boolean): number {
  const { state } = editor
  const { hfField } = state.schema.nodes
  const field = state.schema.marks.instrField
  if (!hfField || !field || !editor.isEditable) return 0
  const tr = closeHistory(state.tr)
  let done = 0
  if (locked) {
    for (const { pos, node } of touchedHfFields(state).reverse()) {
      const mark = field.create({ instr: String(node.attrs.kind), locked: true })
      tr.replaceWith(
        pos,
        pos + node.nodeSize,
        state.schema.text(shownNumber(state, node.attrs.kind), [...node.marks, mark]),
      )
      done++
    }
  } else {
    const { from, to } = state.selection
    for (const range of collectFields(state.doc, field).reverse()) {
      const kind = liveHfFieldOf(range.instr)
      if (!kind || range.mark.attrs.locked !== true || range.from > to || range.to < from) continue
      const marks = (state.doc.nodeAt(range.from)?.marks ?? []).filter((m) => m.type !== field)
      tr.replaceWith(range.from, range.to, hfField.create({ kind }, null, marks))
      done++
    }
  }
  if (done > 0) editor.view.dispatch(tr)
  return done
}

/** Unlink Fields on the live page fields the selection touches: each becomes the number it shows */
export function unlinkHfPageFields(editor: Editor): number {
  const { state } = editor
  if (!state.schema.nodes.hfField || !editor.isEditable) return 0
  const hits = touchedHfFields(state)
  if (hits.length === 0) return 0
  const tr = closeHistory(state.tr)
  for (const { pos, node } of hits.reverse())
    tr.replaceWith(
      pos,
      pos + node.nodeSize,
      state.schema.text(shownNumber(state, node.attrs.kind), node.marks),
    )
  editor.view.dispatch(tr)
  return hits.length
}
