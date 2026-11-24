/**
 * Word's Insert Field Characters (Cmd+F9): a pair of field braces whose code
 * is typed in place. Until the field is updated its text is the code, not a
 * result; a save writes it as a field Word computes on open.
 */
import type { Editor } from '@tiptap/core'
import type { Mark } from '@tiptap/pm/model'
import { TextSelection, type EditorState } from '@tiptap/pm/state'
import { collectFields, PAGE_FIELD_INSTR, typedFieldCode, type FieldRange } from './field-codes'
import { nextFieldId } from './source-citations'

/** the field still shows the code being typed */
export const isFieldCodeOpen = (mark: Mark): boolean => mark.attrs.codeEdit === true

export function insertFieldChars(editor: Editor): boolean {
  const { state, view } = editor
  const markType = state.schema.marks.instrField
  const { from, to, $from, $to } = state.selection
  if (!markType || !editor.isEditable || !$from.parent.isTextblock || !$from.sameParent($to))
    return false
  // no field inside a field: a nested code has nowhere to be kept
  if (markType.isInSet($from.marks()) || state.doc.rangeHasMark(from, to, markType)) return false
  const text = ` ${state.doc.textBetween(from, to, '')} `
  const mark = markType.create({ instr: '', codeEdit: true, fieldId: nextFieldId(state.doc) })
  const tr = state.tr.replaceWith(from, to, state.schema.text(text, [...$from.marks(), mark]))
  // between the braces, after a code taken from the selection
  tr.setSelection(TextSelection.create(tr.doc, from + text.length - 1))
  view.dispatch(tr.scrollIntoView())
  return true
}

export function openFieldCodes(state: EditorState): FieldRange[] {
  return collectFields(state.doc).filter((field) => isFieldCodeOpen(field.mark))
}

/**
 * Update Field on a typed code: the field takes its result. `valueOf` gives
 * what this editor can compute; a code it cannot is left for Word, marked
 * dirty. Braces with nothing in them go away. Returns how many were closed.
 */
export function closeFieldCodes(
  editor: Editor,
  valueOf: (instr: string) => string,
  only?: FieldRange,
): number {
  const { state, view } = editor
  const markType = state.schema.marks.instrField
  if (!markType || !editor.isEditable) return 0
  const open = only ? [only] : openFieldCodes(state)
  if (open.length === 0) return 0
  const tr = state.tr
  for (const field of [...open].reverse()) {
    const instr = typedFieldCode(state.doc.textBetween(field.from, field.to, ''))
    if (!instr) {
      tr.delete(field.from, field.to)
      continue
    }
    const value = valueOf(instr)
    const mark = markType.create({ instr, dirty: !value || PAGE_FIELD_INSTR.test(instr) })
    const kept = markType.removeFromSet(state.doc.nodeAt(field.from)?.marks ?? [])
    tr.replaceWith(field.from, field.to, state.schema.text(value || ' ', [...kept, mark]))
  }
  view.dispatch(tr)
  return open.length
}
