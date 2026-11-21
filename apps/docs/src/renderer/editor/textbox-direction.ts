import type { Editor } from '@tiptap/core'
import { TEXTBOX_VERTS, type TextboxDisplay, type TextboxVert } from '@genoffice/docx-engine'
import { selectedBoxIndex } from './selected-box'

/** the directions Shape Format > Text Direction offers, in Word's order; null is horizontal */
export const TEXTBOX_DIRECTIONS: readonly (TextboxVert | null)[] = [
  null,
  'eaVert',
  'vert',
  'vert270',
  'wordArtVert',
]

export const HORIZONTAL = 'horz'

/** a command payload as a direction; undefined when it names none */
export function textboxDirectionOf(payload: string | undefined): TextboxVert | null | undefined {
  if (payload === HORIZONTAL) return null
  return TEXTBOX_VERTS.find((vert) => vert === payload)
}

function selectedBox(editor: Editor): { boxes: TextboxDisplay[]; index: number } | null {
  const boxes = editor.getAttributes('docProtected')?.textboxes as TextboxDisplay[] | null
  if (!Array.isArray(boxes) || boxes.length === 0) return null
  const index = selectedBoxIndex(editor.state)
  const box = boxes[index]
  // a line or a picture has no text to turn
  if (!box || box.picture || box.readOnly || !box.paras) return null
  return { boxes, index }
}

/** the text direction of the selected box: null is horizontal, undefined no box that holds text */
export function selectedBoxDirection(editor: Editor): TextboxVert | null | undefined {
  const sel = selectedBox(editor)
  return sel ? (sel.boxes[sel.index].vert ?? null) : undefined
}

/** turns the text of the selected box; the box keeps its size, as in Word */
export function setSelectedBoxDirection(editor: Editor, vert: TextboxVert | null): boolean {
  // what is being typed into the box goes into its model first, or the new model would overwrite it
  window.dispatchEvent(new Event('ai-docs-commit-tables'))
  const sel = selectedBox(editor)
  if (!sel || !editor.isEditable) return false
  const { vert: _was, ...rest } = sel.boxes[sel.index]
  const box: TextboxDisplay = vert ? { ...rest, vert } : rest
  return editor
    .chain()
    .focus()
    .updateAttributes('docProtected', {
      textboxes: sel.boxes.map((b, i) => (i === sel.index ? box : b)),
    })
    .run()
}
