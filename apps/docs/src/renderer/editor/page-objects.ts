import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { NodeSelection, type EditorState } from '@tiptap/pm/state'
import type { TextboxDisplay } from '@genoffice/docx-engine'
import { selectedObjectPositions } from './object-multi-select'
import { selectBoxAt, selectedBoxIndex } from './selected-box'

export type PageObjectKind = 'picture' | 'textBox' | 'shape'

/** a drawing object of the document, as the Selection Pane lists it */
export interface PageObject {
  pos: number
  /** which box of a node that packs several drawings; null for a picture node */
  index: number | null
  kind: PageObjectKind
  /** the object's ordinal among its kind, for a name like "Picture 2" */
  n: number
  /** the start of a text box's text */
  text: string
}

const SNIPPET_CHARS = 40

const boxText = (box: TextboxDisplay): string =>
  box.paras
    .map((p) => p.runs.map((r) => r.text).join(''))
    .join(' ')
    .trim()
    .slice(0, SNIPPET_CHARS)

/** the pictures, text boxes and shapes of the document, in document order */
export function pageObjects(doc: PmNode): PageObject[] {
  const out: PageObject[] = []
  const count: Record<PageObjectKind, number> = { picture: 0, textBox: 0, shape: 0 }
  const add = (pos: number, index: number | null, kind: PageObjectKind, text = '') =>
    out.push({ pos, index, kind, n: ++count[kind], text })
  doc.descendants((node, pos) => {
    // a horizontal rule is drawn by the same node, and is not a picture anywhere else
    if (node.type.name === 'docInlineImage' && !node.attrs.rule) add(pos, null, 'picture')
    if (node.type.name !== 'docProtected') return
    if (node.attrs.blockType === 'image') add(pos, null, 'picture')
    const boxes = node.attrs.textboxes as TextboxDisplay[] | null
    boxes?.forEach((box, index) => {
      const text = boxText(box)
      add(pos, index, box.picture ? 'picture' : text ? 'textBox' : 'shape', text)
    })
  })
  return out
}

/** whether `object` is what the selection is on */
export function isPageObjectSelected(state: EditorState, object: PageObject): boolean {
  const sel = state.selection
  if (!(sel instanceof NodeSelection)) return false
  if (sel.from !== object.pos) return selectedObjectPositions(state).includes(object.pos)
  return object.index === null || selectedBoxIndex(state) === object.index
}

/** Selects an object and brings it into view. False when the document changed under the list. */
export function selectPageObject(editor: Editor, object: PageObject): boolean {
  const { state } = editor
  const node = state.doc.nodeAt(object.pos)
  if (!node || (node.type.name !== 'docProtected' && node.type.name !== 'docInlineImage'))
    return false
  const tr =
    object.index === null
      ? state.tr.setSelection(NodeSelection.create(state.doc, object.pos))
      : selectBoxAt(state.tr, object.pos, object.index)
  editor.view.dispatch(tr.scrollIntoView())
  editor.view.focus()
  return true
}

/**
 * Word's Select Data: the chart whose data is to be edited. A chart shows its
 * data table while it is selected, so this selects one: the chart already
 * selected, else the first after the caret, else the first in the document.
 * False when the document has no chart.
 */
export function selectChartForData(editor: Editor): boolean {
  const { state } = editor
  const charts: number[] = []
  state.doc.descendants((node, pos) => {
    if (node.type.name === 'docProtected' && node.attrs.chartDisplay) charts.push(pos)
  })
  if (charts.length === 0) return false
  const { selection } = state
  const pos =
    selection instanceof NodeSelection && charts.includes(selection.from)
      ? selection.from
      : (charts.find((at) => at >= selection.from) ?? charts[0])
  editor.view.dispatch(state.tr.setSelection(NodeSelection.create(state.doc, pos)).scrollIntoView())
  editor.view.focus()
  return true
}
