import type { Editor } from '@tiptap/core'
import type { ResolvedPos } from '@tiptap/pm/model'
import { Selection, TextSelection } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { pasteFromClipboard } from './paste-actions'

export const CARET_TARGETS = [
  'line-start',
  'line-end',
  'para-up',
  'para-down',
  'doc-start',
  'doc-end',
  'screen-up',
  'screen-down',
  'window-start',
  'window-end',
] as const
export type CaretTarget = (typeof CARET_TARGETS)[number]

const EXTEND = '-extend'

/** `para-up` or `para-up-extend`, as a key binding names a move */
export function parseCaretMove(payload: string): { to: CaretTarget; extend: boolean } | null {
  const extend = payload.endsWith(EXTEND)
  const to = CARET_TARGETS.find((t) => t === (extend ? payload.slice(0, -EXTEND.length) : payload))
  return to ? { to, extend } : null
}

/** the editor that has the keyboard: the document, a text box or a header / footer strip */
export function focusedEditor(): Editor | null {
  const dom = document.activeElement?.closest('.ProseMirror') as { editor?: Editor } | null
  return dom?.editor && !dom.editor.isDestroyed ? dom.editor : null
}

/** Word's ParaUp: the start of this paragraph, or from there the start of the one before */
function paragraphUp($head: ResolvedPos): number {
  if (!$head.parent.isTextblock) return Selection.near($head, -1).head
  const start = $head.start()
  if ($head.pos > start) return start
  const before = Selection.findFrom($head.doc.resolve($head.before()), -1, true)
  return before ? before.$head.start() : start
}

/** Word's ParaDown: the start of the next paragraph, or the end of the last one */
function paragraphDown($head: ResolvedPos): number {
  if (!$head.parent.isTextblock) return Selection.near($head, 1).head
  const after = Selection.findFrom($head.doc.resolve($head.after()), 1, true)
  return after ? after.head : $head.end()
}

/** a point of the window as a document position; a point over a ruler or between pages has none, so step toward the text */
function posInWindow(view: EditorView, left: number, top: number, step: 1 | -1): number | null {
  const frame = view.dom.closest('.editor-scroll')?.getBoundingClientRect()
  if (!frame) return null
  for (let y = top; y > frame.top && y < frame.bottom; y += step * 8) {
    const hit = view.posAtCoords({ left, top: y })
    if (hit) return hit.pos
  }
  return null
}

/**
 * Word's PageUp / PageDown and StartOfWindow / EndOfWindow. They are moves in
 * the window, not in the document: the pane scrolls first and the caret takes
 * the position under the point it should land on.
 */
function windowTarget(view: EditorView, to: CaretTarget): number | null {
  const pane = view.dom.closest<HTMLElement>('.editor-scroll')
  if (!pane) return null
  const frame = pane.getBoundingClientRect()
  const page = view.dom.getBoundingClientRect()
  const inset = 6
  if (to === 'window-start' || to === 'window-end') {
    const end = to === 'window-end'
    const left = end
      ? Math.min(page.right, frame.right) - inset
      : Math.max(page.left, frame.left) + inset
    return posInWindow(view, left, end ? frame.bottom - inset : frame.top + inset, end ? -1 : 1)
  }
  const dir = to === 'screen-down' ? 1 : -1
  const caret = view.coordsAtPos(view.state.selection.head)
  const before = pane.scrollTop
  pane.scrollTop = before + dir * pane.clientHeight
  const scrolled = pane.scrollTop - before
  if (scrolled === 0)
    return (dir > 0 ? Selection.atEnd(view.state.doc) : Selection.atStart(view.state.doc)).head
  // the caret keeps its place in the window; what the pane could not scroll, the caret travels
  const wanted = (caret.top + caret.bottom) / 2 + dir * pane.clientHeight - scrolled
  const top = Math.max(frame.top + inset, Math.min(frame.bottom - inset, wanted))
  return posInWindow(view, caret.left, top, dir > 0 ? -1 : 1)
}

export function moveCaret(editor: Editor, to: CaretTarget, extend: boolean): void {
  const { state, view } = editor
  if (to === 'screen-up' || to === 'screen-down' || to === 'window-start' || to === 'window-end') {
    const head = windowTarget(view, to)
    if (head === null) return
    const $to = state.doc.resolve(head)
    const selection = extend
      ? TextSelection.between(state.doc.resolve(state.selection.anchor), $to)
      : TextSelection.near($to, to === 'window-end' ? -1 : 1)
    view.dispatch(state.tr.setSelection(selection))
    return
  }
  if (to === 'line-start' || to === 'line-end') {
    // a line is a matter of layout: the browser's own move, which the editor then reads back
    view.dom.ownerDocument
      .getSelection()
      ?.modify(
        extend ? 'extend' : 'move',
        to === 'line-start' ? 'backward' : 'forward',
        'lineboundary',
      )
    // read back at once, as after a native caret key: the next key must find the new selection
    ;(view as unknown as { domObserver?: { flush(): void } }).domObserver?.flush()
    return
  }
  const { $head, anchor } = state.selection
  const head =
    to === 'para-up'
      ? paragraphUp($head)
      : to === 'para-down'
        ? paragraphDown($head)
        : (to === 'doc-start' ? Selection.atStart(state.doc) : Selection.atEnd(state.doc)).head
  const $to = state.doc.resolve(head)
  const selection = extend
    ? TextSelection.between(state.doc.resolve(anchor), $to)
    : TextSelection.near($to, to === 'para-up' || to === 'doc-start' ? 1 : -1)
  view.dispatch(state.tr.setSelection(selection).scrollIntoView())
}

/**
 * Deletes from the caret to the edge of the word, as ⌥⌫ / ⌥⌦ do. The range
 * is selected and then deleted by the Backspace handling, so a tracked
 * deletion, a list marker or a joined paragraph come out as they do when typed.
 */
export function deleteWord(editor: Editor, dir: 1 | -1): void {
  const { state, view } = editor
  const { $head, empty } = state.selection
  const edge = dir < 0 ? $head.start() : $head.end()
  const domSelection = view.dom.ownerDocument.getSelection()
  if (empty && $head.parent.isTextblock && $head.pos !== edge && domSelection) {
    domSelection.modify('extend', dir < 0 ? 'backward' : 'forward', 'word')
    const node = domSelection.focusNode
    const found = node ? view.posAtDOM(node, domSelection.focusOffset) : $head.pos
    // never past the paragraph: its mark goes on the next press, as in Word
    const head = dir < 0 ? Math.max(found, edge) : Math.min(found, edge)
    if (head === $head.pos) return
    view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, $head.pos, head)))
  }
  editor.commands.keyboardShortcut(dir < 0 ? 'Backspace' : 'Delete')
}

export const TEXT_EDITS = ['cut', 'copy', 'paste', 'select-all'] as const
export type TextEdit = (typeof TEXT_EDITS)[number]

/** the clipboard keys and Select All, carried out by the editor that has the keyboard */
export function editText(editor: Editor, edit: TextEdit): void {
  if (edit === 'select-all') editor.commands.selectAll()
  // the page may not paste by command; the ribbon's Paste reads the clipboard itself
  else if (edit === 'copy') editor.view.dom.ownerDocument.execCommand(edit)
  else if (!editor.isEditable) return
  else if (edit === 'paste') void pasteFromClipboard(editor)
  else editor.view.dom.ownerDocument.execCommand(edit)
}
