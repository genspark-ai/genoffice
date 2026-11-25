import type { Editor } from '@tiptap/core'
import type { Slice } from '@tiptap/pm/model'
import { TextSelection } from '@tiptap/pm/state'

/**
 * Word's Copy Text: with text selected it asks where the copy goes, with a
 * bare caret it asks what to copy to the caret. Return carries it out, Escape
 * or an edit leaves the mode. The clipboard is not used.
 */
type Mode = { editor: Editor; stop: () => void } & (
  { kind: 'to'; slice: Slice } | { kind: 'from'; at: number }
)

let mode: Mode | null = null

export const copyTextModeOn = (): boolean => !!mode

export function leaveCopyText(): void {
  mode?.stop()
  mode = null
}

/** enters the mode; returns which question it asks. `onEnd` runs when the mode is left, however that happens */
export function startCopyText(editor: Editor, onEnd?: () => void): 'to' | 'from' {
  leaveCopyText()
  const { selection } = editor.state
  const stop = () => {
    editor.off('update', leaveCopyText)
    editor.off('destroy', leaveCopyText)
    onEnd?.()
  }
  editor.on('update', leaveCopyText)
  editor.on('destroy', leaveCopyText)
  mode = selection.empty
    ? { editor, stop, kind: 'from', at: selection.from }
    : { editor, stop, kind: 'to', slice: selection.content() }
  return mode.kind
}

function finish(): void {
  if (!mode) return
  const { editor } = mode
  const { state, view } = editor
  const { selection } = state
  const taken = mode
  leaveCopyText()
  if (taken.kind === 'to') {
    view.dispatch(state.tr.replaceSelection(taken.slice).scrollIntoView())
    return
  }
  if (selection.empty) return
  const tr = state.tr.setSelection(TextSelection.create(state.doc, taken.at))
  view.dispatch(tr.replaceSelection(selection.content()).scrollIntoView())
}

/** a key pressed while the mode is on; returns whether the mode took it */
export function copyTextKey(event: KeyboardEvent): boolean {
  if (!mode || !(event.target instanceof Node) || !mode.editor.view.dom.contains(event.target))
    return false
  if (event.key === 'Escape') {
    leaveCopyText()
    return true
  }
  if (event.key !== 'Enter' || event.metaKey || event.ctrlKey || event.altKey || event.shiftKey)
    return false
  finish()
  return true
}
