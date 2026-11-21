import type { Editor } from '@tiptap/core'
import type { Node as PMNode } from '@tiptap/pm/model'
import type { ActionCommandId } from '../../shared/commands'
import { editsDocument } from './menu-state'

/**
 * Edit > Repeat runs the last command again on the current selection. An editing
 * command is remembered when its handler changed the document before returning, and
 * stays repeatable only while the document is still the one it left behind:
 * typing, a dialog's own apply, undo or an AI edit all end it.
 */
interface LastCommand {
  command: ActionCommandId
  payload: string | undefined
  /** the editors the command could have written to, each with the document it left */
  docs: [Editor, PMNode][]
  /** sub-editors opened after the command ran: any edit in one came after it */
  entered: WeakSet<Editor>
}

let last: LastCommand | null = null

const snapshot = (editors: (Editor | null)[]): [Editor, PMNode][] =>
  editors.filter((e): e is Editor => !!e && !e.isDestroyed).map((e) => [e, e.state.doc])

/** runs a command's handler and remembers the command when the document changed under it */
export function runRecorded(
  editors: () => (Editor | null)[],
  command: ActionCommandId,
  payload: string | undefined,
  run: () => void,
): void {
  // New, Open and the like also replace the document, and are not something to run again
  if (!editsDocument(command)) return run()
  const before = snapshot(editors())
  const prior = last
  run()
  // the handler ran another command, which is the one to repeat
  if (last !== prior) return
  const after = snapshot(editors())
  const changed = after.some(([editor, doc]) => before.find(([e]) => e === editor)?.[1] !== doc)
  if (changed) last = { command, payload, docs: after, entered: new WeakSet() }
}

/** the command Repeat would run now, if the document has not moved on since it ran */
export function repeatable(
  editors: (Editor | null)[],
): { command: ActionCommandId; payload: string | undefined } | null {
  if (!last) return null
  const recorded = last.docs
  // a text box or header entered since is a new place to repeat in, not a change
  for (const entry of snapshot(editors)) {
    if (recorded.some(([editor]) => editor === entry[0])) continue
    recorded.push(entry)
    last.entered.add(entry[0])
  }
  // a closed text box wrote any edit into the main document; a strip reported its own on closing
  const [main, ...boxes] = recorded
  last.docs = [main, ...boxes.filter(([editor]) => !editor.isDestroyed)]
  const same = last.docs.every(([editor, doc]) => !editor.isDestroyed && editor.state.doc === doc)
  return same ? { command: last.command, payload: last.payload } : null
}

/**
 * A header or footer strip keeps its edits out of the main document, so it says on
 * closing whether it was edited; call before the editor is destroyed.
 */
export function subEditorClosing(editor: Editor, edited: boolean): void {
  if (!last || !edited) return
  const seen = last.docs.find(([e]) => e === editor)
  if (!seen || last.entered.has(editor) || seen[1] !== editor.state.doc) last = null
}
