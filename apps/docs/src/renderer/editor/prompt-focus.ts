/**
 * The editor's `focus` command. The stock one focuses a frame later and writes the selection to
 * the page there: a click made in between is undone, keys typed in between go nowhere, a dialog
 * opened in between loses the focus. This one focuses ahead of any further input, unless something
 * else took the focus; the scroll stays a frame later, for layout to follow.
 */
import { Extension, isTextSelection, resolveFocusPosition } from '@tiptap/core'
import type { Editor, RawCommands } from '@tiptap/core'

let inputs = 0
let clicks = 0
/** the target of a key or button going down, until that event has run its course */
let pressed: EventTarget | null = null
let counting = false

const queued: Array<() => void> = []
let port: MessagePort | null = null

/** runs once the event being handled is over; a timer would be held back in a window that is not in front */
function afterEvent(run: () => void): void {
  if (!port) {
    const channel = new MessageChannel()
    channel.port1.onmessage = () => queued.shift()?.()
    ;(channel.port1 as { unref?: () => void }).unref?.()
    port = channel.port2
  }
  queued.push(run)
  port.postMessage(null)
}

function countInputs(): void {
  if (counting || typeof window === 'undefined') return
  counting = true
  const count = (e: Event): void => {
    // the mousedown of a press is the same input as its pointerdown
    if (e.type !== 'mousedown') inputs++
    if (e.type === 'wheel') return
    if (e.type === 'pointerdown') clicks++
    pressed = e.target
    afterEvent(() => (pressed = null))
  }
  for (const type of ['pointerdown', 'mousedown', 'keydown', 'wheel'])
    window.addEventListener(type, count, { capture: true, passive: true })
  // input is handled ahead of other tasks: the release may come before the press was cleared
  for (const type of ['pointerup', 'mouseup', 'keyup'])
    window.addEventListener(type, () => (pressed = null), { capture: true, passive: true })
}

const unclaimed = (active: Element | null): boolean =>
  !active || active === document.body || !active.isConnected

function focusPromptly(editor: Editor, scrollIntoView: boolean): void {
  countInputs()
  const before = document.activeElement
  const from = pressed
  const at = inputs
  const clicked = clicks
  const take = (): void => {
    if (editor.isDestroyed || editor.view.hasFocus()) return
    const active = document.activeElement
    const own = active === before || (from instanceof Node && !!active?.contains(from))
    if (unclaimed(active) || own) editor.view.focus()
  }
  // a key or button still going down acts on whatever has the focus when its handlers are done:
  // Enter in a dialog would be typed into the page, the pressed button would take the focus back
  if (from) afterEvent(() => clicks === clicked && take())
  else queueMicrotask(take)
  requestAnimationFrame(() => {
    if (editor.isDestroyed || inputs !== at) return
    // dropped since, by a control that closed while it held the focus
    if (!editor.view.hasFocus() && unclaimed(document.activeElement)) editor.view.focus()
    if (scrollIntoView) editor.commands.scrollIntoView()
  })
}

const focus: RawCommands['focus'] =
  (position = null, options = {}) =>
  ({ editor, view, tr, dispatch }) => {
    const scrollIntoView = options.scrollIntoView ?? true
    try {
      if ((view.hasFocus() && position === null) || position === false) return true
    } catch {
      return false
    }
    // a node or cell selection is left as it is
    if (dispatch && position === null && !isTextSelection(editor.state.selection)) {
      focusPromptly(editor, scrollIntoView)
      return true
    }
    const selection = resolveFocusPosition(tr.doc, position) || editor.state.selection
    const same = editor.state.selection.eq(selection)
    if (dispatch) {
      if (!same) tr.setSelection(selection)
      // setSelection drops the stored marks
      if (same && tr.storedMarks) tr.setStoredMarks(tr.storedMarks)
      focusPromptly(editor, scrollIntoView)
    }
    return true
  }

export const PromptFocusExtension = Extension.create({
  name: 'promptFocus',
  // commands of a lower priority replace the core ones of the same name
  priority: 99,
  addCommands() {
    return { focus }
  },
})
