import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Selection, TextSelection } from '@tiptap/pm/state'

/** 0 is the mode alone; each further F8 takes the next unit around where the mode began */
const UNITS = ['none', 'word', 'sentence', 'paragraph', 'document'] as const

interface Mode {
  editor: Editor
  origin: number
  level: number
  stop: () => void
}

let mode: Mode | null = null
const listeners = new Set<() => void>()

/** Word's Extend Selection mode is on in this editor (in any editor when none is given) */
export function extendModeOn(editor?: Editor): boolean {
  return !!mode && (!editor || mode.editor === editor)
}

/** told when the mode goes on or off */
export function onExtendModeChange(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

export function leaveExtendMode(): void {
  if (!mode) return
  mode.stop()
  mode = null
  for (const listener of listeners) listener()
}

const LEAF = '￼'

function segmentAt(
  text: string,
  offset: number,
  granularity: 'word' | 'sentence',
): [number, number] {
  let found: [number, number] = [0, text.length]
  for (const seg of new Intl.Segmenter(undefined, { granularity }).segment(text)) {
    found = [seg.index, seg.index + seg.segment.length]
    if (offset < found[1]) break
  }
  return found
}

/** [from, to] of the unit of that level around `pos` */
export function unitRange(doc: PmNode, pos: number, level: number): [number, number] {
  const unit = UNITS[level]
  if (unit === 'document') return [Selection.atStart(doc).from, Selection.atEnd(doc).to]
  const $pos = doc.resolve(pos)
  if (unit === 'none' || !$pos.parent.isTextblock) return [pos, pos]
  const start = $pos.start()
  if (unit === 'paragraph') return [start, $pos.end()]
  const text = $pos.parent.textBetween(0, $pos.parent.content.size, LEAF, LEAF)
  const [from, to] = segmentAt(text, $pos.parentOffset, unit)
  return [start + from, start + to]
}

function select(editor: Editor, anchor: number, head: number): void {
  const { state, view } = editor
  const selection = TextSelection.between(state.doc.resolve(anchor), state.doc.resolve(head))
  view.dispatch(state.tr.setSelection(selection).scrollIntoView())
}

function enter(editor: Editor): void {
  leaveExtendMode()
  const dom = editor.view.dom
  // a click extends to where it lands instead of placing the caret
  const onMouseDown = (event: MouseEvent) => {
    if (event.button !== 0 || !mode) return
    const at = editor.view.posAtCoords({ left: event.clientX, top: event.clientY })
    if (!at) return
    event.preventDefault()
    event.stopImmediatePropagation()
    select(editor, editor.state.selection.anchor, at.pos)
    editor.view.focus()
  }
  dom.addEventListener('mousedown', onMouseDown, true)
  editor.on('update', leaveExtendMode)
  editor.on('destroy', leaveExtendMode)
  mode = {
    editor,
    origin: editor.state.selection.anchor,
    level: 0,
    stop: () => {
      dom.removeEventListener('mousedown', onMouseDown, true)
      editor.off('update', leaveExtendMode)
      editor.off('destroy', leaveExtendMode)
    },
  }
  for (const listener of listeners) listener()
}

/** F8: the mode on, then the word, the sentence, the paragraph and the document around its start */
export function growSelection(editor: Editor): void {
  if (mode?.editor !== editor) {
    enter(editor)
    return
  }
  if (mode.level >= UNITS.length - 1) return
  mode.level++
  const { from, to } = editor.state.selection
  const [start, end] = unitRange(editor.state.doc, mode.origin, mode.level)
  select(editor, Math.min(from, start), Math.max(to, end))
}

/** ⇧F8: back one unit, down to the caret the mode began at */
export function shrinkSelection(editor: Editor): void {
  if (mode?.editor !== editor || mode.level === 0) return
  mode.level--
  const [start, end] = unitRange(editor.state.doc, mode.origin, mode.level)
  select(editor, start, end)
}

/** the head moves to just past the next `char`, as typing a character does in Word's mode */
function extendToChar(editor: Editor, char: string): void {
  const { doc, selection } = editor.state
  let found = -1
  doc.nodesBetween(selection.head, doc.content.size, (node, pos) => {
    if (found >= 0) return false
    if (!node.isText) return true
    const from = Math.max(selection.head - pos, 0)
    const at = node.text!.indexOf(char, from)
    if (at >= 0) found = pos + at + char.length
    return false
  })
  if (found >= 0) select(editor, selection.anchor, found)
}

const ARROWS: Record<string, [direction: string, plain: string]> = {
  ArrowLeft: ['left', 'character'],
  ArrowRight: ['right', 'character'],
  ArrowUp: ['backward', 'line'],
  ArrowDown: ['forward', 'line'],
}

/**
 * A key pressed while the mode is on. Escape leaves it with the selection as it is, an arrow
 * extends as its shifted twin would, a character extends to its next occurrence. `caretMove`
 * is the caret move the chord runs outside the mode, when it has one. Returns whether the key
 * was taken.
 */
export function extendModeKey(
  event: KeyboardEvent,
  caretMove: string | null,
  runCaretMove: (payload: string) => void,
): boolean {
  if (!mode || !(event.target instanceof Node) || !mode.editor.view.dom.contains(event.target))
    return false
  const { editor } = mode
  if (event.key === 'Escape') {
    leaveExtendMode()
    return true
  }
  if (caretMove) {
    runCaretMove(caretMove.endsWith('-extend') ? caretMove : `${caretMove}-extend`)
    return true
  }
  const arrow = ARROWS[event.key]
  if (arrow) {
    if (event.shiftKey) return false
    const horizontal = arrow[1] === 'character'
    const granularity =
      horizontal && event.metaKey ? 'lineboundary' : horizontal && event.altKey ? 'word' : arrow[1]
    const { view } = editor
    view.dom.ownerDocument.getSelection()?.modify('extend', arrow[0], granularity)
    ;(view as unknown as { domObserver?: { flush(): void } }).domObserver?.flush()
    return true
  }
  if (event.key.length === 1 && !event.metaKey && !event.ctrlKey) {
    extendToChar(editor, event.key)
    return true
  }
  return false
}
