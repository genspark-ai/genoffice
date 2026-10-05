/**
 * A keystroke that landed on the streaming tail of a large document while it
 * was still loading used to vanish whole. The boundary guard refused the
 * transaction, so the document never changed, no undo step was recorded, and
 * Ctrl+Z had nothing to bring back — the user could not even tell that
 * anything had happened. The guard may still refuse the *position* (an edit
 * there would land in front of blocks still to come), but the text itself is
 * held until the tail has streamed in and is then inserted as a normal,
 * undoable edit.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { PHASED_APPEND } from '../src/renderer/editor/streaming-tail-guard'
import { appendStreamedNodes } from '../src/renderer/file-actions'
import {
  PHASE1_BLOCKS,
  PHASED_MIN_BLOCKS,
  cancelPhasedContent,
  isPhasedContentPending,
  setContentPhased,
  type PhasedContentHost,
} from '../src/renderer/phased-content'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}
const para = (t: string): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [{ type: 'text', text: t }],
})

const liveEditors: Editor[] = []
afterEach(() => {
  cancelPhasedContent()
  for (const e of liveEditors.splice(0)) e.destroy()
})

function createEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [para('seed')] },
  })
  liveEditors.push(editor)
  return editor
}

function blockStart(editor: Editor, index: number): number {
  let pos = 0
  for (let i = 0; i < index; i++) pos += editor.state.doc.child(i).nodeSize
  return pos
}
const endOfBlock = (editor: Editor, index: number): number =>
  blockStart(editor, index) + editor.state.doc.child(index).nodeSize - 1

/** feed characters the way the browser does: through handleTextInput */
function type(editor: Editor, text: string): void {
  for (const ch of text) {
    const { view } = editor
    const { from, to } = view.state.selection
    const deflt = () => view.state.tr.insertText(ch, from, to)
    if (!view.someProp('handleTextInput', (f) => f(view, from, to, ch, deflt)))
      view.dispatch(deflt())
  }
}

/** mounts the first phase and returns the scheduled tail chunks */
function startStreamingOpen(editor: Editor, blocks: number): Array<() => void> {
  const chunks: Array<() => void> = []
  const host: PhasedContentHost = {
    setContent: (doc) =>
      editor
        .chain()
        .setMeta('addToHistory', false)
        .setMeta(PHASED_APPEND, true)
        .setContent(doc as never)
        .run(),
    appendNodes: (nodes) => appendStreamedNodes(editor, nodes as never),
    isDestroyed: () => editor.isDestroyed,
    resetHistory: () => {},
    setLoading: () => {},
    getDirty: () => false,
    setDirty: () => {},
  }
  const doc = { type: 'doc', content: Array.from({ length: blocks }, (_, i) => para(`block ${i}`)) }
  setContentPhased(host, doc as never, (cb) => chunks.push(cb))
  return chunks
}

const TAIL_DOC = PHASED_MIN_BLOCKS + 10

describe('typing on the streaming tail of a phased open', () => {
  it('keeps the character and inserts it as an undoable edit once the tail lands', () => {
    const editor = createEditor()
    const chunks = startStreamingOpen(editor, TAIL_DOC)
    const mounted = editor.state.doc.childCount
    expect(mounted).toBe(PHASE1_BLOCKS)
    expect(mounted).toBeLessThan(TAIL_DOC)
    expect(isPhasedContentPending()).toBe(true)

    // the caret sits at the end of the last mounted block: the tail boundary
    editor.commands.setTextSelection(endOfBlock(editor, mounted - 1))
    type(editor, 'Z')

    // the tail has not landed, so nothing may have been inserted in its place
    expect(editor.state.doc.child(mounted - 1).textContent).toBe(`block ${mounted - 1}`)

    while (chunks.length) chunks.shift()!()
    expect(isPhasedContentPending()).toBe(false)
    expect(editor.state.doc.childCount).toBe(TAIL_DOC)

    // the character survived the stream, at the position it was typed
    expect(editor.state.doc.child(mounted - 1).textContent).toBe(`block ${mounted - 1}Z`)

    // and it is a normal edit, so Ctrl+Z brings it back
    editor.commands.undo()
    expect(editor.state.doc.child(mounted - 1).textContent).toBe(`block ${mounted - 1}`)
    expect(editor.state.doc.childCount).toBe(TAIL_DOC)
  })

  it('inserts a keystroke in a mounted block immediately, without holding it', () => {
    const editor = createEditor()
    const chunks = startStreamingOpen(editor, TAIL_DOC)
    expect(isPhasedContentPending()).toBe(true)

    editor.commands.setTextSelection(endOfBlock(editor, 0))
    type(editor, 'Z')

    // typed where the guard allows it: applied at once, never queued
    expect(editor.state.doc.child(0).textContent).toBe('block 0Z')
    while (chunks.length) chunks.shift()!()
    expect(editor.state.doc.child(0).textContent).toBe('block 0Z')
    expect(editor.state.doc.childCount).toBe(TAIL_DOC)
  })
})
