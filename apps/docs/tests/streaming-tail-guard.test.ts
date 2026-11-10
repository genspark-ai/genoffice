import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { PHASED_APPEND, touchesStreamingTail } from '../src/renderer/editor/streaming-tail-guard'
import { appendStreamedNodes } from '../src/renderer/file-actions'
import {
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
function inside(editor: Editor, index: number, offset = 1): number {
  let pos = 0
  for (let i = 0; i < index; i++) pos += editor.state.doc.child(i).nodeSize
  return pos + offset
}

describe('touchesStreamingTail', () => {
  it('flags edits in or after the last block, not earlier ones or the appends themselves', () => {
    const editor = createEditor()
    editor.commands.setContent({ type: 'doc', content: [para('one'), para('two'), para('three')] })
    expect(touchesStreamingTail(editor.state.tr.insertText('x', inside(editor, 0, 2)))).toBe(false)
    expect(touchesStreamingTail(editor.state.tr.insertText('x', inside(editor, 1, 2)))).toBe(false)
    expect(touchesStreamingTail(editor.state.tr.insertText('x', inside(editor, 2, 2)))).toBe(true)
    const end = editor.state.doc.content.size
    expect(
      touchesStreamingTail(editor.state.tr.insert(end, editor.schema.nodeFromJSON(para('new')))),
    ).toBe(true)
    expect(
      touchesStreamingTail(
        editor.state.tr
          .insert(end, editor.schema.nodeFromJSON(para('chunk')))
          .setMeta(PHASED_APPEND, true),
      ),
    ).toBe(false)
    expect(touchesStreamingTail(editor.state.tr.setMeta('x', 1))).toBe(false)
  })
})

describe('editing while a phased open streams', () => {
  it('accepts edits in mounted blocks, refuses the tail, keeps the user undo stack', () => {
    const editor = createEditor()
    const chunks: Array<() => void> = []
    let resets = 0
    let dirty = false
    const host: PhasedContentHost = {
      // loadFile resets the history right after the first mount; keep the
      // mount out of it here the simple way
      setContent: (doc) =>
        editor
          .chain()
          .setMeta('addToHistory', false)
          .setMeta(PHASED_APPEND, true)
          .setContent(doc as never)
          .run(),
      appendNodes: (nodes) => appendStreamedNodes(editor, nodes as never),
      isDestroyed: () => editor.isDestroyed,
      resetHistory: () => resets++,
      setLoading: () => {},
      getDirty: () => dirty,
      setDirty: (d) => {
        dirty = d
      },
    }
    const blocks = Array.from({ length: PHASED_MIN_BLOCKS + 10 }, (_, i) => para(`block ${i}`))
    setContentPhased(host, { type: 'doc', content: blocks } as never, (cb) => chunks.push(cb))
    expect(isPhasedContentPending()).toBe(true)
    const mounted = editor.state.doc.childCount
    expect(mounted).toBeLessThan(blocks.length)

    // an edit in the first block goes through
    editor.view.dispatch(editor.state.tr.insertText('!', inside(editor, 0, 2)))
    expect(editor.state.doc.child(0).textContent).toBe('b!lock 0')
    dirty = true
    // an edit in the last mounted block is refused
    const lastText = editor.state.doc.child(mounted - 1).textContent
    editor.view.dispatch(editor.state.tr.insertText('!', inside(editor, mounted - 1, 2)))
    expect(editor.state.doc.child(mounted - 1).textContent).toBe(lastText)
    expect(editor.state.doc.childCount).toBe(mounted)

    // the tail lands; the user's undo stack survives (no history reset)
    while (chunks.length) chunks.shift()!()
    expect(isPhasedContentPending()).toBe(false)
    expect(editor.state.doc.childCount).toBe(blocks.length)
    expect(editor.state.doc.child(blocks.length - 1).textContent).toBe(`block ${blocks.length - 1}`)
    expect(resets).toBe(0)
    editor.commands.undo()
    expect(editor.state.doc.child(0).textContent).toBe('block 0')
    expect(editor.state.doc.childCount).toBe(blocks.length)
  })
})

describe('a refused chunk during a phased open', () => {
  it('falls back to the full mount through the guard', () => {
    const editor = createEditor()
    const chunks: Array<() => void> = []
    let appends = 0
    let resets = 0
    const host: PhasedContentHost = {
      setContent: (doc) =>
        editor
          .chain()
          .setMeta('addToHistory', false)
          .setMeta(PHASED_APPEND, true)
          .setContent(doc as never)
          .run(),
      appendNodes: () => {
        appends++
        throw new Error('refused')
      },
      isDestroyed: () => editor.isDestroyed,
      resetHistory: () => resets++,
      setLoading: () => {},
      getDirty: () => true,
      setDirty: () => {},
    }
    const blocks = Array.from({ length: PHASED_MIN_BLOCKS + 10 }, (_, i) => para(`block ${i}`))
    setContentPhased(host, { type: 'doc', content: blocks } as never, (cb) => chunks.push(cb))
    editor.view.dispatch(editor.state.tr.insertText('!', inside(editor, 0, 2)))
    while (chunks.length) chunks.shift()!()
    expect(appends).toBe(1)
    expect(isPhasedContentPending()).toBe(false)
    // the whole document landed even though the guard was armed at the time
    expect(editor.state.doc.childCount).toBe(blocks.length)
    expect(editor.state.doc.child(blocks.length - 1).textContent).toBe(`block ${blocks.length - 1}`)
    // the remount discarded the mid-stream edit, so the history goes with it:
    // an undo must not bring the truncated document back
    expect(resets).toBe(1)
  })
})
