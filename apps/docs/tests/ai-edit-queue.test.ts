import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  addQueueAnchor,
  clearQueueAnchors,
  queueAnchorRange,
  removeQueueAnchors,
} from '../src/renderer/editor/ai-queue-anchors'
import {
  buildQueueInstruction,
  buildQueueSummary,
  liveItems,
  resolveQueue,
  resolveQueueItem,
  type DocsEditQueueItem,
} from '../src/renderer/ai/edit-queue'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const para = (t: string): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [text(t)],
})
const heading = (t: string): JsonNode => ({
  type: 'docHeading',
  attrs: { docxIndex: null, level: 1 },
  content: [text(t)],
})

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.add(editor)
  return editor
}

const fixture = () => [
  heading('Chapter 1'),
  para('First paragraph with some text.'),
  para('Second paragraph, the anchored one.'),
  para('Third paragraph at the end.'),
]

/** positions of block `index`'s text [start, end] */
function blockTextRange(editor: Editor, index: number): { from: number; to: number } {
  let pos = 0
  for (let i = 0; i < index; i++) pos += editor.state.doc.child(i).nodeSize
  return { from: pos + 1, to: pos + 1 + editor.state.doc.child(index).content.size }
}

const item = (qid: string, instruction = 'shorten it'): DocsEditQueueItem => ({
  qid,
  instruction,
  capturedText: 'captured',
})

describe('queue anchors', () => {
  it('anchors migrate through edits made before them', () => {
    const editor = createEditor(fixture())
    const { from, to } = blockTextRange(editor, 2)
    addQueueAnchor(editor, 'q1', from, to)
    // type at the start of the document: the anchor must shift, not stretch
    editor.commands.insertContentAt(1, 'XYZ ')
    const range = queueAnchorRange(editor.state, 'q1')!
    expect(range.from).toBe(from + 4)
    expect(range.to).toBe(to + 4)
    const resolved = resolveQueueItem(editor, item('q1'))
    expect(resolved.target).toEqual({
      startIndex: 2,
      endIndex: 2,
      excerpt: 'Second paragraph, the anchored one.',
    })
  })

  it('an anchor whose text is deleted resolves as gone', () => {
    const editor = createEditor(fixture())
    const { from, to } = blockTextRange(editor, 2)
    addQueueAnchor(editor, 'q1', from, to)
    editor.commands.deleteRange({ from: from - 1, to: to + 1 })
    expect(queueAnchorRange(editor.state, 'q1')).toBeNull()
    expect(resolveQueueItem(editor, item('q1')).target).toBeNull()
  })

  it('the excerpt follows live content edited inside the anchor', () => {
    const editor = createEditor(fixture())
    const { from, to } = blockTextRange(editor, 1)
    addQueueAnchor(editor, 'q1', from, to)
    editor.commands.insertContentAt(from + 6, 'EDITED ')
    const resolved = resolveQueueItem(editor, item('q1'))
    expect(resolved.target?.excerpt).toContain('EDITED')
  })

  it('remove and clear drop anchors', () => {
    const editor = createEditor(fixture())
    const a = blockTextRange(editor, 1)
    const b = blockTextRange(editor, 2)
    addQueueAnchor(editor, 'q1', a.from, a.to)
    addQueueAnchor(editor, 'q2', b.from, b.to)
    removeQueueAnchors(editor, ['q1'])
    expect(queueAnchorRange(editor.state, 'q1')).toBeNull()
    expect(queueAnchorRange(editor.state, 'q2')).not.toBeNull()
    clearQueueAnchors(editor)
    expect(queueAnchorRange(editor.state, 'q2')).toBeNull()
  })

  it('anchor transactions do not enter undo history', () => {
    const editor = createEditor(fixture())
    const { from, to } = blockTextRange(editor, 1)
    addQueueAnchor(editor, 'q1', from, to)
    expect(editor.can().undo()).toBe(false)
  })
})

describe('batch instruction', () => {
  it('lists edits bottom-up with block labels and quoted target text', () => {
    const editor = createEditor(fixture())
    const a = blockTextRange(editor, 1)
    const b = blockTextRange(editor, 3)
    addQueueAnchor(editor, 'q1', a.from, a.to)
    addQueueAnchor(editor, 'q2', b.from, b.to)
    const entries = liveItems(
      resolveQueue(editor, [item('q1', 'polish this'), item('q2', 'translate this')]),
    )
    const instruction = buildQueueInstruction(entries)
    // bottom-up: block 3's edit is listed before block 1's
    expect(instruction.indexOf('Block 3')).toBeLessThan(instruction.indexOf('Block 1'))
    expect(instruction).toContain('"Third paragraph at the end."')
    expect(instruction).toContain('Requested change: translate this')
    expect(instruction).toContain('Requested change: polish this')
  })

  it('summary echoes excerpts and instructions in queue order', () => {
    const editor = createEditor(fixture())
    const a = blockTextRange(editor, 1)
    addQueueAnchor(editor, 'q1', a.from, a.to)
    const entries = liveItems(resolveQueue(editor, [item('q1', 'polish this')]))
    const summary = buildQueueSummary('Batch of 1 edits:', entries)
    expect(summary).toContain('Batch of 1 edits:')
    expect(summary).toContain('polish this')
  })
})

describe('textless anchors (image/chart blocks)', () => {
  const image = (): JsonNode => ({
    type: 'docProtected',
    attrs: { docxIndex: null, blockType: 'image', label: 'Image', imageWidthPx: 10 },
  })

  it('resolves a node-selection anchor over an image with a type placeholder', () => {
    const editor = createEditor([para('intro'), image(), para('outro')])
    let imagePos = -1
    editor.state.doc.forEach((node, offset) => {
      if (node.type.name === 'docProtected') imagePos = offset
    })
    addQueueAnchor(editor, 'img', imagePos, imagePos + 1)
    const resolved = resolveQueueItem(editor, item('img', 'replace this image'))
    expect(resolved.target).not.toBeNull()
    expect(resolved.target!.startIndex).toBe(1)
    expect(resolved.target!.excerpt).toContain('image')
  })

  it('focusing an anchored image dispatches a node selection without throwing', async () => {
    const { NodeSelection, TextSelection } = await import('@tiptap/pm/state')
    const { selectionForAnchor } = await import('../src/renderer/ai/edit-queue')
    const editor = createEditor([para('intro text'), image()])
    let imagePos = -1
    editor.state.doc.forEach((node, offset) => {
      if (node.type.name === 'docProtected') imagePos = offset
    })
    addQueueAnchor(editor, 'img', imagePos, imagePos + 1)
    const { from, to } = blockTextRange(editor, 0)
    addQueueAnchor(editor, 'txt', from, to)
    const imgSel = selectionForAnchor(editor, 'img')
    expect(imgSel).toBeInstanceOf(NodeSelection)
    editor.view.dispatch(editor.state.tr.setSelection(imgSel!))
    expect(selectionForAnchor(editor, 'txt')).toBeInstanceOf(TextSelection)
    expect(selectionForAnchor(editor, 'missing')).toBeNull()
  })
})
