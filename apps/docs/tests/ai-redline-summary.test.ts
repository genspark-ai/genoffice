import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { buildRevisionsContext } from '../src/renderer/ai/protocol'
import { executeTool } from '../src/renderer/ai/tools'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const text = (t: string, marks?: JsonNode['marks']): JsonNode => ({
  type: 'text',
  text: t,
  ...(marks && marks.length > 0 ? { marks } : {}),
})
const para = (content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
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

const NUM_IDS = { bullet: null, ordered: null }

const fixture = () => [
  para([text('Unchanged intro.')]),
  para([
    text('The fee is '),
    text('5%', [{ type: 'del', attrs: { author: 'Bob', date: '2026-08-20T10:00:00Z' } }]),
    text('8%', [{ type: 'ins', attrs: { author: 'Alice', date: '2026-08-21T10:00:00Z' } }]),
    text(' of revenue.'),
  ]),
  para([text('Closing paragraph.')]),
]

describe('read_revisions', () => {
  it('lists pending revisions with kind, author, date, block index and text', async () => {
    const editor = createEditor(fixture())
    const exec = await executeTool(editor, { id: 't', name: 'read_revisions', input: {} }, NUM_IDS)
    expect(exec.isError).toBeFalsy()
    expect(exec.mutated).toBe(false)
    expect(exec.output).toContain('block 1 | deleted by Bob on 2026-08-20: "5%"')
    expect(exec.output).toContain('block 1 | inserted by Alice on 2026-08-21: "8%"')
  })

  it('text inserted then deleted while tracking is not labeled a replacement', () => {
    const editor = createEditor([
      para([
        text('Keep this. '),
        text('ephemeral', [
          { type: 'ins', attrs: { author: 'Alice', date: '2026-08-21T10:00:00Z' } },
          { type: 'del', attrs: { author: 'Bob', date: '2026-08-22T10:00:00Z' } },
        ]),
      ]),
    ])
    const output = buildRevisionsContext(editor)
    expect(output).toContain('inserted then deleted (not in the base document)')
    expect(output).not.toContain('replaced')
  })

  it('reports an empty document state instead of an empty string', () => {
    const editor = createEditor([para([text('Nothing tracked here.')])])
    expect(buildRevisionsContext(editor)).toContain('no tracked revisions')
  })
})
