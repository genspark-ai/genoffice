import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { executeTool } from '../src/renderer/ai/tools'
import { aiDocContentNodes } from '../src/renderer/file-actions'

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

function createEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: null },
          content: [{ type: 'text', text: 'x' }],
        },
      ],
    },
  })
  editors.add(editor)
  return editor
}

const NUM_IDS = { bullet: null, ordered: null }

type DesktopStub = { desktop?: unknown }

/** run one create_document call with a stubbed desktop bridge */
async function runWithDesktop(
  createDocument: ReturnType<typeof vi.fn>,
  input: Record<string, unknown>,
) {
  const editor = createEditor()
  const w = window as unknown as DesktopStub
  const saved = w.desktop
  w.desktop = { createDocument }
  try {
    return await executeTool(editor, { id: 't1', name: 'create_document', input }, NUM_IDS)
  } finally {
    w.desktop = saved
  }
}

describe('create_document tool', () => {
  it('defaults to docx and forwards the request', async () => {
    const createDocument = vi.fn(async () => ({ ok: true }))
    const exec = await runWithDesktop(createDocument, {
      title: 'Summary',
      content: '<h1>Summary</h1><p>Body</p>',
    })
    expect(exec.isError).toBeUndefined()
    expect(createDocument).toHaveBeenCalledWith({
      type: 'docx',
      title: 'Summary',
      content: '<h1>Summary</h1><p>Body</p>',
    })
    expect(exec.output).toContain('Summary.docx')
    expect(exec.mutated).toBe(false)
  })

  it('reports the written path for direct-write types', async () => {
    const createDocument = vi.fn(async () => ({ ok: true, path: '/tmp/Summary.pdf' }))
    const exec = await runWithDesktop(createDocument, {
      type: 'pdf',
      title: 'Summary',
      content: '<p>Body</p>',
    })
    expect(exec.isError).toBeUndefined()
    expect(exec.output).toContain('/tmp/Summary.pdf')
  })

  it('rejects bad input without calling the bridge', async () => {
    const createDocument = vi.fn(async () => ({ ok: true }))
    for (const input of [
      { type: 'xlsx', title: 'T', content: '<p>x</p>' },
      { title: '  ', content: '<p>x</p>' },
      { title: 'T', content: '   ' },
    ]) {
      const exec = await runWithDesktop(createDocument, input)
      expect(exec.isError).toBe(true)
    }
    expect(createDocument).not.toHaveBeenCalled()
  })

  it('rejects docx content that parses into no blocks, so the model can retry', async () => {
    const createDocument = vi.fn(async () => ({ ok: true }))
    const exec = await runWithDesktop(createDocument, { title: 'T', content: '<p></p>' })
    expect(exec.isError).toBe(true)
    expect(createDocument).not.toHaveBeenCalled()
  })

  it('skips the HTML validation for markdown targets', async () => {
    const createDocument = vi.fn(async () => ({ ok: true, path: '/tmp/notes.md' }))
    const exec = await runWithDesktop(createDocument, {
      type: 'md',
      title: 'notes',
      content: '# Notes\n\n- a\n- b',
    })
    expect(exec.isError).toBeUndefined()
    expect(createDocument).toHaveBeenCalledOnce()
  })

  it('surfaces a main-process failure', async () => {
    const createDocument = vi.fn(async () => ({ ok: false, error: 'disk full' }))
    const exec = await runWithDesktop(createDocument, { title: 'T', content: '<p>x</p>' })
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('disk full')
  })
})

describe('aiDocContentNodes (boot-time fill of the new docx tab)', () => {
  it('parses restricted HTML into blocks with the aiChanged highlight stripped', () => {
    const nodes = aiDocContentNodes('<h1>Title</h1><p>Body</p>')
    expect(nodes.map((n) => n.type)).toEqual(['docHeading', 'docParagraph'])
    for (const node of nodes) expect(node.attrs?.aiChanged).toBe(false)
  })

  it('falls back to plain-text paragraphs when the fragment throws (pdf chat cannot pre-parse)', () => {
    const nodes = aiDocContentNodes('<p>Before</p><formula>\\frac{</formula>')
    expect(nodes.length).toBeGreaterThan(0)
    expect(nodes.every((n) => n.type === 'docParagraph')).toBe(true)
    const text = JSON.stringify(nodes)
    expect(text).toContain('Before')
  })

  it('keeps adjacent minified blocks separate in the plain-text salvage', () => {
    const nodes = aiDocContentNodes('<p>One</p><p>Two</p><formula>\\frac{</formula>')
    const texts = nodes.map((n) => JSON.stringify(n))
    expect(texts.some((t) => t.includes('One') && !t.includes('Two'))).toBe(true)
    expect(texts.some((t) => t.includes('Two') && !t.includes('One'))).toBe(true)
  })

  it('returns nothing for markup with no text at all', () => {
    expect(aiDocContentNodes('<p></p>')).toEqual([])
  })
})
