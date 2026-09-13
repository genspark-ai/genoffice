import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { undoDepth } from '@tiptap/pm/history'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  parseDocx,
} from '@genoffice/docx-engine'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import {
  buildDocWriterRequest,
  countFragmentBlocks,
  DraftLanding,
  extractFragment,
  type AiDocWriter,
} from '../src/renderer/ai/doc-writer'
import { executeTool } from '../src/renderer/ai/tools'

const NUM_IDS = { bullet: BLANK_BULLET_NUM_ID, ordered: BLANK_ORDERED_NUM_ID }

async function createBlankEditor() {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const parsed = await parseDocx(await buildBlankDocx())
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return editor
}

const texts = (editor: Editor) => {
  const out: string[] = []
  editor.state.doc.forEach((n) => out.push(n.textContent))
  return out
}

const tick = () => new Promise((r) => setTimeout(r, 350))

describe('extractFragment', () => {
  it('drops chatter and fences around the fragment', () => {
    expect(extractFragment('Here you go:\n```html\n<h1>T</h1>\n<p>a</p>\n```\n')).toBe(
      '<h1>T</h1>\n<p>a</p>',
    )
    expect(extractFragment('```html\n<p>half')).toBe('<p>half')
    expect(extractFragment('<p>a</p>\r\n<p>b</p>')).toBe('<p>a</p>\n<p>b</p>')
  })
  it('counts top-level blocks roughly', () => {
    expect(
      countFragmentBlocks(
        '<h1>T</h1><p>a</p><ul><li>x</li></ul><table><tr><td>1</td></tr></table>',
      ),
    ).toBe(4)
  })
})

describe('buildDocWriterRequest', () => {
  it('carries the plan, title, context and request; the system prompt has the HTML rules', () => {
    const { system, user } = buildDocWriterRequest(
      { plan: 'intro, body', title: 'Guide', context: 'founded 1999', instruction: 'write it' },
      '\n\nAnswer in French.',
    )
    expect(system).toContain('restricted HTML fragment only')
    expect(system).toContain('Only these tags are allowed')
    expect(system.endsWith('Answer in French.')).toBe(true)
    expect(user).toContain('Title: Guide')
    expect(user).toContain('intro, body')
    expect(user).toContain('founded 1999')
    expect(user).toContain('write it')
  })
})

describe('DraftLanding', () => {
  it('renders the growing fragment in place, swaps only changed blocks, and leaves no trace on finish', async () => {
    const editor = await createBlankEditor()
    const before = editor.state.doc
    const depth = undoDepth(editor.state)
    const draft = new DraftLanding(editor, NUM_IDS, { kind: 'whole' })
    draft.update('<h1>Title</h1>\n<p>first para')
    await tick()
    expect(texts(editor)).toEqual(['Title', 'first para', ''])
    const titleNode = editor.state.doc.child(0)
    draft.update('<h1>Title</h1>\n<p>first paragraph</p>\n<p>second')
    await tick()
    expect(texts(editor)).toEqual(['Title', 'first paragraph', 'second', ''])
    // the unchanged heading was kept, not re-inserted
    expect(editor.state.doc.child(0)).toBe(titleNode)
    // the draft is neither history nor a tracked change
    expect(undoDepth(editor.state)).toBe(depth)
    draft.finish()
    expect(editor.state.doc.eq(before)).toBe(true)
  })

  it("follows the user's edits above the draft and reports the block it sits after", async () => {
    const editor = await createBlankEditor()
    editor.commands.setContent('<p>one</p><p>two</p>')
    const draft = new DraftLanding(editor, NUM_IDS, { kind: 'after', index: 0 })
    draft.update('<p>draft</p>')
    await tick()
    expect(texts(editor)).toEqual(['one', 'draft', 'two'])
    // the user inserts a paragraph at the very start
    editor.view.dispatch(editor.state.tr.insert(0, editor.schema.nodes.docParagraph!.create()))
    expect(draft.indexBefore()).toBe(1)
    draft.update('<p>draft</p><p>more</p>')
    await tick()
    expect(texts(editor)).toEqual(['', 'one', 'draft', 'more', 'two'])
    draft.finish()
    expect(texts(editor)).toEqual(['', 'one', 'two'])
  })
})

describe('DraftLanding under user edits', () => {
  it('re-anchors on the live document when the user edits inside the draft', async () => {
    const editor = await createBlankEditor()
    editor.commands.setContent('<p>one</p>')
    const draft = new DraftLanding(editor, NUM_IDS, { kind: 'after', index: 0 })
    draft.update('<p>alpha</p><p>beta</p>')
    await tick()
    expect(texts(editor)).toEqual(['one', 'alpha', 'beta'])
    // the user types inside the first draft block: its size changes
    const alphaPos = editor.state.doc.child(0).nodeSize + 1 + 'alpha'.length
    editor.view.dispatch(editor.state.tr.insertText(' edited', alphaPos))
    draft.update('<p>alpha</p><p>beta</p><p>gamma</p>')
    await tick()
    expect(texts(editor)).toEqual(['one', 'alpha', 'beta', 'gamma'])
    draft.finish()
    expect(texts(editor)).toEqual(['one'])
  })

  it('finish returns the last fragment that rendered', async () => {
    const editor = await createBlankEditor()
    const draft = new DraftLanding(editor, NUM_IDS, { kind: 'whole' })
    draft.update('<p>ok</p>')
    await tick()
    draft.update('<p>ok</p><formula>\\frac{</formula>')
    await tick()
    expect(texts(editor)).toEqual(['ok', ''])
    expect(draft.finish()).toBe('<p>ok</p>')
  })
})

describe('write_document tool', () => {
  const writerOf = (html: string, truncated = false): AiDocWriter => ({
    write: async (_spec, onProgress) => {
      onProgress(html.slice(0, 10))
      onProgress(html)
      return { ok: true, html, truncated }
    },
  })

  it('lands the streamed fragment as one regular insert on a blank document', async () => {
    const editor = await createBlankEditor()
    const exec = await executeTool(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'a title and a paragraph' } },
      NUM_IDS,
      undefined,
      undefined,
      null,
      undefined,
      undefined,
      writerOf('<h1>T</h1><p>body</p>'),
    )
    expect(exec.isError).toBeFalsy()
    expect(exec.mutated).toBe(true)
    expect(exec.output).toContain('2 block(s)')
    expect(texts(editor)).toEqual(['T', 'body'])
    // one undo step removes the whole write
    expect(editor.commands.undo()).toBe(true)
    expect(texts(editor)).toEqual([''])
  })

  it('refuses a whole-document write on a non-blank document without replaceDocument', async () => {
    const editor = await createBlankEditor()
    editor.commands.setContent('<p>keep</p>')
    const exec = await executeTool(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p' } },
      NUM_IDS,
      undefined,
      undefined,
      null,
      undefined,
      undefined,
      writerOf('<p>new</p>'),
    )
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('afterBlockIndex')
    expect(texts(editor)).toEqual(['keep'])
  })

  it('inserts after a block and flags a kept partial as incomplete', async () => {
    const editor = await createBlankEditor()
    editor.commands.setContent('<p>one</p><p>two</p>')
    const exec = await executeTool(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p', afterBlockIndex: 0 } },
      NUM_IDS,
      undefined,
      undefined,
      null,
      undefined,
      undefined,
      writerOf('<p>half</p>', true),
    )
    expect(exec.isError).toBeFalsy()
    expect(exec.output).toContain('INCOMPLETE')
    expect(texts(editor)).toEqual(['one', 'half', 'two'])
  })

  it('keeps the rendered draft when a kept partial ends mid-formula', async () => {
    const editor = await createBlankEditor()
    const writer: AiDocWriter = {
      write: async (_spec, onProgress) => {
        onProgress('<p>ok</p>')
        await tick()
        onProgress('<p>ok</p><formula>\\frac{</formula>')
        return { ok: true, html: '<p>ok</p><formula>\\frac{</formula>', truncated: true }
      },
    }
    const exec = await executeTool(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p' } },
      NUM_IDS,
      undefined,
      undefined,
      null,
      undefined,
      undefined,
      writer,
    )
    expect(exec.isError).toBeFalsy()
    expect(texts(editor)).toEqual(['ok'])
  })

  it('keeps text the user typed into a blank document while the draft streamed', async () => {
    const editor = await createBlankEditor()
    const writer: AiDocWriter = {
      write: async (_spec, onProgress) => {
        onProgress('<p>draft</p>')
        await tick()
        // the user types below the draft meanwhile
        editor.view.dispatch(editor.state.tr.insertText('mine', editor.state.doc.content.size - 1))
        return { ok: true, html: '<p>draft</p><p>more</p>' }
      },
    }
    const exec = await executeTool(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p' } },
      NUM_IDS,
      undefined,
      undefined,
      null,
      undefined,
      undefined,
      writer,
    )
    expect(exec.isError).toBeFalsy()
    expect(texts(editor)).toEqual(['draft', 'more', 'mine'])
  })

  it('a failed or discarded write leaves the document unchanged', async () => {
    const editor = await createBlankEditor()
    const writer: AiDocWriter = {
      write: async (_spec, onProgress) => {
        onProgress('<p>draft</p>')
        await tick()
        return { ok: false, error: 'the user discarded the partial content' }
      },
    }
    const exec = await executeTool(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p' } },
      NUM_IDS,
      undefined,
      undefined,
      null,
      undefined,
      undefined,
      writer,
    )
    expect(exec.isError).toBe(true)
    expect(texts(editor)).toEqual([''])
  })
})
