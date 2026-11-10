import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { undoDepth } from '@tiptap/pm/history'
import { buildExtensions } from '../src/renderer/editor/extensions'
import {
  buildDocWriterRequest,
  countMarkdownBlocks,
  DraftLanding,
  extractMarkdown,
  type AiDocWriter,
} from '../src/renderer/ai/doc-writer'
import { executeTool } from '../src/renderer/ai/tools'
import { MARKDOWN_RULES } from '../src/renderer/ai/markdown-skill'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

function createEditor(md = ''): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
    }),
    content: '',
  })
  if (md) editor.commands.setContent(md, { contentType: 'markdown' })
  editors.push(editor)
  return editor
}

const texts = (editor: Editor) => {
  const out: string[] = []
  editor.state.doc.forEach((n) => out.push(n.textContent))
  return out
}

const tick = () => new Promise((r) => setTimeout(r, 350))

describe('extractMarkdown', () => {
  it('unwraps a whole-reply markdown fence but keeps a leading code block', () => {
    expect(extractMarkdown('```markdown\n# T\n\ntext\n```\n')).toBe('# T\n\ntext')
    expect(extractMarkdown('```js\nlet a\n```\n\nafter')).toBe('```js\nlet a\n```\n\nafter')
    expect(extractMarkdown('\n\n# T\r\nx')).toBe('# T\nx')
  })
  it('counts blank-line separated blocks', () => {
    expect(countMarkdownBlocks('# T\n\npara\n\n- a\n- b\n\n')).toBe(3)
  })
})

describe('buildDocWriterRequest', () => {
  it('carries plan, title and context; the system prompt has the markdown rules and language', () => {
    const { system, user } = buildDocWriterRequest(
      { plan: 'intro, body', title: 'Guide', context: 'founded 1999' },
      MARKDOWN_RULES,
      '\n\nAnswer in French.',
    )
    expect(system).toContain('Reply with the markdown only')
    expect(system).toContain('pure GFM plus math')
    expect(system.endsWith('Answer in French.')).toBe(true)
    expect(user).toContain('Title: Guide')
    expect(user).toContain('intro, body')
    expect(user).toContain('founded 1999')
  })
})

describe('DraftLanding', () => {
  it('renders the growing markdown in place, keeps unchanged blocks, leaves no trace on finish', async () => {
    const editor = createEditor()
    const before = editor.state.doc
    const depth = undoDepth(editor.state)
    const draft = new DraftLanding(editor, { kind: 'whole' })
    draft.update('# Title\n\nfirst para')
    await tick()
    expect(texts(editor)).toEqual(['Title', 'first para', ''])
    const titleNode = editor.state.doc.child(0)
    draft.update('# Title\n\nfirst paragraph\n\n- item')
    await tick()
    expect(texts(editor)).toEqual(['Title', 'first paragraph', 'item', ''])
    expect(editor.state.doc.child(0)).toBe(titleNode)
    expect(undoDepth(editor.state)).toBe(depth)
    draft.finish()
    expect(editor.state.doc.eq(before)).toBe(true)
  })

  it("follows the user's edits above the draft and reports the block it sits after", async () => {
    const editor = createEditor('one\n\ntwo')
    const draft = new DraftLanding(editor, { kind: 'after', index: 0 })
    draft.update('draft')
    await tick()
    expect(texts(editor)).toEqual(['one', 'draft', 'two'])
    editor.view.dispatch(editor.state.tr.insert(0, editor.schema.nodes.paragraph!.create()))
    expect(draft.indexBefore()).toBe(1)
    draft.update('draft\n\nmore')
    await tick()
    expect(texts(editor)).toEqual(['', 'one', 'draft', 'more', 'two'])
    draft.finish()
    expect(texts(editor)).toEqual(['', 'one', 'two'])
  })
})

describe('DraftLanding under user edits', () => {
  it('re-anchors on the live document when the user edits inside the draft', async () => {
    const editor = createEditor('one')
    const draft = new DraftLanding(editor, { kind: 'after', index: 0 })
    draft.update('alpha\n\nbeta')
    await tick()
    expect(texts(editor)).toEqual(['one', 'alpha', 'beta'])
    const alphaPos = editor.state.doc.child(0).nodeSize + 1 + 'alpha'.length
    editor.view.dispatch(editor.state.tr.insertText(' edited', alphaPos))
    draft.update('alpha\n\nbeta\n\ngamma')
    await tick()
    expect(texts(editor)).toEqual(['one', 'alpha', 'beta', 'gamma'])
    expect(draft.finish()).toBe('alpha\n\nbeta\n\ngamma')
    expect(texts(editor)).toEqual(['one'])
  })
})

describe('write_document tool', () => {
  const writerOf = (markdown: string, truncated = false): AiDocWriter => ({
    write: async (_spec, onProgress) => {
      onProgress(markdown.slice(0, 5))
      onProgress(markdown)
      return { ok: true, markdown, truncated }
    },
  })

  it('lands the streamed markdown as one regular insert on a blank document', async () => {
    const editor = createEditor()
    const exec = await executeTool(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'a title and a paragraph' } },
      undefined,
      undefined,
      writerOf('# T\n\nbody'),
    )
    expect(exec.isError).toBeFalsy()
    expect(exec.mutated).toBe(true)
    expect(texts(editor)).toEqual(['T', 'body'])
    expect(editor.commands.undo()).toBe(true)
    expect(texts(editor)).toEqual([''])
  })

  it('refuses a whole-document write on a non-blank document without replaceDocument', async () => {
    const editor = createEditor('keep')
    const exec = await executeTool(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p' } },
      undefined,
      undefined,
      writerOf('new'),
    )
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('afterIndex')
    expect(texts(editor)).toEqual(['keep'])
  })

  it('replaces everything with replaceDocument and inserts after a block with afterIndex', async () => {
    const editor = createEditor('one\n\ntwo')
    const all = await executeTool(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p', replaceDocument: true } },
      undefined,
      undefined,
      writerOf('# New\n\nbody'),
    )
    expect(all.isError).toBeFalsy()
    expect(texts(editor)).toEqual(['New', 'body'])
    const after = await executeTool(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p', afterIndex: 0 } },
      undefined,
      undefined,
      writerOf('half', true),
    )
    expect(after.isError).toBeFalsy()
    expect(after.output).toContain('INCOMPLETE')
    expect(texts(editor)).toEqual(['New', 'half', 'body'])
  })

  it('keeps text the user typed into a blank document while the draft streamed', async () => {
    const editor = createEditor()
    const writer: AiDocWriter = {
      write: async (_spec, onProgress) => {
        onProgress('draft')
        await tick()
        editor.view.dispatch(editor.state.tr.insertText('mine', editor.state.doc.content.size - 1))
        return { ok: true, markdown: 'draft\n\nmore' }
      },
    }
    const exec = await executeTool(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p' } },
      undefined,
      undefined,
      writer,
    )
    expect(exec.isError).toBeFalsy()
    expect(texts(editor)).toEqual(['draft', 'more', 'mine'])
  })

  it('a discarded write leaves the document unchanged', async () => {
    const editor = createEditor()
    const writer: AiDocWriter = {
      write: async (_spec, onProgress) => {
        onProgress('draft')
        await tick()
        return { ok: false, error: 'the user discarded the partial content' }
      },
    }
    const exec = await executeTool(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p' } },
      undefined,
      undefined,
      writer,
    )
    expect(exec.isError).toBe(true)
    expect(texts(editor)).toEqual([''])
  })
})
