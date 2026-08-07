import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { buildDocContext, executeTool, markDocSeen } from '../src/renderer/ai/tools'
import { deriveAutoFileName } from '../src/renderer/App'

// Undestroyed views leave DOMObserver flush timers that fire after jsdom teardown
// ("document is not defined" unhandled error) — destroy every editor we create.
const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

function createEditor(md = ''): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: '',
  })
  if (md) editor.commands.setContent(md, { contentType: 'markdown' })
  editors.push(editor)
  return editor
}

const call = (name: string, input: Record<string, unknown> = {}) => ({
  id: 't1',
  name,
  input,
})

describe('get_document_context', () => {
  it('reports a blank document', () => {
    const editor = createEditor()
    expect(buildDocContext(editor)).toContain('The document is currently blank.')
  })

  it('lists numbered blocks with type and preview', () => {
    const editor = createEditor('# Title\n\nHello world.\n\n- a\n- b')
    const ctx = buildDocContext(editor)
    expect(ctx).toContain('0 | h1 | Title')
    expect(ctx).toContain('1 | paragraph | Hello world.')
    expect(ctx).toContain('2 | bulletList |')
  })
})

describe('insert_content', () => {
  it('replaces the empty paragraph on a blank document', () => {
    const editor = createEditor()
    const result = executeTool(
      editor,
      call('insert_content', { afterIndex: -1, markdown: '# Hi\n\nBody.' }),
    )
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(editor.getMarkdown()).toContain('# Hi')
    expect(editor.state.doc.childCount).toBe(2)
  })

  it('inserts after the given block', () => {
    const editor = createEditor('# A\n\nfirst')
    executeTool(editor, call('insert_content', { afterIndex: 0, markdown: 'inserted' }))
    const md = editor.getMarkdown()
    expect(md.indexOf('inserted')).toBeGreaterThan(md.indexOf('# A'))
    expect(md.indexOf('inserted')).toBeLessThan(md.indexOf('first'))
  })

  it('rejects an out-of-range index', () => {
    const editor = createEditor('# A')
    const result = executeTool(editor, call('insert_content', { afterIndex: 9, markdown: 'x' }))
    expect(result.isError).toBe(true)
  })
})

describe('replace_blocks', () => {
  it('rewrites a block range', () => {
    const editor = createEditor('# A\n\nold text\n\nkeep me')
    const result = executeTool(
      editor,
      call('replace_blocks', { startIndex: 1, endIndex: 1, markdown: 'new text' }),
    )
    expect(result.mutated).toBe(true)
    const md = editor.getMarkdown()
    expect(md).toContain('new text')
    expect(md).not.toContain('old text')
    expect(md).toContain('keep me')
  })

  it('deletes a range with empty markdown', () => {
    const editor = createEditor('# A\n\ndelete me\n\nkeep me')
    executeTool(editor, call('replace_blocks', { startIndex: 1, endIndex: 1, markdown: '' }))
    const md = editor.getMarkdown()
    expect(md).not.toContain('delete me')
    expect(md).toContain('keep me')
  })

  it('deleting every block leaves an empty paragraph', () => {
    const editor = createEditor('# A\n\nb')
    executeTool(editor, call('replace_blocks', { startIndex: 0, endIndex: 1, markdown: '' }))
    expect(editor.state.doc.childCount).toBe(1)
  })
})

describe('staleness guard', () => {
  it('refuses index writes after a user edit and recovers via get_document_context', () => {
    const editor = createEditor('# A')
    markDocSeen(editor)
    // simulate a user edit after the AI last saw the doc
    editor.commands.insertContentAt(editor.state.doc.content.size, 'user typed')
    const blocked = executeTool(editor, call('insert_content', { afterIndex: 0, markdown: 'x' }))
    expect(blocked.isError).toBe(true)
    expect(blocked.output).toContain('changed')
    executeTool(editor, call('get_document_context'))
    const ok = executeTool(editor, call('insert_content', { afterIndex: 0, markdown: 'x' }))
    expect(ok.isError).toBeUndefined()
  })
})

describe('read_blocks paging', () => {
  it('pages long output with a continue notice', () => {
    const editor = createEditor(`# T\n\n${'lorem ipsum '.repeat(3000)}`)
    const result = executeTool(editor, call('read_blocks', { startIndex: 0, endIndex: 1 }))
    expect(result.output).toContain('continue with offset=')
    const offset = Number(/offset=(\d+)/.exec(result.output)![1])
    const rest = executeTool(editor, call('read_blocks', { startIndex: 0, endIndex: 1, offset }))
    expect(rest.output.length).toBeGreaterThan(0)
  })
})

describe('deriveAutoFileName', () => {
  it('uses the first heading', () => {
    const editor = createEditor('# 阿里巴巴集团介绍\n\nbody')
    expect(deriveAutoFileName(editor)).toBe('阿里巴巴集团介绍')
  })

  it('falls back to the first words of a paragraph', () => {
    const editor = createEditor('just some plain opening words here to use\n\nmore')
    expect(deriveAutoFileName(editor)).toBe('just some plain opening words here to use')
  })

  it('returns empty for a blank document', () => {
    const editor = createEditor()
    expect(deriveAutoFileName(editor)).toBe('')
  })
})
