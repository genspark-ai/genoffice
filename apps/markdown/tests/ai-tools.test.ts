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

const ops = (...list: Record<string, unknown>[]) => call('apply_ops', { ops: list })
const insert = (afterIndex: number, markdown: string) =>
  ops({ op: 'insertContent', after: afterIndex, markdown })

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

describe('apply_ops insertContent', () => {
  it('replaces the empty paragraph on a blank document', () => {
    const editor = createEditor()
    const result = executeTool(editor, insert(-1, '# Hi\n\nBody.'))
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(editor.getMarkdown()).toContain('# Hi')
    expect(editor.state.doc.childCount).toBe(2)
  })

  it('inserts after the given block', () => {
    const editor = createEditor('# A\n\nfirst')
    executeTool(editor, insert(0, 'inserted'))
    const md = editor.getMarkdown()
    expect(md.indexOf('inserted')).toBeGreaterThan(md.indexOf('# A'))
    expect(md.indexOf('inserted')).toBeLessThan(md.indexOf('first'))
  })

  it('rejects an out-of-range index', () => {
    const editor = createEditor('# A')
    const result = executeTool(editor, insert(9, 'x'))
    expect(result.isError).toBe(true)
    expect(result.mutated).toBeFalsy()
  })

  it('rejects malformed ops before touching the document', () => {
    const editor = createEditor('# A')
    const unknownOp = executeTool(editor, ops({ op: 'explode', target: 'selection' }))
    expect(unknownOp.isError).toBe(true)
    expect(unknownOp.output).toContain('unknown op')
    const extraField = executeTool(editor, ops({ op: 'deleteBlocks', target: 'selection', x: 1 }))
    expect(extraField.output).toContain('unknown field')
    const badTarget = executeTool(editor, ops({ op: 'deleteBlocks', target: { start: -1 } }))
    expect(badTarget.isError).toBe(true)
    expect(editor.getMarkdown()).toContain('# A')
  })
})

describe('apply_ops batches', () => {
  it('indexes refer to the document before the call, whatever the op order', () => {
    const editor = createEditor('# A\n\nb\n\nc')
    const result = executeTool(
      editor,
      ops(
        { op: 'insertContent', after: -1, markdown: 'intro' },
        { op: 'replaceBlocks', target: { start: 2 }, markdown: 'C!' },
        { op: 'setBlockType', target: { start: 1 }, type: 'heading', level: 2 },
      ),
    )
    expect(result.isError).toBeUndefined()
    const texts: string[] = []
    editor.state.doc.forEach((n) => texts.push(`${n.type.name}:${n.textContent}`))
    expect(texts).toEqual(['paragraph:intro', 'heading:A', 'heading:b', 'paragraph:C!'])
    expect(result.output).toContain('ops[2] setBlockType')
    expect(result.output).toContain('Block indexes may have changed')
  })

  it('stops at the first failing op and reports what ran', () => {
    const editor = createEditor('one\n\ntwo')
    const result = executeTool(
      editor,
      ops(
        { op: 'replaceText', target: { start: 0 }, find: 'one', replace: '1' },
        { op: 'replaceText', target: { start: 1 }, find: 'missing', replace: 'x' },
        { op: 'replaceText', target: { start: 1 }, find: 'two', replace: '2' },
      ),
    )
    expect(result.isError).toBe(true)
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('ops[0] replaceText: Replaced 1')
    expect(result.output).toContain('ops[1] replaceText FAILED')
    expect(result.output).toContain('1 later op(s) were not executed')
    expect(editor.getMarkdown()).toContain('two')
  })

  it('a deleted target is reported instead of hitting a neighbour', () => {
    const editor = createEditor('a\n\nb\n\nc')
    const result = executeTool(
      editor,
      ops(
        { op: 'deleteBlocks', target: { start: 1 } },
        { op: 'replaceText', target: { start: 1 }, find: 'b', replace: 'x' },
      ),
    )
    expect(result.output).toContain('removed by an earlier op')
    expect(editor.getMarkdown()).toContain('c')
  })
})

describe('model output is sanitized to pure GFM', () => {
  it('raw HTML in tool input degrades to plain text', () => {
    const editor = createEditor()
    executeTool(
      editor,
      insert(-1, '<p style="text-align: center"><span style="color: red">note</span> here</p>'),
    )
    const md = editor.getMarkdown()
    expect(md).toContain('note here')
    expect(md).not.toContain('<')
  })

  it('legacy ::: fenced divs in tool input are stripped, keeping the body', () => {
    const editor = createEditor()
    executeTool(editor, insert(-1, ':::callout {type="warning"}\nBe careful.\n:::'))
    const md = editor.getMarkdown()
    expect(md).toContain('Be careful.')
    expect(md).not.toContain(':::')
  })
})

describe('apply_ops replaceBlocks', () => {
  it('rewrites a block range', () => {
    const editor = createEditor('# A\n\nold text\n\nkeep me')
    const result = executeTool(
      editor,
      ops({ op: 'replaceBlocks', target: { start: 1, end: 1 }, markdown: 'new text' }),
    )
    expect(result.mutated).toBe(true)
    const md = editor.getMarkdown()
    expect(md).toContain('new text')
    expect(md).not.toContain('old text')
    expect(md).toContain('keep me')
  })

  it('deletes a range with empty markdown', () => {
    const editor = createEditor('# A\n\ndelete me\n\nkeep me')
    executeTool(editor, ops({ op: 'replaceBlocks', target: { start: 1 }, markdown: '' }))
    const md = editor.getMarkdown()
    expect(md).not.toContain('delete me')
    expect(md).toContain('keep me')
  })

  it('deleting every block leaves an empty paragraph', () => {
    const editor = createEditor('# A\n\nb')
    executeTool(editor, ops({ op: 'deleteBlocks', target: { start: 0, end: 1 } }))
    expect(editor.state.doc.childCount).toBe(1)
  })
})

describe('staleness guard', () => {
  it('refuses index writes after a user edit and recovers via get_document_context', () => {
    const editor = createEditor('# A')
    markDocSeen(editor)
    // simulate a user edit after the AI last saw the doc
    editor.commands.insertContentAt(editor.state.doc.content.size, 'user typed')
    const blocked = executeTool(editor, insert(0, 'x'))
    expect(blocked.isError).toBe(true)
    expect(blocked.output).toContain('changed')
    // selection-addressed ops never go stale — they read the live document
    const sel = executeTool(editor, ops({ op: 'setStyle', target: 'selection', style: 'bold' }))
    expect(sel.isError).toBeUndefined()
    executeTool(editor, call('get_document_context'))
    const ok = executeTool(editor, insert(0, 'x'))
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

describe('selection context', () => {
  it('names the covered block range', () => {
    const editor = createEditor('# A\n\nfirst para\n\nsecond para')
    editor.commands.setTextSelection({ from: 6, to: editor.state.doc.content.size - 2 })
    const ctx = buildDocContext(editor)
    expect(ctx).toMatch(/## User selection \(blocks 1-2\)/)
  })
})

describe('math markdown', () => {
  it('parses $...$ into math nodes and round-trips', () => {
    const editor = createEditor()
    executeTool(editor, insert(-1, 'Energy: $E=mc^2$'))
    let mathNodes = 0
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'inlineMath') mathNodes++
    })
    expect(mathNodes).toBe(1)
    expect(editor.getMarkdown()).toContain('$E=mc^2$')
  })
})

const replaceText = (blockIndex: number, find: string, replace: string) =>
  ops({ op: 'replaceText', target: { start: blockIndex }, find, replace })

describe('apply_ops replaceText', () => {
  it('replaces every occurrence in one block and keeps surrounding marks', () => {
    const editor = createEditor('# A\n\nThe **TODO** item and another TODO here.')
    const result = executeTool(editor, replaceText(1, 'TODO', 'DONE')) as {
      isError?: boolean
      output: string
    }
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('2 occurrence(s)')
    const md = editor.getMarkdown()
    expect(md).toContain('**DONE**')
    expect(md).toContain('another DONE here')
    expect(md).not.toContain('TODO')
  })

  it('deletes when replace is empty', () => {
    const editor = createEditor('alpha beta gamma')
    executeTool(editor, replaceText(0, ' beta', ''))
    expect(editor.getMarkdown()).toContain('alpha gamma')
  })

  it('only touches the addressed block', () => {
    const editor = createEditor('same text\n\nsame text')
    executeTool(editor, replaceText(1, 'same', 'other'))
    const md = editor.getMarkdown()
    expect(md.indexOf('same text')).toBeLessThan(md.indexOf('other text'))
  })

  it('does not match across list-item boundaries', () => {
    const editor = createEditor('- one\n- two')
    const result = executeTool(editor, replaceText(0, 'one\ntwo', 'x')) as { isError?: boolean }
    expect(result.isError).toBe(true)
  })

  it('reports not-found with guidance', () => {
    const editor = createEditor('hello world')
    const result = executeTool(editor, replaceText(0, 'absent', 'x')) as {
      isError?: boolean
      output: string
    }
    expect(result.isError).toBe(true)
    expect(result.output).toContain('not found')
  })
})

describe('apply_ops setStyle', () => {
  it('bolds every match in the range', () => {
    const editor = createEditor('a TODO here\n\nanother TODO there')
    const result = executeTool(
      editor,
      ops({ op: 'setStyle', target: { start: 0, end: 1 }, find: 'TODO', style: 'bold' }),
    ) as { isError?: boolean; output: string }
    expect(result.isError).toBeUndefined()
    const md = editor.getMarkdown()
    expect((md.match(/\*\*TODO\*\*/g) ?? []).length).toBe(2)
  })

  it('removes a style with remove: true', () => {
    const editor = createEditor('a **TODO** here')
    executeTool(
      editor,
      ops({ op: 'setStyle', target: { start: 0 }, find: 'TODO', style: 'bold', mode: 'remove' }),
    )
    expect(editor.getMarkdown()).not.toContain('**')
  })

  it('rejects unknown styles', () => {
    const editor = createEditor('text')
    const result = executeTool(
      editor,
      ops({ op: 'setStyle', target: { start: 0 }, find: 'text', style: 'underline' }),
    ) as { isError?: boolean }
    expect(result.isError).toBe(true)
  })

  it('styles the whole target when find is omitted', () => {
    const editor = createEditor('plain words')
    executeTool(editor, ops({ op: 'setStyle', target: { start: 0 }, style: 'italic' }))
    expect(editor.getMarkdown()).toContain('*plain words*')
  })
})

describe('insert_image', () => {
  const PNG =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
  const withApi = (api: Record<string, unknown>) => {
    ;(window as unknown as { markdownApi: unknown }).markdownApi = api
  }

  it('errors cleanly when the document has never been saved', async () => {
    withApi({
      fetchImage: async () => ({ base64: PNG, mime: 'image/png' }),
      saveImage: async () => null,
    })
    const editor = createEditor('# A')
    const result = await executeTool(
      editor,
      call('insert_image', { url: 'https://example.com/x.png' }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('save the document first')
  })

  it('inserts a saved relative path as an image block', async () => {
    withApi({
      fetchImage: async () => ({ base64: PNG, mime: 'image/png' }),
      saveImage: async () => 'assets/pic.png',
    })
    const editor = createEditor('# A\n\npara')
    const result = await executeTool(
      editor,
      call('insert_image', { url: 'https://example.com/x.png', afterIndex: 0, alt: 'chart' }),
    )
    expect(result.isError).toBeUndefined()
    expect(result.mutated).toBe(true)
    expect(editor.getMarkdown()).toContain('![chart](assets/pic.png)')
  })
})

describe('blank/selection edge cases (Bugbot #871)', () => {
  it('an image-only document is not blank: context lists it, inserts append', () => {
    const editor = createEditor('![pic](assets/pic.png)')
    expect(buildDocContext(editor)).not.toContain('currently blank')
    executeTool(editor, insert(-1, 'caption'))
    const md = editor.getMarkdown()
    expect(md).toContain('![pic](assets/pic.png)')
    expect(md).toContain('caption')
  })

  it('a node selection with no text still reports the selected block', () => {
    const editor = createEditor('intro\n\n![pic](assets/pic.png)')
    let imagePos = -1
    editor.state.doc.forEach((node, offset) => {
      if (node.type.name === 'image') imagePos = offset
    })
    editor.commands.setNodeSelection(imagePos)
    const ctx = buildDocContext(editor)
    expect(ctx).toContain('## User selection (block 1)')
    expect(ctx).toContain('non-text block is selected: image')
  })
})

describe('review follow-ups (#871)', () => {
  it('a queued anchor over an image resolves with a type placeholder, not orphaned', async () => {
    const { addQueueAnchor } = await import('../src/renderer/editor/aiQueueAnchors')
    const { resolveQueueItem } = await import('../src/renderer/ai/edit-queue')
    const editor = createEditor('intro\n\n![pic](assets/pic.png)')
    let imagePos = -1
    editor.state.doc.forEach((node, offset) => {
      if (node.type.name === 'image') imagePos = offset
    })
    addQueueAnchor(editor, 'q1', imagePos, imagePos + 1)
    const r = resolveQueueItem(editor, { qid: 'q1', instruction: 'replace it', capturedText: '' })
    expect(r.target).not.toBeNull()
    expect(r.target!.startIndex).toBe(1)
    expect(r.target!.excerpt).toContain('image')
  })

  it('a deleted anchor still resolves to orphaned', async () => {
    const { addQueueAnchor } = await import('../src/renderer/editor/aiQueueAnchors')
    const { resolveQueueItem } = await import('../src/renderer/ai/edit-queue')
    const editor = createEditor('intro text here')
    addQueueAnchor(editor, 'q2', 1, 6)
    editor.view.dispatch(editor.state.tr.delete(1, 6))
    const r = resolveQueueItem(editor, { qid: 'q2', instruction: 'x', capturedText: 'intro' })
    expect(r.target).toBeNull()
  })

  it('rejects a webp download despite the jpeg content-type fallback', async () => {
    const webp = Buffer.from('RIFF\0\0\0\0WEBPVP8 ', 'binary').toString('base64')
    ;(window as unknown as { markdownApi: unknown }).markdownApi = {
      fetchImage: async () => ({ base64: webp, mime: 'image/jpeg' }),
      saveImage: async () => 'assets/x.jpg',
    }
    const editor = createEditor('# A')
    const result = await executeTool(
      editor,
      call('insert_image', { url: 'https://example.com/x.webp' }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('unsupported image format')
  })

  it('afterIndex null falls back to the end of the document', async () => {
    const PNG =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
    ;(window as unknown as { markdownApi: unknown }).markdownApi = {
      fetchImage: async () => ({ base64: PNG, mime: 'image/png' }),
      saveImage: async () => 'assets/pic.png',
    }
    const editor = createEditor('# A\n\npara')
    await executeTool(
      editor,
      call('insert_image', { url: 'https://x.com/a.png', afterIndex: null }),
    )
    expect(editor.getMarkdown().trimEnd().endsWith('![](assets/pic.png)')).toBe(true)
  })
})

describe('selectionForAnchor', () => {
  it('selects the node for an anchored image and text for an anchored passage', async () => {
    const { addQueueAnchor } = await import('../src/renderer/editor/aiQueueAnchors')
    const { selectionForAnchor } = await import('../src/renderer/ai/edit-queue')
    const { NodeSelection, TextSelection } = await import('@tiptap/pm/state')
    const editor = createEditor('intro text\n\n![pic](assets/pic.png)')
    let imagePos = -1
    editor.state.doc.forEach((node, offset) => {
      if (node.type.name === 'image') imagePos = offset
    })
    addQueueAnchor(editor, 'img', imagePos, imagePos + 1)
    addQueueAnchor(editor, 'txt', 1, 6)
    const imgSel = selectionForAnchor(editor, 'img')
    expect(imgSel).toBeInstanceOf(NodeSelection)
    // dispatching the focus selection must not throw on a block atom
    editor.view.dispatch(editor.state.tr.setSelection(imgSel!))
    const txtSel = selectionForAnchor(editor, 'txt')
    expect(txtSel).toBeInstanceOf(TextSelection)
    expect(selectionForAnchor(editor, 'missing')).toBeNull()
  })
})

describe('frontmatter tools', () => {
  const fmStore = (initial = '') => {
    let inner = initial
    return {
      read: () => inner,
      write: (v: string) => {
        inner = v
      },
    }
  }

  it('reads back what was just written within the same run', () => {
    const editor = createEditor('# A')
    const fm = fmStore()
    const empty = executeTool(editor, call('read_frontmatter'), undefined, fm) as {
      output: string
    }
    expect(empty.output).toContain('no frontmatter')
    const set = executeTool(
      editor,
      ops({ op: 'setFrontmatter', yaml: 'title: Hello\ntags: [a, b]\n' }),
      undefined,
      fm,
    ) as { mutated?: boolean }
    expect(set.mutated).toBe(true)
    const read = executeTool(editor, call('read_frontmatter'), undefined, fm) as {
      output: string
    }
    expect(read.output).toBe('title: Hello\ntags: [a, b]')
  })

  it('an empty yaml removes the block', () => {
    const editor = createEditor('# A')
    const fm = fmStore('title: Old')
    executeTool(editor, ops({ op: 'setFrontmatter', yaml: '  \n' }), undefined, fm)
    expect(fm.read()).toBe('')
  })

  it('fails cleanly without frontmatter access', () => {
    const editor = createEditor('# A')
    const result = executeTool(editor, call('read_frontmatter')) as { isError?: boolean }
    expect(result.isError).toBe(true)
  })
})
