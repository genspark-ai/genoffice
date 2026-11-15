// Modified for ForaOffice from GenOffice v0.12.852.
import { afterEach, describe, expect, it } from 'vitest'
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
  extractFragment,
  type AiDocWriter,
} from '../src/renderer/ai/doc-writer'
import { DraftLanding } from '../src/renderer/runtime/draft-landing'
import { DocumentDraft } from '../src/renderer/runtime/document-draft'
import { createDocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { createNativeDocsWorkflows } from '../src/renderer/ai/native-workflows'
import { HTML_RULES } from '../src/renderer/ai/protocol'
import type { AgentToolCall } from '@genoffice/agent-core'

const NUM_IDS = { bullet: BLANK_BULLET_NUM_ID, ordered: BLANK_ORDERED_NUM_ID }

/** Editors created during the current test; destroyed in afterEach so ProseMirror's
 * DOMObserver timers can't fire after the JSDOM environment is torn down. */
const liveEditors: Editor[] = []
const runtimes: ReturnType<typeof createDocsRuntime>[] = []

async function executeWrite(
  editor: Editor,
  call: AgentToolCall,
  writer: AiDocWriter,
  signal?: AbortSignal,
) {
  const runtime = createDocsRuntime({
    current: () => ({ editor, identity: editor, numIds: NUM_IDS }),
  })
  runtimes.push(runtime)
  const client = runtime.connect({ id: crypto.randomUUID() })
  client.context()
  return createNativeDocsWorkflows(client, () => writer).execute(call, signal)
}

afterEach(() => {
  for (const runtime of runtimes.splice(0)) runtime.dispose()
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

async function createBlankEditor() {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const parsed = await parseDocx(await buildBlankDocx())
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  liveEditors.push(editor)
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
    expect(system).toContain(HTML_RULES)
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
  it('preserves a manually edited preview block and removes only untouched draft content', async () => {
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
    expect(texts(editor)).toEqual(['one', 'alpha edited', 'beta'])
    draft.finish()
    expect(texts(editor)).toEqual(['one', 'alpha edited'])
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

describe('DocumentDraft read-only metadata', () => {
  it('rejects an update before changing the accepted sequence, content, or preview', async () => {
    const editor = await createBlankEditor()
    const draft = new DocumentDraft(editor, NUM_IDS, { kind: 'whole' })
    const accepted = '<p>Accepted content</p>'
    draft.update(1, accepted)
    await tick()
    const before = draft.inspect()
    const preview = editor.state.doc

    expect(() =>
      draft.update(2, '<p>Replacement prefix</p><p data-doc-readonly="field">Field result</p>'),
    ).toThrow(/read-only document metadata/i)
    expect(draft.inspect()).toEqual(before)
    expect(editor.state.doc.eq(preview)).toBe(true)
    await tick()
    expect(editor.state.doc.eq(preview)).toBe(true)
    expect(() => draft.update(1, accepted)).not.toThrow()

    draft.stop()
    expect(draft.finish(true)).toEqual({ ok: true, blocks: 1, partial: true })
    expect(texts(editor)).toEqual(['Accepted content'])
  })

  it.each([false, true])(
    'cannot commit the prefix of a rejected first update (partial=%s)',
    async (partial) => {
      const editor = await createBlankEditor()
      const original = editor.state.doc
      const draft = new DocumentDraft(editor, NUM_IDS, { kind: 'whole' })

      expect(() =>
        draft.update(
          1,
          '<p>Unaccepted prefix</p><span data-doc-readonly="protected">Object</span>',
        ),
      ).toThrow(/read-only document metadata/i)
      expect(draft.inspect()).toMatchObject({ sequence: 0, characters: 0 })
      draft.stop()
      expect(draft.finish(partial)).toMatchObject({ ok: false, errorCode: 'invalid_input' })
      expect(editor.state.doc.eq(original)).toBe(true)
    },
  )
})

describe('write_document tool', () => {
  const writerOf = (html: string, truncated = false): AiDocWriter => ({
    write: async (_spec, onProgress) => {
      onProgress.onProgress(html.slice(0, 10))
      onProgress.onProgress(html)
      if (truncated) await onProgress.onStopped()
      return { ok: true, html, truncated }
    },
  })

  it('lands the streamed fragment as one regular insert on a blank document', async () => {
    const editor = await createBlankEditor()
    const exec = await executeWrite(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'a title and a paragraph' } },
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
    const exec = await executeWrite(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p' } },
      writerOf('<p>new</p>'),
    )
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('afterBlockIndex')
    expect(texts(editor)).toEqual(['keep'])
  })

  it('inserts after a block and flags a kept partial as incomplete', async () => {
    const editor = await createBlankEditor()
    editor.commands.setContent('<p>one</p><p>two</p>')
    const exec = await executeWrite(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p', afterBlockIndex: 0 } },
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
        onProgress.onProgress('<p>ok</p>')
        await tick()
        onProgress.onProgress('<p>ok</p><formula>\\frac{</formula>')
        return { ok: true, html: '<p>ok</p><formula>\\frac{</formula>', truncated: true }
      },
    }
    const exec = await executeWrite(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p' } },
      writer,
    )
    expect(exec.isError).toBeFalsy()
    expect(texts(editor)).toEqual(['ok'])
  })

  it('refuses a native partial containing read-only metadata instead of committing its preview', async () => {
    const editor = await createBlankEditor()
    const original = editor.state.doc
    const html = '<p>Prefix</p><p data-doc-readonly="field">Field result</p>'
    const exec = await executeWrite(
      editor,
      { id: 'readonly', name: 'write_document', input: { plan: 'p' } },
      {
        write: async (_spec, onProgress) => {
          onProgress.onProgress('<p>Prefix</p>')
          await tick()
          onProgress.onProgress(html)
          return { ok: true, html, truncated: true }
        },
      },
    )
    expect(exec).toMatchObject({ isError: true, committed: false })
    expect(exec.output).toMatch(/read-only document metadata/i)
    expect(editor.state.doc.eq(original)).toBe(true)
  })

  it('keeps text the user typed into a blank document while the draft streamed', async () => {
    const editor = await createBlankEditor()
    const writer: AiDocWriter = {
      write: async (_spec, onProgress) => {
        onProgress.onProgress('<p>draft</p>')
        await tick()
        // the user types below the draft meanwhile
        editor.view.dispatch(editor.state.tr.insertText('mine', editor.state.doc.content.size - 1))
        return { ok: true, html: '<p>draft</p><p>more</p>' }
      },
    }
    const exec = await executeWrite(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p' } },
      writer,
    )
    expect(exec.isError).toBeFalsy()
    expect(texts(editor)).toEqual(['draft', 'more', 'mine'])
  })

  it('a failed or discarded write leaves the document unchanged', async () => {
    const editor = await createBlankEditor()
    const writer: AiDocWriter = {
      write: async (_spec, onProgress) => {
        onProgress.onProgress('<p>draft</p>')
        await tick()
        return { ok: false, error: 'the user discarded the partial content' }
      },
    }
    const exec = await executeWrite(
      editor,
      { id: 't', name: 'write_document', input: { plan: 'p' } },
      writer,
    )
    expect(exec.isError).toBe(true)
    expect(texts(editor)).toEqual([''])
  })

  it.each([true, false])(
    'settles a user-stopped native writer through the shared draft decision (keep=%s)',
    async (keep) => {
      const editor = await createBlankEditor()
      const controller = new AbortController()
      const result = await executeWrite(
        editor,
        { id: 'partial', name: 'write_document', input: { plan: 'Write' } },
        {
          write: async (_spec, observer, signal) => {
            observer.onProgress('<p>Partial text</p>')
            controller.abort()
            expect(signal?.aborted).toBe(true)
            await observer.onStopped()
            await tick()
            expect(texts(editor)).toContain('Partial text')
            return keep
              ? { ok: true, html: '<p>Partial text</p>', truncated: true }
              : { ok: false, error: 'The user discarded the partial content' }
          },
        },
        controller.signal,
      )
      expect(result.isError).toBe(keep ? undefined : true)
      expect(texts(editor)).toEqual(keep ? ['Partial text'] : [''])
      if (keep) {
        expect(editor.commands.undo()).toBe(true)
        expect(texts(editor)).toEqual([''])
      }
    },
  )

  it('refuses a late stop decision after document disposal', async () => {
    const editor = await createBlankEditor()
    const result = await executeWrite(
      editor,
      { id: 'closed', name: 'write_document', input: { plan: 'Write' } },
      {
        write: async (_spec, observer) => {
          observer.onProgress('<p>Preview</p>')
          await tick()
          runtimes.at(-1)!.dispose()
          await expect(observer.onStopped()).rejects.toThrow()
          return { ok: true, html: '<p>Late text</p>', truncated: true }
        },
      },
    )
    expect(result).toMatchObject({ isError: true, errorCode: 'canceled', committed: false })
    expect(texts(editor)).toEqual([''])
  })
})
