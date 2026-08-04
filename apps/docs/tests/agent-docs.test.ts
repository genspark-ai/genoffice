import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { AgentLoop, type AgentStreamCallbacks, type AgentTransport } from '@genoffice/agent-core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createDocsSkill } from '../src/renderer/ai/docs-skill'
import { buildDocContext, countWords } from '../src/renderer/ai/protocol'
import { executeTool } from '../src/renderer/ai/tools'

/**
 * End-to-end through the local stack: AgentLoop -> docs skill -> tools ->
 * real tiptap editor. The "model" is scripted, so these tests verify that
 * the capabilities behind typical requests (word count, heading colors,
 * rewrite, insert) actually work when the model picks the right tool.
 */

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const heading = (t: string, level = 1): JsonNode => ({
  type: 'docHeading',
  attrs: { docxIndex: null, level },
  content: [text(t)],
})
const para = (t: string): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [text(t)],
})

/** Editors created during the current test; destroyed in afterEach so ProseMirror's
 * DOMObserver timers can't fire after the JSDOM environment is torn down. */
const liveEditors: Editor[] = []

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  liveEditors.push(editor)
  return editor
}

afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

/** 0 h1 | 1 p | 2 h2 | 3 p */
const fixture = () => [
  heading('Chapter 1 Overview', 1),
  para('GenSpark is an AI office suite.'),
  heading('Risk Notes', 2),
  para('This document is for reference only.'),
]

function scriptedTransport(script: Array<(cb: AgentStreamCallbacks) => void>): AgentTransport {
  let turn = 0
  return {
    stream(_request, cb) {
      const step = script[turn++]
      if (step) queueMicrotask(() => step(cb))
      return { cancel: () => queueMicrotask(() => cb.onDone()) }
    },
  }
}

const flush = () => new Promise((r) => setTimeout(r, 0))
const NUM_IDS = { bullet: null, ordered: null }

function makeLoop(editor: Editor, transport: AgentTransport, onDone: (text: string) => void) {
  return new AgentLoop({
    transport,
    skill: createDocsSkill(
      () => editor,
      () => NUM_IDS,
    ),
    events: { onDone: (r) => onDone(r.text) },
  })
}

describe('word-count stats (answer-style requests)', () => {
  it('each turn context carries full-text stats matching the status bar so the model can quote them directly', async () => {
    const editor = createEditor(fixture())
    const context = buildDocContext(editor)
    const expected = countWords(editor.state.doc.textContent)
    expect(context).toContain(`Full-text stats: words ${expected}`)
    expect(context).toContain('0|h1|Chapter 1 Overview')
    expect(context).toContain('3|p|This document is for reference only.')
  })

  it('get_document_context tool returns the same stats and does not modify the document', async () => {
    const editor = createEditor(fixture())
    const before = JSON.stringify(editor.getJSON())
    const exec = await executeTool(
      editor,
      { id: 't1', name: 'get_document_context', input: {} },
      NUM_IDS,
    )
    expect(exec.mutated).toBe(false)
    expect(exec.output).toContain(`words ${countWords(editor.state.doc.textContent)}`)
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })

  it('answer-only turn: the model replies with plain text and the document is unchanged', async () => {
    const editor = createEditor(fixture())
    const before = JSON.stringify(editor.getJSON())
    let final = ''
    const loop = makeLoop(
      editor,
      scriptedTransport([
        (cb) => {
          cb.onDelta('The document has 17 words.')
          cb.onDone()
        },
      ]),
      (t) => (final = t),
    )
    loop.run('How many words are there now')
    await flush()
    expect(final).toBe('The document has 17 words.')
    expect(JSON.stringify(editor.getJSON())).toBe(before)
    // The context made it into the user message sent to the model
    expect((loop.messages[0] as { text: string }).text).toContain('Full-text stats')
  })
})

describe('changing heading colors (formatting-command requests)', () => {
  it('after the model calls apply_commands, all headings turn red with the aiChanged highlight', async () => {
    const editor = createEditor(fixture())
    let final = ''
    const loop = makeLoop(
      editor,
      scriptedTransport([
        (cb) => {
          cb.onToolCall({
            id: 't1',
            name: 'apply_commands',
            input: {
              commands: [
                {
                  updateTextStyle: {
                    target: { nodeType: 'docHeading' },
                    style: { color: 'FF0000' },
                    fields: ['color'],
                  },
                },
              ],
            },
          })
          cb.onDone()
        },
        (cb) => {
          cb.onDelta('All headings are now red.')
          cb.onDone()
        },
      ]),
      (t) => (final = t),
    )
    loop.run('Turn all headings red')
    await flush()
    await flush()

    expect(final).toBe('All headings are now red.')
    for (const blockIndex of [0, 2]) {
      const block = editor.state.doc.child(blockIndex)
      expect(block.attrs.aiChanged).toBe(true)
      const mark = block.child(0).marks.find((m) => m.type.name === 'docTextStyle')
      expect(mark?.attrs.color).toBe('FF0000')
    }
    // Body text is unaffected
    expect(
      editor.state.doc
        .child(1)
        .child(0)
        .marks.some((m) => m.type.name === 'docTextStyle'),
    ).toBe(false)
    // The tool result was passed back to the model
    const toolMsg = loop.messages[2] as { role: 'tool'; results: Array<{ output: string }> }
    expect(toolMsg.results[0].output).toContain('已更新 2 个块的文字样式')
  })

  it('invalid commands return an error result, the document is unchanged, and the loop continues', async () => {
    const editor = createEditor(fixture())
    const before = JSON.stringify(editor.getJSON())
    const loop = makeLoop(
      editor,
      scriptedTransport([
        (cb) => {
          cb.onToolCall({
            id: 't1',
            name: 'apply_commands',
            input: { commands: [{ updateTextStyle: { style: {}, fields: [] } }] },
          })
          cb.onDone()
        },
        (cb) => {
          cb.onDelta('Invalid command.')
          cb.onDone()
        },
      ]),
      () => {},
    )
    loop.run('Make random edits')
    await flush()
    await flush()
    const toolMsg = loop.messages[2] as { role: 'tool'; results: Array<{ isError?: boolean }> }
    expect(toolMsg.results[0].isError).toBe(true)
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })
})

describe('content read/write tools', () => {
  it('read_blocks returns the full restricted HTML', async () => {
    const editor = createEditor(fixture())
    const exec = await executeTool(
      editor,
      { id: 't', name: 'read_blocks', input: { startBlockIndex: 0, endBlockIndex: 1 } },
      NUM_IDS,
    )
    expect(exec.output).toContain('<h1>Chapter 1 Overview</h1>')
    expect(exec.output).toContain('<p>GenSpark is an AI office suite.</p>')
  })

  it('read_blocks pages oversized content: offset continuation reassembles the full HTML', async () => {
    // one paragraph well past the 24k read cap
    const long = 'A'.repeat(30_000)
    const editor = createEditor([heading('Long chapter', 1), para(long)])
    const first = await executeTool(
      editor,
      { id: 't1', name: 'read_blocks', input: { startBlockIndex: 0, endBlockIndex: 1 } },
      NUM_IDS,
    )
    expect(first.isError).toBeUndefined()
    const match = first.output.match(/offset=(\d+)/)
    expect(first.output).toContain('truncated')
    expect(match).not.toBeNull()
    const offset = Number(match![1])
    const second = await executeTool(
      editor,
      {
        id: 't2',
        name: 'read_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 1, offset },
      },
      NUM_IDS,
    )
    expect(second.isError).toBeUndefined()
    expect(second.output).toContain('end of range')
    const stitched =
      first.output.slice(0, first.output.lastIndexOf('\n…(truncated')) +
      second.output.slice(0, second.output.lastIndexOf('\n(end of range'))
    expect(stitched).toContain('<h1>Long chapter</h1>')
    expect(stitched).toContain(long)
    // an offset beyond the content is an explicit error, not an empty read
    const beyond = await executeTool(
      editor,
      {
        id: 't3',
        name: 'read_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 1, offset: 10_000_000 },
      },
      NUM_IDS,
    )
    expect(beyond.isError).toBe(true)
  })

  it('replace_blocks rewrites the specified blocks', async () => {
    const editor = createEditor(fixture())
    const exec = await executeTool(
      editor,
      {
        id: 't',
        name: 'replace_blocks',
        input: { startBlockIndex: 1, endBlockIndex: 1, html: '<p>Intro rewritten.</p>' },
      },
      NUM_IDS,
    )
    expect(exec.mutated).toBe(true)
    const block = editor.state.doc.child(1)
    expect(block.textContent).toBe('Intro rewritten.')
    expect(block.attrs.aiChanged).toBe(true)
    expect(editor.state.doc.childCount).toBe(4)
  })

  it('insert_content inserts after the specified block', async () => {
    const editor = createEditor(fixture())
    const exec = await executeTool(
      editor,
      {
        id: 't',
        name: 'insert_content',
        input: { html: '<h2>New Section</h2><p>New content.</p>', afterBlockIndex: 3 },
      },
      NUM_IDS,
    )
    expect(exec.mutated).toBe(true)
    expect(editor.state.doc.childCount).toBe(6)
    expect(editor.state.doc.child(4).textContent).toBe('New Section')
    expect(editor.state.doc.child(5).textContent).toBe('New content.')
  })

  it('an out-of-range end index is an error, not silently clamped', async () => {
    const editor = createEditor(fixture())
    const before = JSON.stringify(editor.getJSON())
    for (const name of ['read_blocks', 'replace_blocks']) {
      const exec = await executeTool(
        editor,
        { id: 't', name, input: { startBlockIndex: 2, endBlockIndex: 9, html: '<p>x</p>' } },
        NUM_IDS,
      )
      expect(exec.isError).toBe(true)
      expect(exec.output).toContain('4 blocks')
    }
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })
})

describe('external-edit guard (document freshness baseline)', () => {
  const read = (editor: Editor) =>
    executeTool(editor, { id: 'r', name: 'get_document_context', input: {} }, NUM_IDS)
  const replace = (editor: Editor) =>
    executeTool(
      editor,
      {
        id: 'w',
        name: 'replace_blocks',
        input: { startBlockIndex: 1, endBlockIndex: 1, html: '<p>rewritten</p>' },
      },
      NUM_IDS,
    )

  it('index-addressed writes fail after a user edit, and succeed again after a re-read', async () => {
    const editor = createEditor(fixture())
    await read(editor)
    // simulate a manual user edit between tool calls
    editor.view.dispatch(editor.state.tr.insertText('typed by user ', 2))
    const stale = await executeTool(
      editor,
      {
        id: 'w',
        name: 'apply_commands',
        input: { commands: [{ deleteBlocks: { target: { blockIndexes: [3] } } }] },
      },
      NUM_IDS,
    )
    expect(stale.isError).toBe(true)
    expect(stale.output).toContain('edited by the user')
    expect(editor.state.doc.childCount).toBe(4) // nothing deleted
    await executeTool(
      editor,
      { id: 'r2', name: 'read_blocks', input: { startBlockIndex: 0, endBlockIndex: 3 } },
      NUM_IDS,
    )
    const retry = await replace(editor)
    expect(retry.isError).toBeUndefined()
    expect(editor.state.doc.child(1).textContent).toBe('rewritten')
  })

  it("the AI's own consecutive writes do not trip the guard", async () => {
    const editor = createEditor(fixture())
    await read(editor)
    expect((await replace(editor)).isError).toBeUndefined()
    const second = await executeTool(
      editor,
      { id: 'w2', name: 'insert_content', input: { html: '<p>appendix</p>', afterBlockIndex: 3 } },
      NUM_IDS,
    )
    expect(second.isError).toBeUndefined()
    expect(editor.state.doc.childCount).toBe(5)
  })
})

describe('blank-document detection', () => {
  it('an image-only document is not blank: insert_content appends instead of wiping it', async () => {
    const editor = createEditor([
      {
        type: 'docProtected',
        attrs: { docxIndex: null, blockType: 'image', label: 'Image' },
      },
    ])
    const exec = await executeTool(
      editor,
      { id: 't', name: 'insert_content', input: { html: '<p>caption</p>', afterBlockIndex: 0 } },
      NUM_IDS,
    )
    expect(exec.isError).toBeUndefined()
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(0).attrs.blockType).toBe('image')
    expect(buildDocContext(editor)).not.toContain('blank')
  })

  it('a single empty paragraph is still blank: insert_content replaces the template paragraph', async () => {
    const editor = createEditor([{ type: 'docParagraph', attrs: { docxIndex: null } }])
    expect(buildDocContext(editor)).toContain('blank')
    const exec = await executeTool(
      editor,
      { id: 't', name: 'insert_content', input: { html: '<p>hello</p>' } },
      NUM_IDS,
    )
    expect(exec.isError).toBeUndefined()
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.child(0).textContent).toBe('hello')
  })
})

describe('web_search backend failures', () => {
  it("method 'error' surfaces as a tool error instead of '(no results)'", async () => {
    const editor = createEditor(fixture())
    const w = window as unknown as { desktop?: unknown }
    const saved = w.desktop
    w.desktop = {
      webSearch: async () => ({ results: [], method: 'error', error: 'Serper 502' }),
    }
    try {
      const exec = await executeTool(
        editor,
        { id: 't', name: 'web_search', input: { query: 'genspark' } },
        NUM_IDS,
      )
      expect(exec.isError).toBe(true)
      expect(exec.output).toContain('Serper 502')
      expect(exec.output).not.toContain('(no results)')
    } finally {
      w.desktop = saved
    }
  })
})

describe('insert_image freshness baseline', () => {
  /** jsdom never decodes images; fake one that reports a fixed natural size */
  class FakeImage {
    onload: (() => void) | null = null
    onerror: (() => void) | null = null
    naturalWidth = 100
    naturalHeight = 80
    set src(_v: string) {
      queueMicrotask(() => this.onload?.())
    }
  }

  const withImageStubs = async (fn: (release: () => void) => Promise<void>) => {
    const w = window as unknown as { desktop?: unknown }
    const savedDesktop = w.desktop
    const savedImage = globalThis.Image
    let release!: () => void
    w.desktop = {
      fetchImage: () =>
        new Promise((resolve) => {
          release = () => resolve({ mime: 'image/png', base64: 'AAAA' })
        }),
    }
    globalThis.Image = FakeImage as unknown as typeof Image
    try {
      await fn(() => release())
    } finally {
      w.desktop = savedDesktop
      globalThis.Image = savedImage
    }
  }

  const replaceFirstPara = (editor: Editor) =>
    executeTool(
      editor,
      {
        id: 'w',
        name: 'replace_blocks',
        input: { startBlockIndex: 2, endBlockIndex: 2, html: '<p>rewritten</p>' },
      },
      NUM_IDS,
    )

  // the regression: settling insert_image with markDocSeen after the download
  // baptized user edits made mid-flight, letting index writes hit shifted blocks
  it('user edits during the download keep index-addressed writes stale', async () => {
    const editor = createEditor(fixture())
    await executeTool(editor, { id: 'r', name: 'get_document_context', input: {} }, NUM_IDS)
    await withImageStubs(async (release) => {
      const pending = executeTool(
        editor,
        { id: 't', name: 'insert_image', input: { url: 'https://example.com/a.png' } },
        NUM_IDS,
      )
      // the user types while the download is in flight
      editor.view.dispatch(editor.state.tr.insertText('typed by user ', 2))
      release()
      const exec = await pending
      expect(exec.isError).toBeUndefined()
      const stale = await replaceFirstPara(editor)
      expect(stale.isError).toBe(true)
      expect(stale.output).toContain('edited by the user')
    })
  })

  it('an undisturbed insert_image keeps the baseline current (no forced re-read)', async () => {
    const editor = createEditor(fixture())
    await executeTool(editor, { id: 'r', name: 'get_document_context', input: {} }, NUM_IDS)
    await withImageStubs(async (release) => {
      const pending = executeTool(
        editor,
        { id: 't', name: 'insert_image', input: { url: 'https://example.com/a.png' } },
        NUM_IDS,
      )
      release()
      const exec = await pending
      expect(exec.isError).toBeUndefined()
      const write = await replaceFirstPara(editor)
      expect(write.isError).toBeUndefined()
    })
  })
})

describe('abort during async tools', () => {
  it('insert_image aborted mid-download writes nothing', async () => {
    const editor = createEditor(fixture())
    const before = JSON.stringify(editor.getJSON())
    const w = window as unknown as { desktop?: { fetchImage(url: string): Promise<unknown> } }
    const saved = w.desktop
    w.desktop = {
      fetchImage: async () => ({ mime: 'image/png', base64: 'AAAA' }),
    }
    try {
      const ctrl = new AbortController()
      ctrl.abort()
      const skill = createDocsSkill(
        () => editor,
        () => NUM_IDS,
      )
      const exec = await skill.executeTool(
        { id: 't', name: 'insert_image', input: { url: 'https://example.com/a.png' } },
        ctrl.signal,
      )
      expect(exec.isError).toBe(true)
      expect(exec.output).toContain('stopped by the user')
      expect(JSON.stringify(editor.getJSON())).toBe(before)
    } finally {
      w.desktop = saved
    }
  })
})
