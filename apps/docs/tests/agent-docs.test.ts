import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { AgentLoop, type AgentStreamCallbacks, type AgentTransport } from '@genoffice/agent-core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createDocsSkill } from '../src/renderer/ai/docs-skill'
import { buildDocContext, countWords } from '../src/renderer/ai/protocol'
import { executeTool } from '../src/renderer/ai/tools'
import { appendStreamedNodes } from '../src/renderer/file-actions'

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
  it('after the model calls apply_ops, all headings turn red with the aiChanged highlight', async () => {
    const editor = createEditor(fixture())
    let final = ''
    const loop = makeLoop(
      editor,
      scriptedTransport([
        (cb) => {
          cb.onToolCall({
            id: 't1',
            name: 'apply_ops',
            input: {
              ops: [{ op: 'setFont', target: { nodeType: 'docHeading' }, color: 'FF0000' }],
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
            name: 'apply_ops',
            input: { ops: [{ op: 'setFont' }] },
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

  it('a streamed load tail is not a user edit; a real edit after it still is', async () => {
    const editor = createEditor(fixture())
    await read(editor)
    appendStreamedNodes(editor, [para('streamed tail')])
    const written = await replace(editor)
    expect(written.isError).toBeFalsy()
    expect(editor.state.doc.child(1).textContent).toBe('rewritten')
    editor.view.dispatch(editor.state.tr.insertText('typed by user ', 2))
    appendStreamedNodes(editor, [para('later tail')])
    const stale = await replace(editor)
    expect(stale.isError).toBe(true)
    expect(stale.output).toContain('edited by the user')
  })

  it('index-addressed writes fail after a user edit, and succeed again after a re-read', async () => {
    const editor = createEditor(fixture())
    await read(editor)
    // simulate a manual user edit between tool calls
    editor.view.dispatch(editor.state.tr.insertText('typed by user ', 2))
    const stale = await executeTool(
      editor,
      {
        id: 'w',
        name: 'apply_ops',
        input: { ops: [{ op: 'deleteBlocks', target: { blockIndexes: [3] } }] },
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
          // real PNG magic bytes: the tool sniffs the payload before trusting the mime
          release = () => resolve({ mime: 'image/png', base64: 'iVBORw0KGgoAAAAA' })
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

describe('selection scope freezing', () => {
  it('tools act on the selection captured at context build, not on a mid-run click elsewhere', async () => {
    const editor = createEditor(fixture())
    const skill = createDocsSkill(
      () => editor,
      () => NUM_IDS,
    )
    // the user selects inside block 1, then the run starts (context build freezes the scope)
    const block0Size = editor.state.doc.child(0).nodeSize
    editor.commands.setTextSelection({ from: block0Size + 2, to: block0Size + 6 })
    const context = skill.buildContext!()
    expect(context).toContain('Current selection: block 1')
    // mid-run the user clicks into block 3 (selection-only change, doc untouched)
    const block3Pos =
      block0Size + editor.state.doc.child(1).nodeSize + editor.state.doc.child(2).nodeSize
    editor.commands.setTextSelection(block3Pos + 2)
    const exec = await skill.executeTool({
      id: 't',
      name: 'apply_ops',
      input: {
        ops: [
          {
            op: 'setParagraphFormat',
            target: { nodeType: 'docParagraph', scope: 'selection' },
            align: 'right',
          },
        ],
      },
    })
    expect(exec.isError).toBeFalsy()
    expect(editor.state.doc.child(1).attrs.align).toBe('right')
    expect(editor.state.doc.child(3).attrs.align).not.toBe('right')
  })

  it('once the AI itself edits the doc, the freeze yields to the PM-remapped live selection', async () => {
    const editor = createEditor(fixture())
    const skill = createDocsSkill(
      () => editor,
      () => NUM_IDS,
    )
    // select inside block 1 (the first body paragraph) and freeze
    const block0Size = editor.state.doc.child(0).nodeSize
    editor.commands.setTextSelection({ from: block0Size + 2, to: block0Size + 6 })
    skill.buildContext!()
    // the AI inserts a paragraph at the doc start: every block shifts by one,
    // and ProseMirror remaps the live selection into the original paragraph
    const insert = await skill.executeTool({
      id: 'i',
      name: 'insert_content',
      input: { html: '<p>Lead-in</p>', afterBlockIndex: -1 },
    })
    expect(insert.isError).toBeFalsy()
    const exec = await skill.executeTool({
      id: 't',
      name: 'apply_ops',
      input: {
        ops: [
          {
            op: 'setParagraphFormat',
            target: { nodeType: 'docParagraph', scope: 'selection' },
            align: 'right',
          },
        ],
      },
    })
    expect(exec.isError).toBeFalsy()
    // the originally selected paragraph now sits at index 2 and must be the one styled
    expect(editor.state.doc.child(2).textContent).toBe('GenSpark is an AI office suite.')
    expect(editor.state.doc.child(2).attrs.align).toBe('right')
  })
})

describe('partial selection: context markers and replace_selection', () => {
  // block 1 'GenSpark is an AI office suite.' starts at pos 20, content at 21; 'AI office' = offsets 15..24
  const SEL = { from: 36, to: 45 }

  it('the context marks the selected span with <sel> and repeats it as "Selected text"', () => {
    const editor = createEditor(fixture())
    editor.commands.setTextSelection(SEL)
    const context = buildDocContext(editor)
    expect(context).toContain('<p>GenSpark is an <sel>AI office</sel> suite.</p>')
    expect(context).toContain('Selected text (9 characters): "AI office"')
    expect(context).toContain('Current selection: block 1 (part of the text only')
  })

  it('a selection covering whole blocks carries no markers', () => {
    const editor = createEditor(fixture())
    editor.commands.setTextSelection({
      from: 21,
      to: 21 + 'GenSpark is an AI office suite.'.length,
    })
    const context = buildDocContext(editor)
    expect(context).toContain('<p>GenSpark is an AI office suite.</p>')
    expect(context).not.toContain('<sel>')
    expect(context).not.toContain('Selected text (')
  })

  it('markers split a styled run and span across blocks per block', () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: null },
        content: [
          text('Revenue grew '),
          { type: 'text', text: '8% year', marks: [{ type: 'bold' }] },
        ],
      },
      para('Costs fell.'),
    ])
    // select from 'grew' (offset 8 of block 0) through 'Costs' (offsets 0..5 of block 1)
    const block0Size = editor.state.doc.child(0).nodeSize
    editor.commands.setTextSelection({ from: 1 + 8, to: block0Size + 1 + 5 })
    const context = buildDocContext(editor)
    expect(context).toContain('<p>Revenue <sel>grew <strong>8% year</strong></sel></p>')
    expect(context).toContain('<p><sel>Costs</sel> fell.</p>')
    expect(context).toContain('Selected text (18 characters): "grew 8% year\nCosts"')
  })

  it('replace_selection replaces exactly the frozen span, even after the live selection moved', async () => {
    const editor = createEditor(fixture())
    const skill = createDocsSkill(
      () => editor,
      () => NUM_IDS,
    )
    editor.commands.setTextSelection(SEL)
    skill.buildContext!()
    editor.commands.setTextSelection(editor.state.doc.content.size - 3)
    const exec = await skill.executeTool({
      id: 'r',
      name: 'replace_selection',
      input: { html: 'AI-native productivity' },
    })
    expect(exec.isError).toBeFalsy()
    expect(exec.mutated).toBe(true)
    expect(editor.state.doc.child(1).textContent).toBe(
      'GenSpark is an AI-native productivity suite.',
    )
    expect(editor.state.doc.child(1).attrs.aiChanged).toBe(true)
    expect(editor.state.doc.child(3).textContent).toBe('This document is for reference only.')
    // the new text is selected so a follow-up scope:'selection' command targets it
    expect(editor.state.selection.from).toBe(SEL.from)
    expect(editor.state.selection.to).toBe(SEL.from + 'AI-native productivity'.length)
  })

  it('replace_selection inherits the formatting of the replaced text unless the fragment styles it', async () => {
    const styled = () => [
      {
        type: 'docParagraph',
        attrs: { docxIndex: null },
        content: [
          text('Revenue grew '),
          {
            type: 'text',
            text: '8%',
            marks: [{ type: 'bold' }, { type: 'docTextStyle', attrs: { color: 'FF0000' } }],
          },
          text(' year over year.'),
        ],
      },
    ]
    const plain = createEditor(styled())
    plain.commands.setTextSelection({ from: 14, to: 16 })
    await executeTool(plain, { id: 'a', name: 'replace_selection', input: { html: '9%' } }, NUM_IDS)
    const p1 = plain.state.doc.child(0)
    expect(p1.textContent).toBe('Revenue grew 9% year over year.')
    expect(p1.child(1).text).toBe('9%')
    expect(
      p1
        .child(1)
        .marks.map((m) => m.type.name)
        .sort(),
    ).toEqual(['bold', 'docTextStyle'])
    expect(p1.child(1).marks.find((m) => m.type.name === 'docTextStyle')?.attrs).toMatchObject({
      color: 'FF0000',
    })

    const fragment = createEditor(styled())
    fragment.commands.setTextSelection({ from: 14, to: 16 })
    await executeTool(
      fragment,
      { id: 'b', name: 'replace_selection', input: { html: '<em>9%</em>' } },
      NUM_IDS,
    )
    const p2 = fragment.state.doc.child(0)
    // the fragment's own formatting replaces the inherited bold; run-level style (color) is kept
    expect(
      p2
        .child(1)
        .marks.map((m) => m.type.name)
        .sort(),
    ).toEqual(['docTextStyle', 'italic'])
  })

  it('replace_selection in tracking mode keeps the old text struck through and inserts the new text as ins', async () => {
    const editor = createEditor(fixture())
    editor.commands.setTextSelection(SEL)
    const exec = await executeTool(
      editor,
      { id: 't', name: 'replace_selection', input: { html: 'AI-native' } },
      NUM_IDS,
      { author: 'AI Assistant' },
    )
    expect(exec.isError).toBeFalsy()
    const runs: Array<[string, string[]]> = []
    editor.state.doc.child(1).forEach((child) => {
      runs.push([child.text ?? '', child.marks.map((m) => m.type.name).sort()])
    })
    expect(runs).toEqual([
      ['GenSpark is an ', []],
      ['AI office', ['del']],
      ['AI-native', ['ins']],
      [' suite.', []],
    ])
    expect(editor.state.doc.child(1).attrs.aiChanged).toBe(false)
  })

  it('replace_selection refuses a caret, a multi-block range, block HTML and the <sel> marker', async () => {
    const editor = createEditor(fixture())
    editor.commands.setTextSelection(30)
    const caret = await executeTool(
      editor,
      { id: '1', name: 'replace_selection', input: { html: 'x' } },
      NUM_IDS,
    )
    expect(caret.isError).toBe(true)
    expect(caret.output).toContain('needs a range selection')

    editor.commands.setTextSelection({ from: 36, to: editor.state.doc.content.size - 3 })
    const multi = await executeTool(
      editor,
      { id: '2', name: 'replace_selection', input: { html: 'x' } },
      NUM_IDS,
    )
    expect(multi.isError).toBe(true)
    expect(multi.output).toContain('replace_blocks')

    editor.commands.setTextSelection(SEL)
    const block = await executeTool(
      editor,
      { id: '3', name: 'replace_selection', input: { html: '<p>a</p><p>b</p>' } },
      NUM_IDS,
    )
    expect(block.isError).toBe(true)
    expect(block.output).toContain('replace_blocks')

    const marker = await executeTool(
      editor,
      { id: '4', name: 'replace_selection', input: { html: '<sel>AI</sel>' } },
      NUM_IDS,
    )
    expect(marker.isError).toBe(true)
    expect(marker.output).toContain('<sel>')

    const echoed = await executeTool(
      editor,
      { id: '5', name: 'insert_content', input: { html: '<p>see <sel>this</sel></p>' } },
      NUM_IDS,
    )
    expect(echoed.isError).toBe(true)
    expect(editor.state.doc.child(1).textContent).toBe('GenSpark is an AI office suite.')
  })

  it('a single <p> wrapper and plain text are both accepted as inline replacements', async () => {
    const editor = createEditor(fixture())
    editor.commands.setTextSelection(SEL)
    const exec = await executeTool(
      editor,
      {
        id: 'p',
        name: 'replace_selection',
        input: { html: '<p>an <strong>AI</strong> office</p>' },
      },
      NUM_IDS,
    )
    expect(exec.isError).toBeFalsy()
    const block = editor.state.doc.child(1)
    expect(block.textContent).toBe('GenSpark is an an AI office suite.')
    // the unstyled lead-in merges with the preceding run; the bold slice stays its own node
    expect(block.child(1).text).toBe('AI')
    expect(block.child(1).marks.some((m) => m.type.name === 'bold')).toBe(true)
  })
})

describe('partial selection: table cells and code blocks', () => {
  it('a range inside a table cell keeps whole-block semantics (no markers, no partial note)', async () => {
    const editor = createEditor([
      heading('Chapter 1 Overview', 1),
      {
        type: 'docTable',
        attrs: { docxIndex: null },
        content: [
          {
            type: 'docTableRow',
            content: [
              {
                type: 'docTableCell',
                content: [
                  { type: 'docParagraph', attrs: { docxIndex: null }, content: [text('City GDP')] },
                ],
              },
            ],
          },
        ],
      },
    ])
    // cell paragraph content starts at 20 (table) + 1 (row) + 1 (cell) + 1 (paragraph) + 1
    editor.commands.setTextSelection({ from: 25, to: 28 })
    expect(editor.state.doc.textBetween(25, 28)).toBe('ity')
    const context = buildDocContext(editor)
    expect(context).toContain('Current selection: block 1\n')
    expect(context).not.toContain('<sel>')
    expect(context).not.toContain('Selected text (')
    const exec = await executeTool(
      editor,
      { id: 'c', name: 'replace_selection', input: { html: 'x' } },
      NUM_IDS,
    )
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('replace_blocks')
  })

  it('a code block marks the selected span inside its <pre>', () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: null, shadingFill: 'F2F2F2', borders: 'tblr' },
        content: [text('const a = 1;'), { type: 'hardBreak' }, text('const b = <2>;')],
      },
    ])
    // 'b = <2>' = offsets 19..26 (the hard break occupies one position)
    editor.commands.setTextSelection({ from: 1 + 19, to: 1 + 26 })
    const context = buildDocContext(editor)
    expect(context).toContain('<pre>const a = 1;\nconst <sel>b = &lt;2&gt;</sel>;</pre>')
    expect(context).toContain('Selected text (7 characters): "b = <2>"')
  })
})
