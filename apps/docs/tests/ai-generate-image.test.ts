import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { executeTool } from '../src/renderer/ai/tools'

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

/** run one generate_image call with a stubbed desktop bridge */
async function runWithDesktop(
  desktop: Record<string, unknown>,
  input: Record<string, unknown>,
  signal?: AbortSignal,
) {
  const editor = createEditor()
  const w = window as unknown as DesktopStub
  const saved = w.desktop
  w.desktop = desktop
  try {
    return {
      exec: await executeTool(
        editor,
        { id: 't', name: 'generate_image', input },
        NUM_IDS,
        undefined,
        signal,
      ),
      editor,
    }
  } finally {
    w.desktop = saved
  }
}

describe('generate_image', () => {
  it('rejects an empty prompt without calling the channel', async () => {
    let called = false
    const { exec } = await runWithDesktop(
      {
        aiGenerateImage: () => {
          called = true
          return Promise.resolve({ url: 'https://example.com/a.png' })
        },
      },
      { prompt: '   ' },
    )
    expect(exec.isError).toBe(true)
    expect(called).toBe(false)
  })

  it('surfaces the channel error (not logged in / cloud tools off / generation failure)', async () => {
    const { exec, editor } = await runWithDesktop(
      { aiGenerateImage: () => Promise.resolve({ error: 'Genspark account is not logged in' }) },
      { prompt: 'a watercolor fox' },
    )
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('not logged in')
    expect(editor.state.doc.childCount).toBe(1) // nothing inserted
  })

  it('an abort after generation never writes into the document', async () => {
    const ctrl = new AbortController()
    const { exec, editor } = await runWithDesktop(
      {
        aiGenerateImage: () => {
          ctrl.abort()
          return Promise.resolve({ url: 'https://example.com/a.png' })
        },
        fetchImage: () => Promise.resolve({ base64: 'AAAA', mime: 'image/png' }),
      },
      { prompt: 'a watercolor fox' },
      ctrl.signal,
    )
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('stopped by the user')
    expect(editor.state.doc.childCount).toBe(1)
  })

  it('passes the prompt and aspect ratio through to the channel', async () => {
    let received: unknown = null
    await runWithDesktop(
      {
        aiGenerateImage: (op: unknown) => {
          received = op
          return Promise.resolve({ error: 'stop here' }) // fail before the jsdom-unfriendly decode
        },
      },
      { prompt: 'a watercolor fox', aspectRatio: '16:9' },
    )
    expect(received).toEqual({ prompt: 'a watercolor fox', aspectRatio: '16:9' })
  })
})
