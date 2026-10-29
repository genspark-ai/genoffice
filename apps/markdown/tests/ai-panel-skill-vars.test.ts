import { afterEach, beforeAll, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { AiPanel, type MarkdownAiDeps } from '../src/renderer/ai/AiPanel'
import type { MarkdownApi } from '../src/shared/ipc'

// Skill context variables reach the model only through the send path: the
// panel reads them when the composer sends — the app's localized name, the
// saved file's base name, and the selection the run targets (the redacted view
// the agent itself sees, capped at 500 chars) — and expands them inside the
// attached skill's body. Driven through the mounted composer, in the style of
// ai-panel-slash.test.ts: the slash palette attaches the skill, Enter sends,
// and the request the transport receives must carry the expanded body.

const roots: Root[] = []
// Undestroyed views leave DOMObserver flush timers that fire after jsdom
// teardown ("document is not defined" unhandled error) — destroy every editor.
const editors: Editor[] = []

beforeAll(() => {
  ;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
  Element.prototype.scrollTo ??= () => {}
})

afterEach(() => {
  for (const root of roots.splice(0)) act(() => root.unmount())
  for (const editor of editors.splice(0)) editor.destroy()
  document.body.innerHTML = ''
  delete (window as unknown as { markdownApi?: MarkdownApi }).markdownApi
})

function createEditor(md: string): Editor {
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
  editor.commands.setContent(md, { contentType: 'markdown' })
  editors.push(editor)
  return editor
}

/** select the full text of the document's first paragraph block */
function selectParagraph(editor: Editor): void {
  let range: { from: number; to: number } | undefined
  editor.state.doc.forEach((node, offset) => {
    if (!range && node.type.name === 'paragraph')
      range = { from: offset + 1, to: offset + 1 + node.content.size }
  })
  editor.commands.setTextSelection(range ?? { from: 1, to: 1 })
}

type StreamRequest = { requestId: string; messages: Array<{ role: string; text: string }> }

/** the panel's own IPC surface, minus everything a mounted send needs; the
 *  stream fake answers with one delta and a done chunk so the run settles */
function installApi(skillBody: string): { streams: StreamRequest[] } {
  const streams: StreamRequest[] = []
  let chunkListener: ((chunk: { requestId: string; type: string; text?: string }) => void) | null =
    null
  const api = {
    aiGskStatus: async () => ({ loggedIn: false }),
    getAiSettings: async () => ({
      provider: 'openai',
      providers: { openai: { apiKey: 'k', model: 'gpt' } },
    }),
    setAiSettings: async () => {},
    onAiSettingsChanged: () => () => {},
    openAiModelSettings: async () => {},
    skillsList: async () => [{ name: 'vars', description: 'expand the context' }],
    skillText: async () => skillBody,
    setAiPanelPrefs: async (prefs: unknown) => prefs,
    openSourcePath: async () => {},
    aiStream: async (request: StreamRequest) => {
      streams.push(request)
      setTimeout(() => {
        chunkListener?.({ requestId: request.requestId, type: 'delta', text: 'Done.' })
        chunkListener?.({ requestId: request.requestId, type: 'done' })
      }, 0)
    },
    aiStreamCancel: async () => {},
    onAiStream: (handler: (chunk: { requestId: string; type: string; text?: string }) => void) => {
      chunkListener = handler
      return () => {
        chunkListener = null
      }
    },
  }
  ;(window as unknown as { markdownApi: MarkdownApi }).markdownApi = api as unknown as MarkdownApi
  return { streams }
}

function mountPanel(editor: Editor | null, filePath: string | null): HTMLDivElement {
  const deps: MarkdownAiDeps = {
    getEditor: () => editor,
    getFrontmatter: () => '',
    setFrontmatter: () => {},
    getSnapshot: () => ({ body: '', frontmatter: '' }),
    restoreSnapshot: () => {},
    onRunDone: () => {},
  }
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  roots.push(root)
  act(() => {
    root.render(createElement(AiPanel, { deps, filePath, onCollapse: () => {} }))
  })
  return container
}

function typeInto(textarea: HTMLTextAreaElement, text: string): void {
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
  act(() => {
    setter.call(textarea, text)
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

/** pump timers until the condition holds (the send pipeline is promise-driven) */
async function waitFor(cond: () => boolean): Promise<void> {
  for (let i = 0; i < 60 && !cond(); i++) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10))
    })
  }
}

/** attach the skill through the slash palette, then send the request */
async function attachAndSend(container: HTMLDivElement, request: string): Promise<void> {
  const textarea = container.querySelector<HTMLTextAreaElement>('textarea')!
  typeInto(textarea, '/vars')
  await waitFor(() => container.querySelector('[role="option"]') !== null)
  act(() => {
    textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
  })
  typeInto(textarea, request)
  act(() => {
    textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
  })
}

/** the user message of the single stream the send produced */
const sentUserText = (streams: StreamRequest[]): string => {
  const last = streams[0]!.messages.at(-1)!
  expect(last.role).toBe('user')
  return last.text
}

/** the body of the single skill block the send produced */
const sentSkillBody = (streams: StreamRequest[]): string => {
  const text = sentUserText(streams)
  const body = /<skill name="[^"]*">\n([\s\S]*?)\n<\/skill>/.exec(text)
  expect(body, 'one skill block in the sent instruction').not.toBeNull()
  return body![1]!
}

describe('AiPanel skill context vars', () => {
  it('expands the app, file and selection into the skill body at send time', async () => {
    const today = new Date().toLocaleDateString('zh')
    const { streams } = installApi(
      'Work in ${APP} (${LANGUAGE}) on ${FILE_NAME} today ${TODAY}.\nSelection: ${SELECTION}\nUnknown: ${NOPE}',
    )
    const editor = createEditor('# Notes\n\nThe quarterly numbers look strong.')
    selectParagraph(editor)
    const container = mountPanel(editor, 'D:\\GitHub\\notes\\Q3 report.md')

    await attachAndSend(container, 'Polish this')
    await waitFor(() => streams.length === 1)
    await waitFor(
      () => container.querySelector('.ai-msg-assistant:not(.ai-msg-streaming)') !== null,
    )

    expect(streams.length).toBe(1)
    const text = sentUserText(streams)
    expect(text).toContain(`Work in Markdown (中文) on Q3 report.md today ${today}.`)
    expect(text).toContain('Selection: The quarterly numbers look strong.')
    expect(text).toContain('Unknown: ${NOPE}')
    expect(text).toContain('User request:\nPolish this')
  })

  it('caps the selection at 500 chars and keeps an unsaved file out of FILE_NAME', async () => {
    const { streams } = installApi('File: ${FILE_NAME}\nSelection: ${SELECTION}')
    const editor = createEditor('x'.repeat(600))
    selectParagraph(editor)
    const container = mountPanel(editor, null)

    await attachAndSend(container, 'Trim this')
    await waitFor(() => streams.length === 1)

    expect(streams.length).toBe(1)
    const body = sentSkillBody(streams)
    expect(body).toContain('File: ${FILE_NAME}')
    expect(body).toContain(`Selection: ${'x'.repeat(500)}…`)
    expect(body).not.toContain('x'.repeat(501))
  })

  it('leaves ${SELECTION} literal when no editor is attached', async () => {
    const { streams } = installApi('File: ${FILE_NAME}\nSelection: ${SELECTION}')
    const container = mountPanel(null, 'notes/todo.md')

    await attachAndSend(container, 'Summarize')
    await waitFor(() => streams.length === 1)

    expect(streams.length).toBe(1)
    const text = sentUserText(streams)
    expect(text).toContain('File: todo.md')
    expect(text).toContain('Selection: ${SELECTION}')
  })
})
