// Attached skills expand their parameters against the panel's send context:
// positional $N (0-indexed) splits the request with shell-style quoting,
// ${FILE_NAME} takes the open document's name, and ${SELECTION} the text
// selected at send (redacted like the scope quote, capped at 500 chars).
// The drive is a real send through the stubbed preload bridge; the assertions
// read the model-bound request the IPC transport saw.
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'

type ChunkListener = (chunk: { requestId: string; type: string; text?: string }) => void

/** the slice of AgentStreamRequest the assertions need */
type StreamRequest = {
  requestId: string
  messages: Array<{ role: string; text: string }>
}

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((p) => [p.id, { apiKey: 'k', model: p.defaultModel }]),
  ) as AiSettings['providers'],
}

const DOC_TEXT = 'EVs market research'
// $1 is the second positional (0-indexed, like the agent-core tests pin down);
// "the tone" is quoted so it also exercises the shell-style arg splitting
const REQUEST = 'fix "the tone" today'
const SKILL_BODY = 'Rewrite $1 of ${FILE_NAME}. Selection: ${SELECTION}.'

function createEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: DOC_TEXT }],
        },
      ],
    },
  })
}

function mount(element: React.ReactElement): {
  container: HTMLElement
  root: Root
  cleanup: () => void
} {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  return {
    container,
    root,
    cleanup: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

function typeInto(textarea: HTMLTextAreaElement, text: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
  act(() => {
    setter.call(textarea, text)
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function press(textarea: HTMLTextAreaElement, key: string): KeyboardEvent {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })
  act(() => textarea.dispatchEvent(event))
  return event
}

const flush = (): Promise<void> => new Promise((r) => setTimeout(r, 0))

function panelProps(editor: Editor, fileName?: string) {
  return {
    editor,
    blocks: [],
    settings,
    ...(fileName ? { fileName } : {}),
    open: true,
    onExpand: () => {},
    onCollapse: () => {},
  }
}

/** stubs the preload bridge; returns every model-bound request the transport started */
function stubDesktop(): StreamRequest[] {
  const seen: StreamRequest[] = []
  let listener: ChunkListener | null = null
  ;(window as unknown as Record<string, unknown>).desktop = {
    aiGskStatus: vi.fn(async () => ({ loggedIn: true })),
    skillsList: vi.fn(async () => [{ name: 'polish', description: 'Rewrite in a livelier tone' }]),
    skillText: vi.fn(async () => SKILL_BODY),
    onAiStream: (l: ChunkListener) => {
      listener = l
      return () => {
        listener = null
      }
    },
    aiStream: vi.fn(async (request: StreamRequest) => {
      seen.push(request)
      listener?.({ requestId: request.requestId, type: 'delta', text: 'Done.' })
      listener?.({ requestId: request.requestId, type: 'done' })
    }),
    aiStreamCancel: vi.fn(async () => {}),
  }
  return seen
}

/** attach the skill through the slash palette, then type the request and send */
async function sendWithSkill(mounted: ReturnType<typeof mount>, request: string): Promise<void> {
  const textarea = mounted.container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
  typeInto(textarea, '/polish')
  await act(async () => {
    await flush()
    await flush()
  })
  press(textarea, 'Enter')
  await act(async () => {
    await flush()
  })
  typeInto(textarea, request)
  const send = mounted.container.querySelector<HTMLButtonElement>('.ai-send-btn')!
  await act(async () => {
    send.click()
    await flush()
    await flush()
    await flush()
    await flush()
  })
}

beforeAll(() => {
  // jsdom's opaque origin has no localStorage; the panel reads it at mount
  let store: Record<string, string> = {}
  Object.defineProperty(window, 'localStorage', {
    configurable: true,
    value: {
      getItem: (k: string) => store[k] ?? null,
      setItem: (k: string, v: string) => {
        store[k] = v
      },
      removeItem: (k: string) => {
        delete store[k]
      },
      clear: () => {
        store = {}
      },
    },
  })
  Element.prototype.scrollTo ??= () => {}
})

describe('attached skill context variables', () => {
  it('expands $1, ${FILE_NAME} and ${SELECTION} in the model-bound instruction', async () => {
    const editor = createEditor()
    const seen = stubDesktop()
    try {
      const mounted = mount(createElement(AiPanel, panelProps(editor, 'market-report.docx')))
      // the whole paragraph is selected: ${SELECTION} freezes at send like the scope quote
      act(() => editor.commands.setTextSelection({ from: 1, to: 1 + DOC_TEXT.length }))
      await sendWithSkill(mounted, REQUEST)
      const instruction = seen[0]!.messages.at(-1)!.text
      expect(instruction).toContain('<skill name="polish">')
      expect(instruction).toContain(
        `Rewrite the tone of market-report.docx. Selection: ${DOC_TEXT}.`,
      )
      expect(instruction).not.toContain('$1')
      expect(instruction).not.toContain('${FILE_NAME}')
      expect(instruction).not.toContain('${SELECTION}')
      expect(instruction).toContain(`User request:\n${REQUEST}`)
      act(() => mounted.root.unmount())
    } finally {
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })

  it('leaves ${FILE_NAME} and ${SELECTION} literal when the panel has neither', async () => {
    const editor = createEditor()
    const seen = stubDesktop()
    try {
      const mounted = mount(createElement(AiPanel, panelProps(editor)))
      // no file name prop, caret collapsed: both vars stay literal, $1 still expands
      await sendWithSkill(mounted, REQUEST)
      const instruction = seen[0]!.messages.at(-1)!.text
      expect(instruction).toContain('Rewrite the tone of ${FILE_NAME}. Selection: ${SELECTION}.')
      act(() => mounted.root.unmount())
    } finally {
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })
})
