// The docs panel's approval gate: risk:'high' write tools park at an inline
// allow/deny card before they execute. Denying must leave the document
// untouched and hand the model a do-not-retry result; allowing runs the tool.
// These drive the real send path (AgentLoop + IPC transport + docs skill).
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'
import { TOOL_DENIED_OUTPUT } from '@genoffice/agent-core'

type ChunkListener = (chunk: {
  requestId: string
  type: string
  text?: string
  toolCall?: { id: string; name: string; input: Record<string, unknown> }
}) => void

function settings(): AiSettings {
  return {
    provider: 'gemini',
    providers: Object.fromEntries(
      AI_PROVIDERS.map((p) => [
        p.id,
        p.id === 'gemini'
          ? { apiKey: 'k', model: p.defaultModel }
          : { apiKey: '', model: p.defaultModel },
      ]),
    ) as AiSettings['providers'],
  }
}

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
          content: [{ type: 'text', text: 'original body' }],
        },
      ],
    },
  })
}

const editorText = (editor: Editor): string => editor.state.doc.textContent

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

const flush = (): Promise<void> => new Promise((r) => setTimeout(r, 0))

beforeAll(() => {
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

/**
 * The transport answers turn 1 with an insert_content call, turn 2 with text.
 * Records every request so the assertions can see what the model received.
 */
function stubTransport() {
  const started: Array<{
    requestId: string
    messages: Array<{ role: string; results?: Array<{ output: string }> }>
  }> = []
  let listener: ChunkListener | null = null
  ;(window as unknown as Record<string, unknown>).desktop = {
    aiGskStatus: vi.fn(async () => ({ loggedIn: true })),
    onAiStream: (l: ChunkListener) => {
      listener = l
      return () => {
        listener = null
      }
    },
    aiStream: vi.fn(async (request: { requestId: string } & Record<string, unknown>) => {
      started.push({
        requestId: request.requestId,
        messages: request.messages as Array<{ role: string; results?: Array<{ output: string }> }>,
      })
      queueMicrotask(() => {
        if (started.length === 1) {
          listener?.({
            requestId: request.requestId,
            type: 'tool-call',
            toolCall: {
              id: 't1',
              name: 'insert_content',
              input: { html: '<p>AI inserted</p>', afterBlockIndex: -1 },
            },
          })
        } else {
          listener?.({ requestId: request.requestId, type: 'delta', text: 'ok' })
        }
        listener?.({ requestId: request.requestId, type: 'done' })
      })
    }),
    aiStreamCancel: vi.fn(async () => {}),
  }
  return { started }
}

async function send(container: HTMLElement): Promise<void> {
  const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
  typeInto(textarea, 'add a paragraph at the top')
  const send = container.querySelector<HTMLButtonElement>('.ai-send-btn')!
  await act(async () => {
    send.click()
    for (let i = 0; i < 8; i++) await flush()
  })
}

describe('AI panel tool confirmation', () => {
  it('parks a write tool at the allow/deny card; denying leaves the document untouched', async () => {
    const editor = createEditor()
    try {
      const { started } = stubTransport()
      const mounted = mount(
        createElement(AiPanel, {
          editor,
          blocks: [],
          settings: settings(),
          open: true,
          onExpand: () => {},
          onCollapse: () => {},
        }),
      )
      const { container } = mounted

      await send(container)
      // the card is in the conversation and the loop holds at the gate
      expect(container.querySelector('.ai-confirm-card')).not.toBeNull()
      expect(started).toHaveLength(1)

      const deny = container.querySelector<HTMLButtonElement>('.ai-confirm-deny')!
      await act(async () => {
        deny.click()
        for (let i = 0; i < 8; i++) await flush()
      })

      // the tool never ran: the document still says what it said
      expect(editorText(editor)).toBe('original body')
      // the model received the denial as an error tool result it can act on
      expect(started).toHaveLength(2)
      const toolMessage = started[1]!.messages.find((m) => m.role === 'tool')
      expect(toolMessage?.results?.[0]?.output).toBe(TOOL_DENIED_OUTPUT)
      // the settled card stays in the flow as the record of the decision
      expect(container.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe(
        'denied',
      )
      expect(container.querySelector('.ai-confirm-deny')).toBeNull()
      act(() => mounted.root.unmount())
    } finally {
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })

  it('allowing executes the write tool', async () => {
    const editor = createEditor()
    try {
      const { started } = stubTransport()
      const mounted = mount(
        createElement(AiPanel, {
          editor,
          blocks: [],
          settings: settings(),
          open: true,
          onExpand: () => {},
          onCollapse: () => {},
        }),
      )
      const { container } = mounted

      await send(container)
      const allow = container.querySelector<HTMLButtonElement>('.ai-confirm-allow')!
      await act(async () => {
        allow.click()
        for (let i = 0; i < 8; i++) await flush()
      })

      // the write happened through the real editor path
      expect(editorText(editor)).toContain('AI inserted')
      // the tool result went back to the model (its round-trip continued)
      expect(started).toHaveLength(2)
      const toolMessage = started[1]!.messages.find((m) => m.role === 'tool')
      expect(toolMessage?.results?.[0]?.output).not.toBe(TOOL_DENIED_OUTPUT)
      expect(container.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe(
        'allowed',
      )
      act(() => mounted.root.unmount())
    } finally {
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })

  it('Stop while the card is open settles it as denied without editing', async () => {
    const editor = createEditor()
    try {
      const { started } = stubTransport()
      const mounted = mount(
        createElement(AiPanel, {
          editor,
          blocks: [],
          settings: settings(),
          open: true,
          onExpand: () => {},
          onCollapse: () => {},
        }),
      )
      const { container } = mounted

      await send(container)
      expect(container.querySelector('.ai-confirm-card')).not.toBeNull()
      // Esc is the composer's stop; the stop button is visible while busy
      const stop = container.querySelector<HTMLButtonElement>('.ai-stop-btn')!
      await act(async () => {
        stop.click()
        for (let i = 0; i < 8; i++) await flush()
      })

      expect(editorText(editor)).toBe('original body')
      // the gate settled as denied and the run ended without a second turn
      expect(container.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe(
        'denied',
      )
      expect(started).toHaveLength(1)
      act(() => mounted.root.unmount())
    } finally {
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })
})
