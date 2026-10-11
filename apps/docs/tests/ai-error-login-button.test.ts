// A failed run marks the error bubble with a "Sign in to Genspark" button only
// when Genspark is the active provider — a BYOK user's failure (their own
// provider's rate limit, say) must not advertise a Genspark sign-in (#1948).
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import { AI_PROVIDERS, type AiProviderId, type AiSettings } from '../src/shared/ipc'

type ChunkListener = (chunk: { requestId: string; type: string; error?: string }) => void

function settingsWith(provider: AiProviderId): AiSettings {
  return {
    provider,
    providers: Object.fromEntries(
      AI_PROVIDERS.map((p) => [
        p.id,
        p.id === provider
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
          content: [{ type: 'text', text: 'EVs market research' }],
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

const flush = (): Promise<void> => new Promise((r) => setTimeout(r, 0))

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

describe('AI error bubble sign-in button', () => {
  it('shows no sign-in button for a BYOK provider failure', async () => {
    const editor = createEditor()
    try {
      // rebuild inline so the harness stays alive until the assertions ran
      let listener: ChunkListener | null = null
      ;(window as unknown as Record<string, unknown>).desktop = {
        aiGskStatus: vi.fn(async () => ({ loggedIn: false })),
        onAiStream: (l: ChunkListener) => {
          listener = l
          return () => {
            listener = null
          }
        },
        aiStream: vi.fn(async (request: { requestId: string }) => {
          listener?.({ requestId: request.requestId, type: 'error', error: 'HTTP 429: rate limit' })
        }),
        aiStreamCancel: vi.fn(async () => {}),
      }
      const mounted = mount(
        createElement(AiPanel, {
          editor,
          blocks: [],
          settings: settingsWith('gemini'),
          open: true,
          onExpand: () => {},
          onCollapse: () => {},
        }),
      )
      const { container } = mounted
      const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
      typeInto(textarea, 'summarize this')
      const send = container.querySelector<HTMLButtonElement>('.ai-send-btn')!
      await act(async () => {
        send.click()
        await flush()
        await flush()
        await flush()
        await flush()
      })
      expect(container.querySelector('.ai-msg-error')).not.toBeNull()
      expect(container.querySelector('.ai-login-btn')).toBeNull()
      act(() => mounted.root.unmount())
    } finally {
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })

  it('offers the sign-in button when Genspark is the provider', async () => {
    const editor = createEditor()
    try {
      let listener: ChunkListener | null = null
      ;(window as unknown as Record<string, unknown>).desktop = {
        aiGskStatus: vi.fn(async () => ({ loggedIn: false })),
        onAiStream: (l: ChunkListener) => {
          listener = l
          return () => {
            listener = null
          }
        },
        aiStream: vi.fn(async (request: { requestId: string }) => {
          listener?.({ requestId: request.requestId, type: 'error', error: 'HTTP 429: rate limit' })
        }),
        aiStreamCancel: vi.fn(async () => {}),
      }
      const mounted = mount(
        createElement(AiPanel, {
          editor,
          blocks: [],
          settings: settingsWith('genspark'),
          open: true,
          onExpand: () => {},
          onCollapse: () => {},
        }),
      )
      const container = mounted.container
      const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
      typeInto(textarea, 'summarize this')
      const send = container.querySelector<HTMLButtonElement>('.ai-send-btn')!
      await act(async () => {
        send.click()
        await flush()
        await flush()
        await flush()
        await flush()
      })
      expect(container.querySelector('.ai-msg-error')).not.toBeNull()
      expect(container.querySelector('.ai-login-btn')).not.toBeNull()
      act(() => mounted.root.unmount())
    } finally {
      editor.destroy()
      delete (window as unknown as Record<string, unknown>).desktop
    }
  })
})
