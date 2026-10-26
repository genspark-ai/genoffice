// The empty state advertises a setup CTA only when no provider can serve a
// chat request and the user is not signed into Genspark — a configured BYOK
// user or a signed-in user must never see it.
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import { AI_PROVIDERS, type AiProviderId, type AiSettings } from '../src/shared/ipc'
import { setModuleLang } from '../src/renderer/i18n/locale'

function settingsWith(provider: AiProviderId, apiKey: string): AiSettings {
  return {
    provider,
    providers: Object.fromEntries(
      AI_PROVIDERS.map((p) => [
        p.id,
        p.id === provider
          ? { apiKey, model: p.defaultModel }
          : { apiKey: '', model: p.defaultModel },
      ]),
    ),
  } as AiSettings
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
          content: [{ type: 'text', text: 'Hello' }],
        },
      ],
    },
  })
}

function mount(element: React.ReactElement): { container: HTMLElement; cleanup: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() => root.render(element))
  return {
    container,
    cleanup: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
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

setModuleLang('en')

function mockDesktop(loggedIn: boolean) {
  const desktop = {
    aiGskStatus: vi.fn(async () => ({ loggedIn })),
    onAiStream: () => () => {},
    aiStream: vi.fn(async () => {}),
    aiStreamCancel: vi.fn(async () => {}),
    openAiModelSettings: vi.fn(async () => {}),
  }
  ;(window as unknown as Record<string, unknown>).desktop = desktop
  return desktop
}

function setupCta(container: HTMLElement): HTMLButtonElement | null {
  return container.querySelector('.ai-setup-needed')
}

describe('AI empty-state setup CTA', () => {
  it('shows the CTA when nothing is configured and Genspark is signed out', async () => {
    mockDesktop(false)
    const editor = createEditor()
    const m = mount(
      createElement(AiPanel, {
        editor,
        blocks: [],
        settings: settingsWith('gemini', ''),
        open: true,
        onExpand: () => {},
        onCollapse: () => {},
      }),
    )
    try {
      await act(() => flush())
      expect(setupCta(m.container)).not.toBeNull()
    } finally {
      m.cleanup()
      editor.destroy()
    }
  })

  it('clicking the CTA opens the AI model settings', async () => {
    const desktop = mockDesktop(false)
    const editor = createEditor()
    const m = mount(
      createElement(AiPanel, {
        editor,
        blocks: [],
        settings: settingsWith('gemini', ''),
        open: true,
        onExpand: () => {},
        onCollapse: () => {},
      }),
    )
    try {
      await act(() => flush())
      const cta = setupCta(m.container)!
      await act(async () => {
        cta.click()
        await flush()
      })
      expect(desktop.openAiModelSettings).toHaveBeenCalledOnce()
    } finally {
      m.cleanup()
      editor.destroy()
    }
  })

  it('hides the CTA when a BYOK provider is configured', async () => {
    mockDesktop(false)
    const editor = createEditor()
    const m = mount(
      createElement(AiPanel, {
        editor,
        blocks: [],
        settings: settingsWith('gemini', 'real-key'),
        open: true,
        onExpand: () => {},
        onCollapse: () => {},
      }),
    )
    try {
      await act(() => flush())
      expect(setupCta(m.container)).toBeNull()
    } finally {
      m.cleanup()
      editor.destroy()
    }
  })

  it('hides the CTA when Genspark is signed in', async () => {
    mockDesktop(true)
    const editor = createEditor()
    const m = mount(
      createElement(AiPanel, {
        editor,
        blocks: [],
        settings: settingsWith('gemini', ''),
        open: true,
        onExpand: () => {},
        onCollapse: () => {},
      }),
    )
    try {
      await act(() => flush())
      expect(setupCta(m.container)).toBeNull()
    } finally {
      m.cleanup()
      editor.destroy()
    }
  })
})
