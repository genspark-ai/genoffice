import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { defaultAiSettings, type AiSettings } from '@genoffice/ai-provider'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import type { PdfAiDeps } from '../src/renderer/ai/tools'

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

function settings(provider: AiSettings['provider']): AiSettings {
  return { ...defaultAiSettings(), provider }
}

function deps(): PdfAiDeps {
  return {
    doc: () => null,
    fileName: () => 'test.pdf',
    pageCount: () => 1,
    currentPage: () => 1,
    readOnly: () => false,
    outline: () => null,
    searchIndex: () => null,
    isDeleted: () => false,
    gotoPage: () => true,
    addMarkup: () => {},
    editText: async () => null,
    editFonts: () => [],
    formEdits: () => new Map(),
    applyFormEdit: () => {},
    rotatePage: () => {},
    deletePage: () => true,
    pageGeom: () => null,
    listImages: async () => [],
    isImageClaimed: () => false,
    insertImage: () => {},
    transformImage: () => {},
    replaceImage: () => {},
    deleteImage: () => {},
    searchImages: async () => ({ images: [], method: 'test' }),
    generateImage: async () => ({}),
    fetchImage: async () => null,
  }
}

function installApi(initial: AiSettings, loggedIn = false) {
  let settingsChanged: ((next: AiSettings) => void) | undefined
  const api = {
    getAiSettings: vi.fn(async () => initial),
    setAiSettings: vi.fn(async () => {}),
    onAiSettingsChanged: vi.fn((handler: (next: AiSettings) => void) => {
      settingsChanged = handler
      return () => {
        settingsChanged = undefined
      }
    }),
    aiCodexStatus: vi.fn(async () => ({ loggedIn })),
    aiCodexLogin: vi.fn(async () => ({ loggedIn: true })),
    aiCodexCancelLogin: vi.fn(async () => {}),
    aiCodexLogout: vi.fn(async () => ({ loggedIn: false })),
    aiCodexCapabilities: vi.fn(async () => ({
      models: [
        {
          id: 'gpt-5.5',
          name: 'GPT-5.5',
          reasoningEfforts: ['none', 'high'],
          serviceTiers: [{ id: 'default', name: 'Standard' }],
        },
      ],
    })),
    onAiStream: vi.fn(() => () => {}),
    aiStream: vi.fn(async () => {}),
    aiStreamCancel: vi.fn(async () => {}),
  }
  Object.defineProperty(window, 'pdfApi', { configurable: true, value: api })
  return {
    api,
    emit(next: AiSettings) {
      settingsChanged?.(next)
    },
  }
}

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

afterEach(() => {
  document.body.replaceChildren()
})

describe('PDF Codex provider wiring', () => {
  it('loads shared settings, gates signed-out Send, and keeps drafts editable', async () => {
    const initial = settings('openai-codex')
    const { api } = installApi(initial)
    const { container, cleanup } = mount(
      createElement(AiPanel, { api: deps(), onCollapse: () => {} }),
    )
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(container.querySelector('.ai-provider-select-input')).not.toBeNull()
    expect(container.querySelector<HTMLButtonElement>('.ai-send-btn')?.disabled).toBe(true)
    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
    act(() => {
      setter.call(textarea, 'draft')
      textarea.dispatchEvent(new Event('input', { bubbles: true }))
    })
    expect(textarea.value).toBe('draft')
    expect(api.setAiSettings).not.toHaveBeenCalled()
    cleanup()
  })

  it('shows capabilities and accepts an external settings event without writing it back', async () => {
    const initial = settings('openai-codex')
    const installed = installApi(initial, true)
    const { container, cleanup } = mount(
      createElement(AiPanel, { api: deps(), onCollapse: () => {} }),
    )
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(container.querySelector('.ai-codex-model-trigger')).not.toBeNull()
    installed.emit(settings('genspark'))
    await act(async () => {
      await Promise.resolve()
    })
    expect(container.querySelector<HTMLSelectElement>('.ai-provider-select-input')?.value).toBe(
      'genspark',
    )
    expect(installed.api.setAiSettings).not.toHaveBeenCalled()
    cleanup()
  })
})
