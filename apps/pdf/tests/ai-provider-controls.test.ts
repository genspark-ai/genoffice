import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { defaultAiSettings, type AiSettings } from '@genoffice/ai-provider'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import type { PdfAiDeps } from '../src/renderer/ai/tools'

const activeCleanups = new Set<() => void>()
const activeApiRestorers = new Set<() => void>()

function mount(element: React.ReactElement): {
  container: HTMLElement
  root: Root
  cleanup: () => void
} {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  let mounted = true
  const cleanup = () => {
    if (!mounted) return
    mounted = false
    activeCleanups.delete(cleanup)
    act(() => root.unmount())
    container.remove()
  }
  activeCleanups.add(cleanup)
  return {
    container,
    root,
    cleanup,
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
  const previous = Object.getOwnPropertyDescriptor(window, 'pdfApi')
  Object.defineProperty(window, 'pdfApi', { configurable: true, value: api })
  const restore = () => {
    if (!activeApiRestorers.delete(restore)) return
    if (previous) Object.defineProperty(window, 'pdfApi', previous)
    else Reflect.deleteProperty(window, 'pdfApi')
  }
  activeApiRestorers.add(restore)
  return {
    api,
    emit(next: AiSettings) {
      settingsChanged?.(next)
    },
  }
}

async function waitFor(assertion: () => void): Promise<void> {
  await vi.waitFor(async () => {
    await act(async () => {
      await new Promise<void>((resolve) => queueMicrotask(resolve))
    })
    assertion()
  })
}

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

afterEach(() => {
  for (const cleanup of [...activeCleanups]) cleanup()
  for (const restore of [...activeApiRestorers]) restore()
  document.body.replaceChildren()
})

describe('PDF Codex provider wiring', () => {
  it('loads shared settings, gates signed-out Send, and keeps drafts editable', async () => {
    const initial = settings('openai-codex')
    const { api } = installApi(initial)
    const { container, cleanup } = mount(
      createElement(AiPanel, { api: deps(), onCollapse: () => {} }),
    )
    await waitFor(() => expect(container.querySelector('.ai-provider-select-input')).not.toBeNull())
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
    await waitFor(() => expect(container.querySelector('.ai-codex-model-trigger')).not.toBeNull())
    installed.emit(settings('genspark'))
    await waitFor(() =>
      expect(container.querySelector<HTMLSelectElement>('.ai-provider-select-input')?.value).toBe(
        'genspark',
      ),
    )
    expect(installed.api.setAiSettings).not.toHaveBeenCalled()
    cleanup()
  })
})
