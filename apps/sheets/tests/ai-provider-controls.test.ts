// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { defaultAiSettings, type AiSettings } from '@genoffice/ai-provider'
import { AiChatPanel } from '../src/renderer/ai/AiChatPanel'

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

function installApi({ loggedIn = false, capabilities = false } = {}) {
  const api = {
    aiCodexStatus: vi.fn(async () => ({ loggedIn })),
    aiCodexLogin: vi.fn(async () => ({ loggedIn: true })),
    aiCodexCancelLogin: vi.fn(async () => {}),
    aiCodexLogout: vi.fn(async () => ({ loggedIn: false })),
    aiCodexCapabilities: vi.fn(async () =>
      capabilities
        ? {
            models: [
              {
                id: 'gpt-5.5',
                name: 'GPT-5.5',
                reasoningEfforts: ['none', 'high'],
                serviceTiers: [{ id: 'default', name: 'Standard' }],
              },
            ],
          }
        : { models: [] },
    ),
    setAiSettings: vi.fn(async () => {}),
    readAttachmentImage: vi.fn(async () => ({ ok: false })),
  }
  Object.defineProperty(window, 'desktopApi', { configurable: true, value: api })
  return api
}

function props(overrides: Record<string, unknown> = {}) {
  return {
    isOpen: true,
    hasContent: true,
    chat: [],
    historicChat: [],
    attachments: [],
    attachNotice: null,
    onPickAttachments: () => {},
    onAddAttachmentPaths: () => {},
    onAddPastedImage: () => {},
    onRemoveAttachment: () => {},
    prompt: '',
    preview: null,
    aiBusy: false,
    onPromptChange: () => {},
    onSend: () => {},
    onStop: () => {},
    onNewChat: () => {},
    onUndo: () => {},
    onExpand: () => {},
    onCollapse: () => {},
    settings: settings('genspark'),
    onSettingsChange: () => {},
    ...overrides,
  }
}

afterEach(() => {
  document.body.replaceChildren()
})

beforeAll(() => {
  ;(
    globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
  ).IS_REACT_ACT_ENVIRONMENT = true
  Element.prototype.scrollTo ??= () => {}
})

describe('Sheets Codex provider wiring', () => {
  it('persists provider selection', () => {
    const api = installApi()
    const onSettingsChange = vi.fn()
    const { container, cleanup } = mount(createElement(AiChatPanel, props({ onSettingsChange })))
    const select = container.querySelector<HTMLSelectElement>('.ai-provider-select-input')!
    act(() => {
      select.value = 'openai-codex'
      select.dispatchEvent(new Event('change', { bubbles: true }))
    })
    expect(onSettingsChange).toHaveBeenLastCalledWith(
      expect.objectContaining({ provider: 'openai-codex' }),
    )
    expect(api.setAiSettings).not.toHaveBeenCalled()
    cleanup()
  })

  it('keeps the draft editable while signed-out Codex disables Send', async () => {
    installApi()
    const onPromptChange = vi.fn()
    const { container, cleanup } = mount(
      createElement(AiChatPanel, props({ settings: settings('openai-codex'), onPromptChange })),
    )
    await act(async () => {
      await Promise.resolve()
    })
    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
    act(() => {
      setter.call(textarea, 'draft')
      textarea.dispatchEvent(new Event('input', { bubbles: true }))
    })
    expect(onPromptChange).toHaveBeenCalledWith('draft')
    expect(container.querySelector<HTMLButtonElement>('.ai-send-btn')?.disabled).toBe(true)
    cleanup()
  })

  it('renders the capability-backed model control without echoing external settings', async () => {
    const api = installApi({ loggedIn: true, capabilities: true })
    const { container, root, cleanup } = mount(
      createElement(AiChatPanel, props({ settings: settings('openai-codex') })),
    )
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(container.querySelector('.ai-codex-model-trigger')).not.toBeNull()
    const next = settings('genspark')
    act(() => root.render(createElement(AiChatPanel, props({ settings: next }))))
    expect(container.querySelector<HTMLSelectElement>('.ai-provider-select-input')?.value).toBe(
      'genspark',
    )
    expect(api.setAiSettings).not.toHaveBeenCalled()
    cleanup()
  })
})
