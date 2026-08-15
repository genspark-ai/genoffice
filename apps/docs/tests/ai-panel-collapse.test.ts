// The AI panel stays mounted while collapsed (rail only),
// so the conversation, draft, and in-flight runs survive collapse/expand.
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((p) => [p.id, { apiKey: '', model: p.defaultModel }]),
  ) as AiSettings['providers'],
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

function panelProps(editor: Editor, overrides: Record<string, unknown> = {}) {
  return {
    editor,
    blocks: [],
    settings,
    open: true,
    onExpand: () => {},
    onCollapse: () => {},
    ...overrides,
  }
}

/** Simulate typing into React's controlled textarea */
function typeInto(textarea: HTMLTextAreaElement, text: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
  act(() => {
    setter.call(textarea, text)
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function deferred<T>() {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((res) => {
    resolve = res
  })
  return { promise, resolve }
}

function installDesktop(desktop: Record<string, unknown>) {
  const previous = Object.getOwnPropertyDescriptor(window, 'desktop')
  Object.defineProperty(window, 'desktop', { configurable: true, value: desktop })
  return () => {
    if (previous) Object.defineProperty(window, 'desktop', previous)
    else delete (window as Partial<Window>).desktop
  }
}

beforeAll(() => {
  // jsdom has no scrollTo; the panel auto-scrolls its chat log
  Element.prototype.scrollTo ??= () => {}
})

describe('AiPanel collapse', () => {
  it('orders provider, new conversation, and collapse controls', async () => {
    const editor = createEditor()
    const restoreDesktop = installDesktop({
      onAiStream: () => () => {},
      aiStream: vi.fn().mockResolvedValue(undefined),
      aiStreamCancel: vi.fn().mockResolvedValue(undefined),
    })
    const { container, cleanup } = mount(
      createElement(
        AiPanel,
        panelProps(editor, { preset: { text: 'Start conversation', nonce: 1, autoRun: true } }),
      ),
    )
    await act(async () => {
      await Promise.resolve()
    })
    const actions = container.querySelector('.ai-panel-header-actions')!
    const provider = actions.querySelector('.ai-provider-select')
    const headerButtons = [...actions.querySelectorAll<HTMLButtonElement>('.ai-header-btn')]

    expect(headerButtons).toHaveLength(2)
    expect(provider?.nextElementSibling).toBe(headerButtons[0])
    expect(headerButtons[0]?.nextElementSibling).toBe(headerButtons[1])

    cleanup()
    restoreDesktop()
    editor.destroy()
  })

  it('keeps model and reasoning controls hidden for non-Codex providers', () => {
    const editor = createEditor()
    const { container, cleanup } = mount(createElement(AiPanel, panelProps(editor)))

    expect(container.querySelectorAll('.ai-provider-select')).toHaveLength(1)

    cleanup()
    editor.destroy()
  })

  it('keeps the product title fixed and moves Codex controls into the composer', async () => {
    const editor = createEditor()
    let streamListener:
      ((chunk: { requestId: string; type: 'error'; error: string }) => void) | undefined
    const previousDesktop = Object.getOwnPropertyDescriptor(window, 'desktop')
    const onSettingsChange = vi.fn()
    Object.defineProperty(window, 'desktop', {
      configurable: true,
      value: {
        onAiStream: (listener: typeof streamListener) => {
          streamListener = listener
          return () => {}
        },
        aiStream: vi.fn().mockResolvedValue(undefined),
        aiStreamCancel: vi.fn().mockResolvedValue(undefined),
        aiCodexCapabilities: vi.fn().mockResolvedValue({
          models: [
            { id: 'gpt-5.5', reasoningEfforts: ['none', 'low', 'high'] },
            { id: 'gpt-5.4', reasoningEfforts: ['none', 'medium'] },
          ],
        }),
        aiCodexStatus: vi.fn().mockResolvedValue({ loggedIn: true }),
      },
    })
    const codexSettings: AiSettings = { ...settings, provider: 'openai-codex' }
    const { container, cleanup } = mount(
      createElement(AiPanel, panelProps(editor, { settings: codexSettings, onSettingsChange })),
    )

    await act(async () => {
      await Promise.resolve()
    })

    expect(container.querySelector('.ai-panel-title')?.textContent).toContain('GenSpark AI')
    expect(container.querySelectorAll('.ai-panel-header .ai-provider-select')).toHaveLength(1)
    expect(container.querySelector('.ai-codex-model-trigger')).not.toBeNull()

    act(() => container.querySelector<HTMLButtonElement>('.ai-codex-model-trigger')!.click())
    expect(container.querySelector('.ai-codex-model-popover')).not.toBeNull()
    expect(container.querySelectorAll('[data-codex-model-option]')).toHaveLength(2)
    expect(container.querySelectorAll('[data-codex-reasoning-option]')).toHaveLength(3)

    act(() => container.querySelector<HTMLElement>('[data-codex-model-option="gpt-5.4"]')!.click())
    expect(onSettingsChange).toHaveBeenLastCalledWith(
      expect.objectContaining({
        providers: expect.objectContaining({
          'openai-codex': expect.objectContaining({ model: 'gpt-5.4' }),
        }),
      }),
    )
    act(() => container.querySelector<HTMLElement>('[data-codex-reasoning-option="high"]')!.click())
    expect(onSettingsChange).toHaveBeenLastCalledWith(
      expect.objectContaining({
        providers: expect.objectContaining({
          'openai-codex': expect.objectContaining({ reasoningEffort: 'high' }),
        }),
      }),
    )

    cleanup()
    if (previousDesktop) Object.defineProperty(window, 'desktop', previousDesktop)
    else delete (window as Partial<Window>).desktop
    editor.destroy()
    void streamListener
  })

  it('opens explanatory Track changes choices and persists the selected mode', () => {
    const editor = createEditor()
    localStorage.removeItem('ai-docs-track-changes')
    const { container, cleanup } = mount(createElement(AiPanel, panelProps(editor)))
    const trigger = container.querySelector<HTMLButtonElement>('.ai-track-btn')!

    act(() => trigger.click())
    expect(container.querySelector('.ai-track-popover')).not.toBeNull()
    expect(container.querySelectorAll('[data-track-choice]')).toHaveLength(2)
    expect(container.querySelector('[data-track-choice="on"]')?.textContent).toContain(
      '修订追踪已开启',
    )

    act(() => container.querySelector<HTMLElement>('[data-track-choice="on"]')!.click())
    expect(trigger.classList.contains('on')).toBe(true)
    expect(localStorage.getItem('ai-docs-track-changes')).toBe('1')

    act(() => trigger.click())
    act(() => container.querySelector<HTMLElement>('[data-track-choice="off"]')!.click())
    expect(trigger.classList.contains('on')).toBe(false)
    expect(localStorage.getItem('ai-docs-track-changes')).toBe('0')

    act(() => trigger.click())
    act(() => document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' })))
    expect(container.querySelector('.ai-track-popover')).toBeNull()

    act(() => trigger.click())
    act(() => document.body.dispatchEvent(new Event('pointerdown', { bubbles: true })))
    expect(container.querySelector('.ai-track-popover')).toBeNull()

    cleanup()
    editor.destroy()
  })

  it('preflights Codex auth before enabling send or loading capabilities', async () => {
    const editor = createEditor()
    const status = deferred<{ loggedIn: boolean }>()
    const aiStream = vi.fn()
    const aiCodexCapabilities = vi.fn().mockResolvedValue({
      models: [{ id: 'gpt-5.5', reasoningEfforts: ['none', 'low', 'high'] }],
    })
    const restoreDesktop = installDesktop({
      onAiStream: () => () => {},
      aiStream,
      aiStreamCancel: vi.fn().mockResolvedValue(undefined),
      aiCodexStatus: vi.fn().mockReturnValue(status.promise),
      aiCodexCapabilities,
    })
    const codexSettings: AiSettings = { ...settings, provider: 'openai-codex' }
    const { container, cleanup } = mount(
      createElement(AiPanel, panelProps(editor, { settings: codexSettings })),
    )

    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
    typeInto(textarea, 'Summarize this document')
    expect(container.querySelector<HTMLButtonElement>('.ai-send-btn')!.disabled).toBe(true)
    act(() => {
      textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    })
    expect(aiStream).not.toHaveBeenCalled()
    expect(aiCodexCapabilities).not.toHaveBeenCalled()

    status.resolve({ loggedIn: false })
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })

    const header = container.querySelector('.ai-panel-header')!
    const banner = container.querySelector('.ai-codex-auth-banner')!
    expect(header.nextElementSibling).toBe(banner)
    expect(banner.nextElementSibling).toBe(container.querySelector('.ai-chat'))
    expect(banner.textContent).toContain('登录 ChatGPT 后才能使用 ChatGPT Codex。')
    expect(banner.querySelectorAll('button')).toHaveLength(1)
    expect(container.querySelector<HTMLButtonElement>('.ai-send-btn')!.disabled).toBe(true)
    expect(aiStream).not.toHaveBeenCalled()
    cleanup()
    restoreDesktop()
    editor.destroy()
  })

  it('keeps draft while Codex login is pending and enables send after success', async () => {
    const editor = createEditor()
    const login = deferred<{ loggedIn: boolean }>()
    const aiCodexLogin = vi.fn().mockReturnValue(login.promise)
    const aiCodexCapabilities = vi.fn().mockResolvedValue({
      models: [{ id: 'gpt-5.5', reasoningEfforts: ['none', 'low', 'high'] }],
    })
    const restoreDesktop = installDesktop({
      onAiStream: () => () => {},
      aiStream: vi.fn(),
      aiStreamCancel: vi.fn().mockResolvedValue(undefined),
      aiCodexStatus: vi.fn().mockResolvedValue({ loggedIn: false }),
      aiCodexLogin,
      aiCodexCapabilities,
    })
    const codexSettings: AiSettings = { ...settings, provider: 'openai-codex' }
    const { container, cleanup } = mount(
      createElement(AiPanel, panelProps(editor, { settings: codexSettings })),
    )

    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
    typeInto(textarea, 'Keep this draft')
    const loginButton = container.querySelector<HTMLButtonElement>('.ai-codex-auth-login')!

    act(() => loginButton.click())
    expect(aiCodexLogin).toHaveBeenCalledTimes(1)
    expect(loginButton.disabled).toBe(true)
    expect(loginButton.getAttribute('aria-busy')).toBe('true')
    act(() => loginButton.click())
    expect(aiCodexLogin).toHaveBeenCalledTimes(1)

    login.resolve({ loggedIn: true })
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(container.querySelector('.ai-codex-auth-banner')).toBeNull()
    expect(container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!.value).toBe(
      'Keep this draft',
    )
    expect(aiCodexCapabilities).toHaveBeenCalledTimes(1)
    expect(container.querySelector<HTMLButtonElement>('.ai-send-btn')!.disabled).toBe(false)

    cleanup()
    restoreDesktop()
    editor.destroy()
  })

  it('keeps Codex sign-in failure visible and retryable', async () => {
    const editor = createEditor()
    const aiCodexLogin = vi
      .fn()
      .mockResolvedValueOnce({ loggedIn: false, errorCode: 'auth-temporary' as const })
      .mockResolvedValueOnce({ loggedIn: false, errorCode: 'auth-temporary' as const })
    const restoreDesktop = installDesktop({
      onAiStream: () => () => {},
      aiStream: vi.fn(),
      aiStreamCancel: vi.fn().mockResolvedValue(undefined),
      aiCodexStatus: vi.fn().mockResolvedValue({ loggedIn: false }),
      aiCodexLogin,
      aiCodexCapabilities: vi.fn().mockResolvedValue({ models: [] }),
    })
    const codexSettings: AiSettings = { ...settings, provider: 'openai-codex' }
    const { container, cleanup } = mount(
      createElement(AiPanel, panelProps(editor, { settings: codexSettings })),
    )

    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
    typeInto(textarea, 'Draft survives failure')
    const loginButton = container.querySelector<HTMLButtonElement>('.ai-codex-auth-login')!
    act(() => loginButton.click())
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(container.querySelector('.ai-codex-auth-banner')?.textContent).toContain(
      'ChatGPT 登录暂时不可用',
    )
    expect(container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!.value).toBe(
      'Draft survives failure',
    )
    expect(container.querySelector<HTMLButtonElement>('.ai-codex-auth-login')!.disabled).toBe(false)

    act(() => container.querySelector<HTMLButtonElement>('.ai-codex-auth-login')!.click())
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(aiCodexLogin).toHaveBeenCalledTimes(2)

    cleanup()
    restoreDesktop()
    editor.destroy()
  })

  it('switches away from Codex without restoring stale login state', async () => {
    const editor = createEditor()
    const login = deferred<{ loggedIn: boolean }>()
    const aiCodexStatus = vi.fn().mockResolvedValue({ loggedIn: false })
    const aiCodexLogin = vi.fn().mockReturnValue(login.promise)
    const aiStream = vi.fn().mockResolvedValue(undefined)
    const restoreDesktop = installDesktop({
      onAiStream: () => () => {},
      aiStream,
      aiStreamCancel: vi.fn().mockResolvedValue(undefined),
      aiCodexStatus,
      aiCodexLogin,
      aiCodexCapabilities: vi.fn().mockResolvedValue({ models: [] }),
    })
    const codexSettings: AiSettings = { ...settings, provider: 'openai-codex' }
    const { container, root, cleanup } = mount(
      createElement(AiPanel, panelProps(editor, { settings: codexSettings })),
    )

    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    act(() => container.querySelector<HTMLButtonElement>('.ai-codex-auth-login')!.click())
    const anthropicSettings: AiSettings = { ...settings, provider: 'anthropic' }
    act(() =>
      root.render(createElement(AiPanel, panelProps(editor, { settings: anthropicSettings }))),
    )
    expect(container.querySelector('.ai-codex-auth-banner')).toBeNull()

    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
    typeInto(textarea, 'Anthropic draft')
    act(() => container.querySelector<HTMLButtonElement>('.ai-send-btn')!.click())
    await act(async () => {
      await Promise.resolve()
    })
    expect(aiStream).toHaveBeenCalledTimes(1)

    login.resolve({ loggedIn: true })
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(container.querySelector('.ai-codex-auth-banner')).toBeNull()

    act(() => root.render(createElement(AiPanel, panelProps(editor, { settings: codexSettings }))))
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(aiCodexStatus).toHaveBeenCalledTimes(2)
    expect(container.querySelector('.ai-codex-auth-banner')).not.toBeNull()

    cleanup()
    restoreDesktop()
    editor.destroy()
  })

  it('does not show Codex login action for a non-auth stream error', async () => {
    const editor = createEditor()
    let streamListener:
      | ((chunk: {
          requestId: string
          type: 'error'
          errorCode: 'capabilities-unavailable'
        }) => void)
      | undefined
    const aiStream = vi.fn((request: { requestId: string }) => {
      queueMicrotask(() =>
        streamListener?.({
          requestId: request.requestId,
          type: 'error',
          errorCode: 'capabilities-unavailable',
        }),
      )
      return Promise.resolve()
    })
    const onSettingsChange = vi.fn()
    const restoreDesktop = installDesktop({
      onAiStream: (listener: typeof streamListener) => {
        streamListener = listener
        return () => {}
      },
      aiStream,
      aiStreamCancel: vi.fn().mockResolvedValue(undefined),
      aiCodexStatus: vi.fn().mockResolvedValue({ loggedIn: true }),
      aiCodexCapabilities: vi.fn().mockResolvedValue({
        models: [
          { id: 'gpt-5.5', reasoningEfforts: ['none', 'low', 'high'] },
          { id: 'gpt-5.4', reasoningEfforts: ['none', 'medium'] },
        ],
      }),
    })
    const codexSettings: AiSettings = { ...settings, provider: 'openai-codex' }
    const { container, cleanup } = mount(
      createElement(AiPanel, panelProps(editor, { settings: codexSettings, onSettingsChange })),
    )

    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
      await Promise.resolve()
    })
    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
    typeInto(textarea, 'Summarize this document')
    act(() => container.querySelector<HTMLButtonElement>('.ai-send-btn')!.click())
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 0))
    })

    expect(container.querySelector('.ai-msg-error')?.textContent).toContain('Codex 模型暂不可用')
    expect(container.querySelector('.ai-login-btn')).toBeNull()

    cleanup()
    restoreDesktop()
    editor.destroy()
  })

  it('keeps the draft input across a collapse/expand cycle', () => {
    const editor = createEditor()
    const { container, root, cleanup } = mount(createElement(AiPanel, panelProps(editor)))

    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')
    expect(textarea).not.toBeNull()
    typeInto(textarea!, 'unsent draft')
    expect(textarea!.value).toBe('unsent draft')

    // collapse: only the rail is rendered, but the component stays mounted
    act(() => root.render(createElement(AiPanel, panelProps(editor, { open: false }))))
    expect(container.querySelector('.ai-input-box textarea')).toBeNull()
    expect(container.querySelector('.ai-rail')).not.toBeNull()

    // expand: the draft is still there
    act(() => root.render(createElement(AiPanel, panelProps(editor, { open: true }))))
    const restored = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')
    expect(restored).not.toBeNull()
    expect(restored!.value).toBe('unsent draft')

    cleanup()
    editor.destroy()
  })

  it('expands back through the rail button', () => {
    const editor = createEditor()
    const onExpand = vi.fn()
    const { container, cleanup } = mount(
      createElement(AiPanel, panelProps(editor, { open: false, onExpand })),
    )

    const rail = container.querySelector<HTMLButtonElement>('.ai-rail')
    expect(rail).not.toBeNull()
    act(() => rail!.click())
    expect(onExpand).toHaveBeenCalledTimes(1)

    cleanup()
    editor.destroy()
  })
})
