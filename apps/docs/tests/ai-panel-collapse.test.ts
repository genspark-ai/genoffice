// The AI panel stays mounted while collapsed (rail only),
// so the conversation, draft, and in-flight runs survive collapse/expand.
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
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

const activeCleanups = new Set<() => void>()
const activeDesktopRestorers = new Set<() => void>()

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

async function waitFor(assertion: () => void): Promise<void> {
  await vi.waitFor(async () => {
    await act(async () => {
      await new Promise<void>((resolve) => queueMicrotask(resolve))
    })
    assertion()
  })
}

function installDesktop(desktop: Record<string, unknown>) {
  const previous = Object.getOwnPropertyDescriptor(window, 'desktop')
  Object.defineProperty(window, 'desktop', { configurable: true, value: desktop })
  const restore = () => {
    if (!activeDesktopRestorers.delete(restore)) return
    if (previous) Object.defineProperty(window, 'desktop', previous)
    else Reflect.deleteProperty(window, 'desktop')
  }
  activeDesktopRestorers.add(restore)
  return restore
}

beforeAll(() => {
  // jsdom has no scrollTo; the panel auto-scrolls its chat log
  Element.prototype.scrollTo ??= () => {}
})

afterEach(() => {
  for (const cleanup of [...activeCleanups]) cleanup()
  for (const restore of [...activeDesktopRestorers]) restore()
  Reflect.deleteProperty(window, 'desktop')
  document.body.replaceChildren()
})

describe('AiPanel collapse', () => {
  it('updates settings when selecting Codex model, reasoning, and service tier', async () => {
    const editor = createEditor()
    let streamListener:
      ((chunk: { requestId: string; type: 'error'; error: string }) => void) | undefined
    const onSettingsChange = vi.fn()
    const restoreDesktop = installDesktop({
      onAiStream: (listener: typeof streamListener) => {
        streamListener = listener
        return () => {}
      },
      aiStream: vi.fn().mockResolvedValue(undefined),
      aiStreamCancel: vi.fn().mockResolvedValue(undefined),
      aiCodexCapabilities: vi.fn().mockResolvedValue({
        models: [
          {
            id: 'gpt-5.5',
            reasoningEfforts: ['none', 'low', 'high'],
            serviceTiers: [
              { id: 'default', name: 'Standard' },
              { id: 'priority', name: 'Fast' },
            ],
          },
          {
            id: 'gpt-5.4',
            reasoningEfforts: ['none', 'medium'],
            serviceTiers: [{ id: 'default', name: 'Standard' }],
          },
        ],
      }),
      aiCodexStatus: vi.fn().mockResolvedValue({ loggedIn: true }),
    })
    const codexSettings: AiSettings = { ...settings, provider: 'openai-codex' }
    const { container, cleanup } = mount(
      createElement(AiPanel, panelProps(editor, { settings: codexSettings, onSettingsChange })),
    )

    await waitFor(() => expect(container.querySelector('.ai-codex-model-trigger')).not.toBeNull())

    expect(container.querySelector('.ai-codex-model-trigger')).not.toBeNull()

    act(() => container.querySelector<HTMLButtonElement>('.ai-codex-model-trigger')!.click())
    expect(container.querySelector('.ai-codex-model-popover')).not.toBeNull()

    act(() => container.querySelector<HTMLElement>('[data-codex-model-option="gpt-5.4"]')!.click())
    expect(onSettingsChange).toHaveBeenLastCalledWith(
      expect.objectContaining({
        providers: expect.objectContaining({
          'openai-codex': expect.objectContaining({ model: 'gpt-5.4' }),
        }),
      }),
    )

    act(() => container.querySelector<HTMLButtonElement>('.ai-codex-model-trigger')!.click())
    act(() =>
      container.querySelector<HTMLButtonElement>('[data-codex-menu-item="effort"]')!.click(),
    )
    act(() => container.querySelector<HTMLElement>('[data-codex-reasoning-option="high"]')!.click())
    expect(onSettingsChange).toHaveBeenLastCalledWith(
      expect.objectContaining({
        providers: expect.objectContaining({
          'openai-codex': expect.objectContaining({ reasoningEffort: 'high' }),
        }),
      }),
    )

    act(() => container.querySelector<HTMLButtonElement>('.ai-codex-model-trigger')!.click())
    act(() => container.querySelector<HTMLButtonElement>('[data-codex-menu-item="speed"]')!.click())
    act(() =>
      container.querySelector<HTMLElement>('[data-codex-service-tier-option="priority"]')!.click(),
    )
    expect(onSettingsChange).toHaveBeenLastCalledWith(
      expect.objectContaining({
        providers: expect.objectContaining({
          'openai-codex': expect.objectContaining({ serviceTier: 'priority' }),
        }),
      }),
    )

    cleanup()
    restoreDesktop()
    editor.destroy()
    void streamListener
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
    await waitFor(() => expect(container.querySelector('.ai-codex-auth-banner')).not.toBeNull())

    expect(container.querySelector('.ai-codex-auth-banner')).not.toBeNull()
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

    await waitFor(() => expect(container.querySelector('.ai-codex-auth-login')).not.toBeNull())
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
    await waitFor(() => {
      expect(container.querySelector('.ai-codex-auth-banner')).toBeNull()
      expect(aiCodexCapabilities).toHaveBeenCalledTimes(1)
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

    await waitFor(() => expect(container.querySelector('.ai-codex-auth-login')).not.toBeNull())
    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
    typeInto(textarea, 'Draft survives failure')
    const loginButton = container.querySelector<HTMLButtonElement>('.ai-codex-auth-login')!
    act(() => loginButton.click())
    await waitFor(() => expect(container.querySelector('.ai-codex-auth-banner')).not.toBeNull())
    expect(container.querySelector('.ai-codex-auth-banner')).not.toBeNull()
    expect(container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!.value).toBe(
      'Draft survives failure',
    )
    expect(container.querySelector<HTMLButtonElement>('.ai-codex-auth-login')!.disabled).toBe(false)

    act(() => container.querySelector<HTMLButtonElement>('.ai-codex-auth-login')!.click())
    await waitFor(() => expect(aiCodexLogin).toHaveBeenCalledTimes(2))
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

    await waitFor(() => expect(container.querySelector('.ai-codex-auth-login')).not.toBeNull())
    act(() => container.querySelector<HTMLButtonElement>('.ai-codex-auth-login')!.click())
    const anthropicSettings: AiSettings = { ...settings, provider: 'anthropic' }
    act(() =>
      root.render(createElement(AiPanel, panelProps(editor, { settings: anthropicSettings }))),
    )
    expect(container.querySelector('.ai-codex-auth-banner')).toBeNull()

    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
    typeInto(textarea, 'Anthropic draft')
    act(() => container.querySelector<HTMLButtonElement>('.ai-send-btn')!.click())
    await waitFor(() => expect(aiStream).toHaveBeenCalledTimes(1))
    expect(aiStream).toHaveBeenCalledTimes(1)

    login.resolve({ loggedIn: true })
    await waitFor(() => expect(container.querySelector('.ai-codex-auth-banner')).toBeNull())
    expect(container.querySelector('.ai-codex-auth-banner')).toBeNull()

    act(() => root.render(createElement(AiPanel, panelProps(editor, { settings: codexSettings }))))
    await waitFor(() => {
      expect(aiCodexStatus).toHaveBeenCalledTimes(2)
      expect(container.querySelector('.ai-codex-auth-banner')).not.toBeNull()
    })

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

    await waitFor(() => expect(container.querySelector('.ai-codex-model-trigger')).not.toBeNull())
    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
    typeInto(textarea, 'Summarize this document')
    act(() => container.querySelector<HTMLButtonElement>('.ai-send-btn')!.click())
    await waitFor(() => expect(container.querySelector('.ai-msg-error')).not.toBeNull())

    expect(container.querySelector('.ai-msg-error')).not.toBeNull()
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
