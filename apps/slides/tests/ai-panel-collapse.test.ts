// Slides: the AI panel stays mounted while collapsed (rail only),
// so the conversation, draft, and in-flight runs survive collapse/expand.
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

// react-konva's node entry requires the native 'canvas' package; nothing here draws
vi.mock('react-konva', () => {
  const stub = () => null
  return {
    Stage: stub,
    Layer: stub,
    Rect: stub,
    Group: stub,
    Transformer: stub,
    Line: stub,
    Arrow: stub,
    Text: stub,
    Ellipse: stub,
    Image: stub,
    Path: stub,
    Circle: stub,
    Arc: stub,
  }
})

import { AiPanel } from '../src/renderer/ai/AiPanel'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((p) => [p.id, { apiKey: '', model: p.defaultModel }]),
  ) as AiSettings['providers'],
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

function panelProps(overrides: Record<string, unknown> = {}) {
  return {
    slides: [],
    current: 0,
    selectedIds: [],
    images: new Map<string, HTMLImageElement>(),
    applySlide: () => {},
    applyDeck: () => {},
    fitWidthPx: 960,
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

beforeAll(() => {
  // jsdom has no scrollTo; the panel auto-scrolls its chat log
  Element.prototype.scrollTo ??= () => {}
})

describe('AiPanel collapse (slides)', () => {
  it('keeps the draft input across a collapse/expand cycle', () => {
    const { container, root, cleanup } = mount(createElement(AiPanel, panelProps()))

    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')
    expect(textarea).not.toBeNull()
    typeInto(textarea!, 'unsent draft')
    expect(textarea!.value).toBe('unsent draft')

    // collapse: only the rail is rendered, but the component stays mounted
    act(() => root.render(createElement(AiPanel, panelProps({ open: false }))))
    expect(container.querySelector('.ai-input-box textarea')).toBeNull()
    expect(container.querySelector('.ai-rail')).not.toBeNull()

    // expand: the draft is still there
    act(() => root.render(createElement(AiPanel, panelProps({ open: true }))))
    const restored = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')
    expect(restored).not.toBeNull()
    expect(restored!.value).toBe('unsent draft')

    cleanup()
  })

  it('expands back through the rail button', () => {
    const onExpand = vi.fn()
    const { container, cleanup } = mount(
      createElement(AiPanel, panelProps({ open: false, onExpand })),
    )

    const rail = container.querySelector<HTMLButtonElement>('.ai-rail')
    expect(rail).not.toBeNull()
    act(() => rail!.click())
    expect(onExpand).toHaveBeenCalledTimes(1)

    cleanup()
  })

  it('renders Codex controls and keeps Send disabled while signed out', async () => {
    const previous = Object.getOwnPropertyDescriptor(window, 'slidesApi')
    const api = {
      aiCodexStatus: vi.fn(async () => ({ loggedIn: false })),
      aiCodexLogin: vi.fn(async () => ({ loggedIn: true })),
      aiCodexCancelLogin: vi.fn(async () => {}),
      aiCodexLogout: vi.fn(async () => ({ loggedIn: false })),
      aiCodexCapabilities: vi.fn(async () => ({ models: [] })),
      setAiSettings: vi.fn(async () => {}),
      onAiStream: vi.fn(() => () => {}),
      aiStream: vi.fn(async () => {}),
      aiStreamCancel: vi.fn(async () => {}),
    }
    Object.defineProperty(window, 'slidesApi', { configurable: true, value: api })
    const onSettingsChange = vi.fn()
    const codexSettings: AiSettings = {
      ...settings,
      provider: 'openai-codex',
      providers: {
        ...settings.providers,
        'openai-codex': {
          ...settings.providers['openai-codex'],
          reasoningEffort: 'none',
          serviceTier: 'default',
        },
      },
    }
    const { container, cleanup } = mount(
      createElement(AiPanel, panelProps({ settings: codexSettings, onSettingsChange })),
    )
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(container.querySelectorAll('.ai-provider-select-input option')).toHaveLength(2)
    expect(container.querySelector<HTMLButtonElement>('.ai-send-btn')?.disabled).toBe(true)
    cleanup()
    if (previous) Object.defineProperty(window, 'slidesApi', previous)
    else delete (window as Partial<Window>).slidesApi
  })

  it('renders the capability-backed model control after Codex sign-in', async () => {
    const previous = Object.getOwnPropertyDescriptor(window, 'slidesApi')
    const api = {
      aiCodexStatus: vi.fn(async () => ({ loggedIn: true })),
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
      setAiSettings: vi.fn(async () => {}),
      onAiStream: vi.fn(() => () => {}),
      aiStream: vi.fn(async () => {}),
      aiStreamCancel: vi.fn(async () => {}),
    }
    Object.defineProperty(window, 'slidesApi', { configurable: true, value: api })
    const codexSettings: AiSettings = {
      ...settings,
      provider: 'openai-codex',
      providers: {
        ...settings.providers,
        'openai-codex': {
          ...settings.providers['openai-codex'],
          reasoningEffort: 'none',
          serviceTier: 'default',
        },
      },
    }
    const { container, cleanup } = mount(
      createElement(AiPanel, panelProps({ settings: codexSettings })),
    )
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(container.querySelector('.ai-codex-model-trigger')).not.toBeNull()
    cleanup()
    if (previous) Object.defineProperty(window, 'slidesApi', previous)
    else delete (window as Partial<Window>).slidesApi
  })
})
