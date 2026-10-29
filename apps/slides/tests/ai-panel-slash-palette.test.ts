// Slides: the composer's slash palette — typing `/` opens the skills listbox
// above the textarea, and picking attaches the skill while striking the token
// from the draft (an already-attached skill is never toggled off).
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

/** Simulate typing into React's controlled textarea (the caret lands at the end) */
function typeInto(textarea: HTMLTextAreaElement, text: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
  act(() => {
    setter.call(textarea, text)
    textarea.setSelectionRange(text.length, text.length)
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function press(textarea: HTMLTextAreaElement, key: string) {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })
  act(() => textarea.dispatchEvent(event))
  return event
}

async function settle(): Promise<void> {
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
  })
}

const chipNames = (container: HTMLElement) =>
  [...container.querySelectorAll('.ai-skill-chip-name')].map((n) => n.textContent)

beforeAll(() => {
  // jsdom has no scrollTo; the panel auto-scrolls its chat log
  Element.prototype.scrollTo ??= () => {}
})

describe('AiPanel slash palette (slides)', () => {
  it('opens on a leading slash; picking attaches the skill, strips the token, and refocuses', async () => {
    Object.defineProperty(window, 'slidesApi', {
      configurable: true,
      value: {
        aiGskStatus: vi.fn(async () => ({ loggedIn: true })),
        beginHistoryBatch: vi.fn(async () => false),
        endHistoryBatch: vi.fn(async () => null),
        aiLogRunFailure: vi.fn(async () => undefined),
        skillsList: vi.fn(async () => [{ name: 'tables', description: 'Header rows and totals' }]),
      },
    })
    const { container, cleanup } = mount(createElement(AiPanel, panelProps()))
    const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!

    expect(container.querySelector('.ai-slash-pop')).toBeNull()

    typeInto(textarea, '/tab')
    await settle()
    const pop = container.querySelector<HTMLElement>('.ai-slash-pop')!
    expect(pop.getAttribute('role')).toBe('listbox')
    expect([...pop.querySelectorAll('.ai-slash-skill-name')].map((n) => n.textContent)).toEqual([
      'tables',
    ])

    const event = press(textarea, 'Enter')
    expect(event.defaultPrevented).toBe(true)
    expect(textarea.value).toBe('')
    expect(container.querySelector('.ai-slash-pop')).toBeNull()
    expect(chipNames(container)).toEqual(['tables'])

    // the composer keeps focus with the caret where the token began
    await act(async () => {
      await new Promise((r) => setTimeout(r, 30))
    })
    expect(document.activeElement).toBe(textarea)
    expect(textarea.selectionStart).toBe(0)

    // picking the same skill again leaves it attached exactly once
    typeInto(textarea, '/tab')
    await settle()
    press(textarea, 'Enter')
    expect(chipNames(container)).toEqual(['tables'])

    cleanup()
  })
})
