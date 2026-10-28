// Slides: the AI panel follows the shared responsive width rule — the
// window-following auto width (floor 360) until the user drags, the clamped
// dragged width afterwards — and drops the legacy persisted width key.
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
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

import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((p) => [p.id, { apiKey: '', model: p.defaultModel }]),
  ) as AiSettings['providers'],
}

function setInnerWidth(w: number): void {
  Object.defineProperty(window, 'innerWidth', { configurable: true, writable: true, value: w })
}

function resizeTo(w: number): void {
  setInnerWidth(w)
  act(() => {
    window.dispatchEvent(new Event('resize'))
  })
}

/** jsdom has no PointerEvent; the resizer only reads pointerId/clientX */
class FakePointerEvent extends Event {
  constructor(type: string, init: { clientX?: number; pointerId?: number } = {}) {
    super(type, { bubbles: true })
    Object.assign(this, { clientX: init.clientX ?? 0, pointerId: init.pointerId ?? 1 })
  }
}

const mounted: Array<{ root: Root; dock: HTMLElement }> = []

async function mountPanel(): Promise<HTMLElement> {
  const { AiPanel } = await import('../src/renderer/ai/AiPanel')
  const dock = document.createElement('div')
  dock.className = 'ai-dock'
  document.body.appendChild(dock)
  const root = createRoot(dock)
  act(() =>
    root.render(
      createElement(AiPanel, {
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
      }),
    ),
  )
  mounted.push({ root, dock })
  return dock
}

/** Drag the resizer: pointerdown on the handle, move/up on the window */
function dragTo(dock: HTMLElement, pointerId: number, clientX: number): void {
  const resizer = dock.querySelector<HTMLElement>('.ai-panel-resizer')
  expect(resizer).not.toBeNull()
  act(() => {
    resizer!.dispatchEvent(new FakePointerEvent('pointerdown', { clientX: 360, pointerId }))
    window.dispatchEvent(new FakePointerEvent('pointermove', { clientX, pointerId }))
    window.dispatchEvent(new FakePointerEvent('pointerup', { clientX, pointerId }))
  })
}

beforeAll(() => {
  // jsdom lacks scrollTo (the panel auto-scrolls its chat log) and pointer capture
  Element.prototype.scrollTo ??= () => {}
  Element.prototype.setPointerCapture ??= () => {}
  Element.prototype.releasePointerCapture ??= () => {}
})

afterEach(() => {
  for (const { root, dock } of mounted.splice(0)) {
    act(() => root.unmount())
    dock.remove()
  }
})

describe('AiPanel width rule (slides)', () => {
  it('drops the legacy persisted width key so pre-floor values never come back', async () => {
    localStorage.setItem('slides-ai-panel-width', '240')
    await import('../src/renderer/ai/AiPanel')
    expect(localStorage.getItem('slides-ai-panel-width')).toBeNull()
  })

  it('shows the auto width before any drag and re-derives it on resize', async () => {
    setInnerWidth(1920)
    const dock = await mountPanel()
    expect(dock.style.getPropertyValue('--ai-panel-width')).toBe('422px')
    // narrow window: the width never drops below the 360 floor
    resizeTo(1000)
    expect(dock.style.getPropertyValue('--ai-panel-width')).toBe('360px')
    // wide window: ~22% of the window, still under the 60% cap
    resizeTo(1800)
    expect(dock.style.getPropertyValue('--ai-panel-width')).toBe('396px')
  })

  it('clamps a dragged width and re-clamps it on window resize', async () => {
    setInnerWidth(1000)
    const dock = await mountPanel()
    dragTo(dock, 1, 500)
    expect(dock.style.getPropertyValue('--ai-panel-width')).toBe('500px')
    // never above the 60% cap: shrinking the window re-clamps the dragged width
    resizeTo(800)
    expect(dock.style.getPropertyValue('--ai-panel-width')).toBe('480px')
    // never below 360, however far the drag goes
    dragTo(dock, 2, 100)
    expect(dock.style.getPropertyValue('--ai-panel-width')).toBe('360px')
  })
})
