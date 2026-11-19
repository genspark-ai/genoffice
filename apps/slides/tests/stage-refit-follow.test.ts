// Regression: the stage's fit-to-window follow (ResizeObserver on .stage-wrap)
// must survive an editor remount — switching to reading/sorter view and back
// replaces the .stage-wrap element, and the observer has to re-bind to the new
// node or window resizes stop re-fitting the canvas (slide overflows the pane).
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

import { App } from '../src/renderer/App'

/** Recording ResizeObserver: lets the test fire size changes for a target. */
class FakeResizeObserver {
  static instances: FakeResizeObserver[] = []
  targets = new Set<Element>()
  constructor(private cb: ResizeObserverCallback) {
    FakeResizeObserver.instances.push(this)
  }
  observe(el: Element) {
    this.targets.add(el)
    // Real ResizeObserver delivers an initial entry on observe; the stage
    // follow relies on it to baseline the container's border-box size.
    this.cb([{ target: el } as ResizeObserverEntry], this as unknown as ResizeObserver)
  }
  unobserve(el: Element) {
    this.targets.delete(el)
  }
  disconnect() {
    this.targets.clear()
  }
  static fire(el: Element) {
    for (const inst of FakeResizeObserver.instances) {
      if (inst.targets.has(el)) {
        inst.cb([{ target: el } as ResizeObserverEntry], inst as unknown as ResizeObserver)
      }
    }
  }
}

/** The size fitZoom() sees (content box); fitZoom subtracts 56px (w) / 72px (h) padding. */
let stageSize = { w: 1336, h: 800 } // avail 1280x728 → fit zoom 1 for the blank 1280x720 deck
/** Border box of .stage-wrap: unlike the content box, scrollbars never shrink it. */
let outerSize = { w: 1336, h: 800 }

/** A real window/pane resize: content box and border box change together. */
function resizeStage(w: number, h: number) {
  stageSize = { w, h }
  outerSize = { w, h }
}

const blankSlide = () => ({
  widthPx: 1280,
  heightPx: 720,
  scale: 1,
  nodes: [],
  background: { kind: 'color', color: '#ffffff' },
})

/** Minimal slidesApi: enough for App to boot into the editor on a blank deck. */
function makeSlidesApi() {
  const explicit: Record<string, unknown> = {
    consumePendingOpen: () => Promise.resolve(null),
    newBlank: () => Promise.resolve({ path: '', slides: [blankSlide()], defaultFont: 'Arial' }),
    isDirty: () => Promise.resolve(false),
    getRecentFiles: () => Promise.resolve([]),
    getAiSettings: () => Promise.resolve(null),
    getSections: () => Promise.resolve([]),
    getComments: () => Promise.resolve([]),
    getNotes: () => Promise.resolve(''),
    getAnimations: () => Promise.resolve([]),
    getTransition: () => Promise.resolve(null),
    listFonts: () => Promise.resolve([]),
  }
  const fallbacks = new Map<string, unknown>()
  return new Proxy(explicit, {
    get(target, prop: string) {
      if (prop in target) return target[prop]
      if (!fallbacks.has(prop)) {
        fallbacks.set(
          prop,
          prop.startsWith('on')
            ? () => () => {} // event subscription → unsubscribe
            : () => Promise.resolve(null),
        )
      }
      return fallbacks.get(prop)
    },
  })
}

async function settle() {
  for (let i = 0; i < 8; i++) {
    await act(async () => {
      await new Promise((r) => setTimeout(r, 0))
    })
  }
}

/** Zoom buttons go through the shared preview path: the CSS transform lands on the next
 * rAF and the React commit after a 150ms debounce — flush both before asserting. */
async function flushZoomPreview() {
  await act(async () => {
    await new Promise((r) => requestAnimationFrame(() => r(null)))
    await new Promise((r) => setTimeout(r, 180))
  })
}

function stageZoom(container: HTMLElement): number {
  const el = container.querySelector<HTMLElement>('.stage-scale')
  expect(el, 'stage-scale should be mounted').not.toBeNull()
  const m = /scale\(([\d.]+)\)/.exec(el!.style.transform)
  expect(m, `transform should carry a scale(): ${el!.style.transform}`).not.toBeNull()
  return Number(m![1])
}

function clickByText(container: HTMLElement, selector: string, re: RegExp): void {
  const btn = [...container.querySelectorAll<HTMLButtonElement>(selector)].find((b) =>
    re.test(b.textContent ?? ''),
  )
  expect(btn, `button ${selector} matching ${re} should exist`).not.toBeNull()
  act(() => btn!.click())
}

let root: Root | null = null
let container: HTMLElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  vi.stubGlobal('ResizeObserver', FakeResizeObserver)
  window.matchMedia ??= (query: string) =>
    ({
      matches: false,
      media: query,
      addEventListener: () => {},
      removeEventListener: () => {},
      addListener: () => {},
      removeListener: () => {},
      onchange: null,
      dispatchEvent: () => false,
    }) as MediaQueryList
  Element.prototype.scrollTo ??= () => {}
  Element.prototype.scrollIntoView ??= () => {}
  // fitZoom measures the stage container; jsdom has no layout, so feed it a
  // controllable size (all elements share it — only the stage wrap is measured)
  Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
    configurable: true,
    get: () => stageSize.w,
  })
  Object.defineProperty(HTMLElement.prototype, 'clientHeight', {
    configurable: true,
    get: () => stageSize.h,
  })
  // Border-box size, mocked only for the stage wrap (other elements keep jsdom's
  // 0 so the zoom-preview transform math stays on its offsetWidth-less path)
  Object.defineProperty(HTMLElement.prototype, 'offsetWidth', {
    configurable: true,
    get(this: HTMLElement) {
      return this.classList.contains('stage-wrap') ? outerSize.w : 0
    },
  })
  Object.defineProperty(HTMLElement.prototype, 'offsetHeight', {
    configurable: true,
    get(this: HTMLElement) {
      return this.classList.contains('stage-wrap') ? outerSize.h : 0
    },
  })
  ;(window as unknown as { slidesApi: unknown }).slidesApi = makeSlidesApi()
})

afterEach(() => {
  act(() => root?.unmount())
  container?.remove()
  root = null
  container = null
  FakeResizeObserver.instances.length = 0
})

/** Mount the App and boot into a blank deck at fit zoom 1 (1336x800 container). */
async function bootApp(): Promise<HTMLElement> {
  resizeStage(1336, 800)
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => root!.render(createElement(App)))
  await settle()
  const wrap = container.querySelector<HTMLElement>('.stage-wrap')
  expect(wrap, 'editor should boot into a blank deck').not.toBeNull()
  expect(stageZoom(container!)).toBeCloseTo(1, 5)
  return wrap!
}

describe('stage fit-to-window follow', () => {
  it('hides bleed-only scrollbars until the nominal slide overflows', async () => {
    const wrap = await bootApp()
    expect(wrap.classList.contains('stage-fits-viewport')).toBe(true)

    // A zoom above the geometric fit needs normal canvas panning.
    const plus = [...container!.querySelectorAll<HTMLButtonElement>('.zoom-btn')].at(-1)!
    act(() => plus.click())
    await flushZoomPreview()
    expect(stageZoom(container!)).toBeCloseTo(1.1, 5)
    expect(wrap.classList.contains('stage-fits-viewport')).toBe(false)

    // Growing the viewport makes the unchanged manual zoom fit again. The
    // ResizeObserver must refresh the class even though it does not change zoom.
    resizeStage(1500, 900)
    await act(async () => FakeResizeObserver.fire(wrap))
    expect(stageZoom(container!)).toBeCloseTo(1.1, 5)
    expect(wrap.classList.contains('stage-fits-viewport')).toBe(true)
  })

  it('re-fits on container resize after a reading-view round trip', async () => {
    const wrap = await bootApp()

    // sanity: the follow works before any view-mode round trip
    resizeStage(696, 800) // avail 640 → fit 0.5
    await act(async () => FakeResizeObserver.fire(wrap!))
    expect(stageZoom(container!)).toBeCloseTo(0.5, 5)

    // reading view unmounts the editor; returning remounts a fresh .stage-wrap
    clickByText(container!, '.ribbon-tabs button', /^(View|视图)$/)
    clickByText(container!, 'button.rb-big', /Reading|阅读/)
    await settle()
    expect(container!.querySelector('.stage-wrap')).toBeNull()
    clickByText(container!, 'button.reading-exit', /./)
    await settle()

    const wrap2 = container!.querySelector<HTMLElement>('.stage-wrap')
    expect(wrap2, 'editor should be back').not.toBeNull()
    expect(wrap2).not.toBe(wrap)

    // shrink again: the observer must follow the NEW stage-wrap
    resizeStage(376, 800) // avail 320 → fit 0.25
    await act(async () => FakeResizeObserver.fire(wrap2!))
    expect(stageZoom(container!)).toBeCloseTo(0.25, 5)
  })

  it('clamps a manual zoom back to fit when the container shrinks', async () => {
    const wrap = await bootApp()

    // "+" bottom-bar button: manual zoom, exits fit mode
    const plus = [...container!.querySelectorAll<HTMLButtonElement>('.zoom-btn')].at(-1)!
    act(() => plus.click())
    await flushZoomPreview()
    expect(stageZoom(container!)).toBeCloseTo(1.1, 5)

    // container shrinks: 1.1 no longer fits → clamp down to the new fit
    resizeStage(696, 800) // avail 640 → fit 0.5
    await act(async () => FakeResizeObserver.fire(wrap))
    expect(stageZoom(container!)).toBeCloseTo(0.5, 5)

    // and fit mode is re-entered: the next resize follows again
    resizeStage(376, 800) // avail 320 → fit 0.25
    await act(async () => FakeResizeObserver.fire(wrap))
    expect(stageZoom(container!)).toBeCloseTo(0.25, 5)
  })

  it('leaves a manual zoom smaller than fit alone on container resize', async () => {
    const wrap = await bootApp()

    // "−" bottom-bar button: manual zoom out below fit
    const minus = container!.querySelector<HTMLButtonElement>('.zoom-btn')!
    act(() => minus.click())
    await flushZoomPreview()
    expect(stageZoom(container!)).toBeCloseTo(0.9, 5)

    // resize with fit still 1: 0.9 fits already → the user's choice is kept
    await act(async () => FakeResizeObserver.fire(wrap))
    expect(stageZoom(container!)).toBeCloseTo(0.9, 5)
  })

  it('keeps a zoom above the 1.5 fit cap while it still fits the container', async () => {
    const wrap = await bootApp()

    // container big enough for 200%: uncapped fit ≈ 2.07, fitZoom caps at 1.5
    resizeStage(2700, 1600)
    await act(async () => FakeResizeObserver.fire(wrap))
    expect(stageZoom(container!)).toBeCloseTo(1.5, 5) // fit-mode follow, capped

    // manual zoom to 2.0 — still fits geometrically
    const plus = [...container!.querySelectorAll<HTMLButtonElement>('.zoom-btn')].at(-1)!
    for (let i = 0; i < 5; i++) act(() => plus.click())
    await flushZoomPreview()
    expect(stageZoom(container!)).toBeCloseTo(2.0, 5)

    // a resize that still fits 200% must NOT wipe the manual zoom down to 1.5
    resizeStage(2680, 1590)
    await act(async () => FakeResizeObserver.fire(wrap))
    expect(stageZoom(container!)).toBeCloseTo(2.0, 5)

    // but a real overflow still clamps back to fit
    resizeStage(1336, 800) // uncapped fit 1
    await act(async () => FakeResizeObserver.fire(wrap))
    expect(stageZoom(container!)).toBeCloseTo(1, 5)
  })

  it('keeps a manual zoom when scrollbars appear without a real container resize', async () => {
    const wrap = await bootApp()

    // zoom above fit: the slide overflows and classic scrollbars appear
    const plus = [...container!.querySelectorAll<HTMLButtonElement>('.zoom-btn')].at(-1)!
    act(() => plus.click())
    await flushZoomPreview()
    expect(stageZoom(container!)).toBeCloseTo(1.1, 5)

    // The scrollbars shrink the observed content box, but the border box is
    // unchanged — no real window/pane resize happened, so the manual zoom
    // must NOT snap back to fit (regression: zooming past fit bounced back).
    stageSize = { w: 1321, h: 785 }
    await act(async () => FakeResizeObserver.fire(wrap))
    expect(stageZoom(container!)).toBeCloseTo(1.1, 5)
    expect(wrap.classList.contains('stage-fits-viewport')).toBe(false)

    // A real shrink afterwards still clamps back down to the fit value
    resizeStage(696, 800) // avail 640 → fit 0.5
    await act(async () => FakeResizeObserver.fire(wrap))
    expect(stageZoom(container!)).toBeCloseTo(0.5, 5)
  })
})
