// #763: opening a 409-slide deck mounted a Konva Stage — one canvas — per slide on
// first paint (411 canvases for 409 slides). The rail must mount only the rows
// around the viewport, keep the row geometry intact for everything that measures
// it, and mount the far end of the deck once it is scrolled to.
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
import { THUMB_GAP_PX, WINDOWING_MIN_SLIDES, thumbRowHeight } from '../src/renderer/thumb-window'

const SLIDE_COUNT = 120
const RAIL_VIEWPORT_H = 800

class FakeResizeObserver {
  static instances: FakeResizeObserver[] = []
  targets = new Set<Element>()
  constructor(private cb: ResizeObserverCallback) {
    FakeResizeObserver.instances.push(this)
  }
  observe(el: Element) {
    this.targets.add(el)
    this.cb([{ target: el } as ResizeObserverEntry], this as unknown as ResizeObserver)
  }
  unobserve(el: Element) {
    this.targets.delete(el)
  }
  disconnect() {
    this.targets.clear()
  }
}

const slide = () => ({
  widthPx: 1280,
  heightPx: 720,
  scale: 1,
  nodes: [],
  background: { kind: 'color', color: '#ffffff' },
})

/** Minimal slidesApi: boots the editor on a long deck. */
function makeSlidesApi() {
  const deck = {
    path: '',
    slides: Array.from({ length: SLIDE_COUNT }, slide),
    defaultFont: 'Arial',
  }
  const explicit: Record<string, unknown> = {
    consumePendingOpen: () => Promise.resolve(null),
    newBlank: () => Promise.resolve(deck),
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
        fallbacks.set(prop, prop.startsWith('on') ? () => () => {} : () => Promise.resolve(null))
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

/** The rail updates its window on the next animation frame after a scroll. */
async function flushFrame() {
  await act(async () => {
    await new Promise((r) => requestAnimationFrame(() => r(null)))
  })
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
  // jsdom has no layout: give every element the rail's viewport height
  Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
    configurable: true,
    get: () => 1336,
  })
  Object.defineProperty(HTMLElement.prototype, 'clientHeight', {
    configurable: true,
    get: () => RAIL_VIEWPORT_H,
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

async function bootApp(): Promise<HTMLElement> {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => root!.render(createElement(App)))
  await settle()
  const list = container.querySelector<HTMLElement>('.slide-list')
  expect(list, 'the thumbnail rail should mount for a long deck').not.toBeNull()
  return list!
}

/** The rail's own thumb width: sidebar width minus the horizontal padding and border. */
function railThumbW(list: HTMLElement): number {
  return Math.max(60, Number.parseFloat(list.style.width) - 24)
}
const railRowH = (list: HTMLElement) => thumbRowHeight(slide(), railThumbW(list))

const mountedThumbs = (el: HTMLElement) =>
  el.querySelectorAll('.thumb:not(.thumb-placeholder)').length
const placeholderThumbs = (el: HTMLElement) => el.querySelectorAll('.thumb-placeholder').length

describe('long-deck thumbnail rail (#763)', () => {
  it('mounts a window of rows instead of one per slide', async () => {
    const list = await bootApp()

    const mounted = mountedThumbs(list)
    expect(mounted).toBeGreaterThan(0)
    // The window is the viewport plus overscan: a deck of 409 must not mount 409.
    expect(mounted).toBeLessThan(WINDOWING_MIN_SLIDES)
    expect(mounted * railRowH(list)).toBeLessThan(RAIL_VIEWPORT_H * 3)
  })

  it('keeps every slide in the rail as a placeholder so the geometry is unchanged', async () => {
    const list = await bootApp()

    // Real rows plus placeholders still describe all 409 slides: the rail's scroll
    // height, drag drop targets and the blank-space insertion point all read the DOM.
    expect(mountedThumbs(list) + placeholderThumbs(list)).toBe(SLIDE_COUNT)
    expect(list.querySelectorAll('[data-index]').length).toBe(SLIDE_COUNT)
  })

  it('gives every spacer the height of a whole row', async () => {
    // This is the invariant that was missing: the window is picked from a row model
    // built with `thumbRowHeight` (stage + 4px borders + 10px gap), while the spacer
    // only set the *stage* height. Under `* { box-sizing: border-box }` that put the
    // borders inside the box, so each spacer was 4px short — 4px of drift per row
    // against the model. By row ~300 that is 1,200px (~17 rows against the 600px
    // overscan), which is why a long jump left its target row as a placeholder and
    // why fast scrolling showed blanks. jsdom has no layout to catch that; this
    // asserts the arithmetic the layout would have shown.
    const list = await bootApp()
    const thumbW = railThumbW(list)
    const rowHeight = thumbRowHeight(slide(), thumbW)

    const spacers = [...list.querySelectorAll<HTMLElement>('.thumb-placeholder')]
    expect(spacers.length).toBeGreaterThan(0)
    for (const spacer of spacers) {
      expect(Number.parseFloat(spacer.style.height)).toBe(rowHeight - THUMB_GAP_PX)
    }
    // and they add up to exactly the rows they stand in for
    const asRows = spacers.reduce(
      (sum, el) => sum + Number.parseFloat(el.style.height) + THUMB_GAP_PX,
      0,
    )
    expect(asRows).toBe(spacers.length * rowHeight)
  })

  it('mounts the far end of the deck once it is scrolled to', async () => {
    const list = await bootApp()
    const rowH = railRowH(list)
    const total = rowH * SLIDE_COUNT

    expect(list.querySelector(`[data-index="${SLIDE_COUNT - 1}"]`)?.className).toContain(
      'thumb-placeholder',
    )

    act(() => {
      list.scrollTop = total - RAIL_VIEWPORT_H
      list.dispatchEvent(new Event('scroll'))
    })
    await flushFrame()

    const last = list.querySelector<HTMLElement>(`[data-index="${SLIDE_COUNT - 1}"]`)
    expect(last).not.toBeNull()
    expect(last!.className).not.toContain('thumb-placeholder')
    // and the top of the deck is now the placeholder side
    expect(list.querySelector('[data-index="0"]')?.className).toContain('thumb-placeholder')
  })

  it('mounts the selected slide when it is scrolled into view from elsewhere', async () => {
    const list = await bootApp()
    const rowH = railRowH(list)

    act(() => {
      list.scrollTop = 100 * rowH
      list.dispatchEvent(new Event('scroll'))
    })
    await flushFrame()

    const mountedIndexes = [...list.querySelectorAll<HTMLElement>('.thumb:not(.thumb-placeholder)')]
      .map((el) => Number(el.dataset.index))
      .sort((a, b) => a - b)
    expect(mountedIndexes.length).toBeGreaterThan(0)
    // the mounted rows really are the ones around 100
    expect(mountedIndexes[0]).toBeGreaterThan(80)
    expect(mountedIndexes.at(-1)!).toBeLessThan(120)
  })
})
