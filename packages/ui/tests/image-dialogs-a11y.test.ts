/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  CropDialog,
  cropEdgeArrowDelta,
  cropEdgeValue,
  nudgeCropEdge,
  CROP_EDGES,
  type CropEdge,
  type CropFractions,
  type ImageDialogLabels,
} from '../src/image-dialogs'
import { CROP_EDGE_LABELS } from '../src/strings-crop-edges'

const LABELS: ImageDialogLabels = {
  cancel: 'Cancel',
  apply: 'Apply',
  applying: 'Applying',
  loading: 'Loading',
  loadFailed: 'load failed',
  processFailed: 'process failed',
  cutoutTitle: 'Remove background',
  tolerance: 'Tolerance',
  cutoutHint: (pct) => `${pct}%`,
  cropTitle: 'Crop',
  cropHint: 'hint',
}

const FULL: CropFractions = { l: 0, t: 0, r: 1, b: 1 }

let host: HTMLDivElement
let root: ReturnType<typeof createRoot>
let onApply: ReturnType<typeof vi.fn>
let onCancel: ReturnType<typeof vi.fn>

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  onApply = vi.fn()
  onCancel = vi.fn()
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

/** The dialog only renders handles once the image reports a size, which jsdom
 * never does on its own. */
function mountCrop(lang?: Parameters<typeof CropDialog>[0]['lang']) {
  class FakeImage {
    onload: (() => void) | null = null
    onerror: (() => void) | null = null
    naturalWidth = 400
    naturalHeight = 300
    set src(_value: string) {
      this.onload?.()
    }
  }
  vi.stubGlobal('Image', FakeImage)
  act(() =>
    root.render(
      createElement(CropDialog, { labels: LABELS, image: 'data:,x', onApply, onCancel, lang }),
    ),
  )
  return host
}

const edge = (scope: HTMLElement, pos: string): HTMLElement =>
  scope.querySelector(`[data-crop-edge="${pos}"]`) as HTMLElement

const key = (target: HTMLElement, k: string, init: KeyboardEventInit = {}) => {
  act(() => {
    target.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, ...init }))
  })
}

describe('nudgeCropEdge', () => {
  it('moves only the sides the handle owns', () => {
    const start: CropFractions = { l: 0.2, t: 0.3, r: 0.8, b: 0.7 }
    const expectCrop = (actual: CropFractions, expected: CropFractions) => {
      expect(actual.l).toBeCloseTo(expected.l, 10)
      expect(actual.t).toBeCloseTo(expected.t, 10)
      expect(actual.r).toBeCloseTo(expected.r, 10)
      expect(actual.b).toBeCloseTo(expected.b, 10)
    }
    expectCrop(nudgeCropEdge(start, 'e', 0.05, 0.05, 0, 0), { l: 0.2, t: 0.3, r: 0.85, b: 0.7 })
    expectCrop(nudgeCropEdge(start, 'w', -0.05, 0, 0, 0), { l: 0.15, t: 0.3, r: 0.8, b: 0.7 })
    expectCrop(nudgeCropEdge(start, 'n', 0, -0.05, 0, 0), { l: 0.2, t: 0.25, r: 0.8, b: 0.7 })
    expectCrop(nudgeCropEdge(start, 's', 0, 0.05, 0, 0), { l: 0.2, t: 0.3, r: 0.8, b: 0.75 })
  })

  it('clamps to the image and to the minimum box size', () => {
    expect(nudgeCropEdge(FULL, 'e', 0.5, 0, 0, 0).r).toBe(1)
    expect(nudgeCropEdge(FULL, 'w', -0.5, 0, 0, 0).l).toBe(0)
    const tight = nudgeCropEdge(FULL, 'w', 0.9, 0, 0.2, 0.2)
    expect(tight.l).toBe(0.8)
    const tightS = nudgeCropEdge(FULL, 'n', 0, 0.9, 0.2, 0.2)
    expect(tightS.t).toBe(0.8)
  })

  it('moves both sides of a corner', () => {
    expect(nudgeCropEdge(FULL, 'nw', -0.1, -0.1, 0, 0)).toEqual({ l: 0, t: 0, r: 1, b: 1 })
    const start: CropFractions = { l: 0.2, t: 0.2, r: 0.8, b: 0.8 }
    expect(nudgeCropEdge(start, 'se', 0.1, 0.1, 0, 0)).toEqual({ l: 0.2, t: 0.2, r: 0.9, b: 0.9 })
  })
})

describe('cropEdgeArrowDelta', () => {
  it('ignores a key that does not move the edge', () => {
    expect(cropEdgeArrowDelta('ArrowLeft', 'n', false)).toBeNull()
    expect(cropEdgeArrowDelta('ArrowUp', 'e', false)).toBeNull()
    expect(cropEdgeArrowDelta('Enter', 'e', false)).toBeNull()
  })

  it('takes the coarse step with Shift', () => {
    expect(cropEdgeArrowDelta('ArrowRight', 'e', false)).toEqual({ dx: 0.01, dy: 0 })
    expect(cropEdgeArrowDelta('ArrowRight', 'e', true)).toEqual({ dx: 0.1, dy: 0 })
    expect(cropEdgeArrowDelta('ArrowUp', 'n', true)).toEqual({ dx: 0, dy: -0.1 })
    expect(cropEdgeArrowDelta('ArrowUp', 'nw', true)).toEqual({ dx: 0, dy: -0.1 })
    expect(cropEdgeArrowDelta('ArrowDown', 'se', true)).toEqual({ dx: 0, dy: 0.1 })
  })
})

describe('cropEdgeValue', () => {
  it('reports the side the handle owns', () => {
    const crop: CropFractions = { l: 0.1, t: 0.2, r: 0.7, b: 0.8 }
    expect(cropEdgeValue(crop, 'n')).toBe(0.2)
    expect(cropEdgeValue(crop, 's')).toBe(0.8)
    expect(cropEdgeValue(crop, 'w')).toBe(0.1)
    expect(cropEdgeValue(crop, 'e')).toBe(0.7)
  })
})

describe('shared CropDialog keyboard and focus', () => {
  it('exposes every edge as a named tab stop', () => {
    mountCrop()
    const handles = host.querySelectorAll('[data-crop-edge]')
    expect(handles).toHaveLength(8)
    for (const handle of handles) {
      expect(handle.getAttribute('role')).toBe('slider')
      expect(handle.getAttribute('tabindex')).toBe('0')
      expect(handle.getAttribute('aria-label')).toBeTruthy()
      expect(handle.getAttribute('aria-valuemin')).toBe('0')
      expect(handle.getAttribute('aria-valuemax')).toBe('100')
      const pos = handle.getAttribute('data-crop-edge') as CropEdge
      expect(handle.getAttribute('aria-valuenow')).toBe(
        String(Math.round(cropEdgeValue(FULL, pos) * 100)),
      )
    }
    expect(CROP_EDGES).toHaveLength(8)
    expect(edge(host, 'w').getAttribute('aria-label')).toBe('Left edge')
    expect(edge(host, 'nw').getAttribute('aria-orientation')).toBe('vertical')
  })

  it('localizes the edge names', () => {
    mountCrop('zh')
    expect(edge(host, 'w').getAttribute('aria-label')).toBe('\u5de6\u8fb9\u7f18')
    expect(edge(host, 'se').getAttribute('aria-label')).toBe('\u53f3\u4e0b\u8fb9\u7f18')
    expect(CROP_EDGE_LABELS.en.nw).toBe('Top left edge')
  })

  it('moves focus into the dialog on open', () => {
    const outside = document.createElement('button')
    document.body.append(outside)
    outside.focus()
    mountCrop()
    expect(host.querySelector('[role="dialog"]')!.contains(document.activeElement)).toBe(true)
    outside.remove()
  })

  it('adjusts the focused edge with the arrow keys', () => {
    mountCrop()
    const w = edge(host, 'w')
    key(w, 'ArrowRight')
    expect(Number(w.getAttribute('aria-valuenow'))).toBe(1)
    key(w, 'ArrowRight', { shiftKey: true })
    expect(Number(w.getAttribute('aria-valuenow'))).toBe(11)
    key(w, 'ArrowLeft')
    expect(Number(w.getAttribute('aria-valuenow'))).toBe(10)
  })

  it('leaves an edge alone for an arrow it does not own', () => {
    mountCrop()
    const n = edge(host, 'n')
    key(n, 'ArrowLeft')
    expect(n.getAttribute('aria-valuenow')).toBe('0')
    key(n, 'ArrowDown')
    expect(n.getAttribute('aria-valuenow')).toBe('1')
  })

  it('never drives an edge past the image bounds', () => {
    mountCrop()
    const w = edge(host, 'w')
    for (let i = 0; i < 40; i++) key(w, 'ArrowLeft', { shiftKey: true })
    expect(Number(w.getAttribute('aria-valuenow'))).toBe(0)
  })

  it('does not apply the crop when Enter lands on a handle', () => {
    mountCrop()
    key(edge(host, 'w'), 'Enter')
    expect(onApply).not.toHaveBeenCalled()
  })

  it('closes on Escape without applying', () => {
    mountCrop()
    key(host.querySelector('[role="dialog"]')!, 'Escape')
    expect(onCancel).toHaveBeenCalledTimes(1)
    expect(onApply).not.toHaveBeenCalled()
  })

  it('names the dialog and marks it modal', () => {
    mountCrop()
    const dialog = host.querySelector('[role="dialog"]')!
    expect(dialog.getAttribute('aria-label')).toBe('Crop')
    expect(dialog.getAttribute('aria-modal')).toBe('true')
  })

  it('focuses Apply on open so Enter applies instead of cancelling', () => {
    mountCrop()
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
      drawImage() {},
    } as unknown as CanvasRenderingContext2D)
    vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue('data:image/png;base64,QUJD')
    const active = document.activeElement as HTMLElement
    expect(active.textContent).toBe('Apply')
    key(active, 'Enter')
    // jsdom does not run a button's Enter activation; emulate the browser default
    if (active instanceof HTMLButtonElement) act(() => active.click())
    expect(onCancel).not.toHaveBeenCalled()
    expect(onApply).toHaveBeenCalledTimes(1)
    expect(onApply.mock.calls[0]![0]).toBe('QUJD')
  })

  it('wraps Tab from the last control to the first', () => {
    mountCrop()
    const dialog = host.querySelector('[role="dialog"]') as HTMLElement
    const focusables = [...dialog.querySelectorAll<HTMLElement>('button, [tabindex="0"]')]
    const last = focusables[focusables.length - 1]!
    last.focus()
    const e = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true })
    act(() => {
      dialog.dispatchEvent(e)
    })
    expect(e.defaultPrevented).toBe(true)
    expect(document.activeElement).toBe(focusables[0])
    expect(focusables[0]!.getAttribute('data-crop-edge')).toBe('nw')
  })

  it('wraps Shift+Tab from the first control to the last', () => {
    mountCrop()
    const dialog = host.querySelector('[role="dialog"]') as HTMLElement
    const focusables = [...dialog.querySelectorAll<HTMLElement>('button, [tabindex="0"]')]
    focusables[0]!.focus()
    const e = new KeyboardEvent('keydown', {
      key: 'Tab',
      shiftKey: true,
      bubbles: true,
      cancelable: true,
    })
    act(() => {
      dialog.dispatchEvent(e)
    })
    expect(e.defaultPrevented).toBe(true)
    expect(document.activeElement).toBe(focusables[focusables.length - 1])
    expect(document.activeElement!.textContent).toBe('Apply')
  })

  it('restores focus to the invoking control on close', () => {
    const trigger = document.createElement('button')
    document.body.append(trigger)
    trigger.focus()
    mountCrop()
    act(() => root.unmount())
    expect(document.activeElement).toBe(trigger)
    trigger.remove()
  })
})

describe('CROP_EDGE_LABELS', () => {
  const edges: CropEdge[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']

  it('names all eight edges in every language', () => {
    for (const [lang, names] of Object.entries(CROP_EDGE_LABELS)) {
      expect(Object.keys(names).sort(), lang).toEqual([...edges].sort())
      for (const edgeName of edges) expect(names[edgeName], `${lang}.${edgeName}`).toBeTruthy()
    }
  })

  it('has a distinct name per edge in every language', () => {
    for (const [lang, names] of Object.entries(CROP_EDGE_LABELS)) {
      expect(new Set(edges.map((e) => names[e])).size, lang).toBe(edges.length)
    }
  })
})
