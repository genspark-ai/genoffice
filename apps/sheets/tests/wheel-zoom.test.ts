// Ctrl+wheel in Sheets: a mouse notch steps ten percentage points through the
// worksheet zoom API and is kept from Univer; a trackpad pinch is left to
// Univer's continuous handler; nothing zooms while a cell is being edited.
import { describe, expect, it } from 'vitest'
import { SHEET_ZOOM_MAX, SHEET_ZOOM_MIN, createWheelZoomStepper } from '../src/renderer/wheel-zoom'

function harness(zoom: number, editing = false) {
  const set: number[] = []
  let prevented = 0
  let stopped = 0
  const step = createWheelZoomStepper({
    getZoom: () => zoom,
    setZoom: (z) => {
      set.push(z)
      zoom = z
    },
    isCellEditing: () => editing,
  })
  const wheel = (deltaY: number, timeStamp: number, extra: Partial<WheelEvent> = {}) =>
    step({
      deltaY,
      deltaMode: 0,
      ctrlKey: true,
      metaKey: false,
      timeStamp,
      preventDefault: () => void prevented++,
      stopPropagation: () => void stopped++,
      ...extra,
    })
  return { set, wheel, prevented: () => prevented, stopped: () => stopped }
}

describe('createWheelZoomStepper', () => {
  it('50% plus one Windows notch is 60%, not the cap', () => {
    const h = harness(0.5)
    h.wheel(-100, 0)
    expect(h.set).toEqual([0.6])
    expect(h.stopped()).toBe(1)
    h.wheel(-100, 150)
    h.wheel(100, 300)
    expect(h.set).toEqual([0.6, 0.7, 0.6])
  })

  it('clamps to the slider range and skips a no-op set', () => {
    const lo = harness(SHEET_ZOOM_MIN)
    lo.wheel(100, 0)
    expect(lo.set).toEqual([])
    const hi = harness(3.95)
    hi.wheel(-100, 0)
    expect(hi.set).toEqual([SHEET_ZOOM_MAX])
  })

  it('lets a trackpad pinch through to Univer', () => {
    const h = harness(1)
    h.wheel(-3.4, 0)
    h.wheel(-7.25, 16)
    expect(h.set).toEqual([])
    expect(h.prevented()).toBe(0)
    expect(h.stopped()).toBe(0)
  })

  it('swallows the accumulating part of a Cmd+two-finger stream and steps once', () => {
    const h = harness(1)
    const cmd = { ctrlKey: false, metaKey: true }
    h.wheel(-10, 0, cmd)
    h.wheel(-20, 16, cmd)
    h.wheel(-40, 32, cmd)
    h.wheel(-120, 48, cmd)
    expect(h.set).toEqual([1.1])
    expect(h.stopped()).toBe(4)
  })

  it('ignores plain scrolling and wheel events during cell editing', () => {
    const plain = harness(1)
    plain.wheel(-100, 0, { ctrlKey: false })
    expect(plain.set).toEqual([])
    expect(plain.prevented()).toBe(0)
    const editing = harness(1, true)
    editing.wheel(-100, 0)
    expect(editing.set).toEqual([])
    expect(editing.prevented()).toBe(0)
  })
})
