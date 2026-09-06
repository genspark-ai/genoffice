import { afterEach, describe, expect, it } from 'vitest'
import { clearPrintZoom, printZoom, setPrintZoom } from '../src/renderer/print-zoom'

const setDpr = (dpr: number) =>
  Object.defineProperty(window, 'devicePixelRatio', { value: dpr, configurable: true })

describe('printZoom', () => {
  it('reproduces the screen device scale and is a no-op at 1', () => {
    expect(printZoom(1)).toBe(1)
    expect(printZoom(2)).toBe(2)
    expect(printZoom(1.5)).toBe(1.5)
  })

  it('clamps to the range Chromium print scales can invert', () => {
    expect(printZoom(20)).toBe(10)
    expect(printZoom(0.25)).toBe(0.5)
    expect(printZoom(NaN)).toBe(1)
  })
})

describe('setPrintZoom', () => {
  afterEach(() => {
    document.body.innerHTML = ''
    setDpr(1)
  })

  it('zooms the preview root by the device scale and returns the inverse print scale', () => {
    document.body.innerHTML = '<div class="pagination-preview"></div>'
    setDpr(2)
    expect(setPrintZoom()).toBe(0.5)
    const root = document.querySelector<HTMLElement>('.pagination-preview')!
    expect(root.style.getPropertyValue('--pv-print-zoom')).toBe('2')
    clearPrintZoom()
    expect(root.style.getPropertyValue('--pv-print-zoom')).toBe('')
  })

  it('leaves the preview untouched on standard-density screens', () => {
    document.body.innerHTML = '<div class="pagination-preview"></div>'
    setDpr(1)
    expect(setPrintZoom()).toBe(1)
    expect(
      document
        .querySelector<HTMLElement>('.pagination-preview')!
        .style.getPropertyValue('--pv-print-zoom'),
    ).toBe('')
  })
})
