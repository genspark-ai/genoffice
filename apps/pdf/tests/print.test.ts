import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { printPdf, printScaleForAreas } from '../src/renderer/print'

function fakeDoc(numPages: number, render = vi.fn(() => ({ promise: Promise.resolve() }))) {
  const getPage = vi.fn(async () => ({
    getViewport: ({ scale }: { scale: number }) => ({ width: 612 * scale, height: 792 * scale }),
    render,
  }))
  return { doc: { numPages, getPage } as unknown as PDFDocumentProxy, getPage, render }
}

beforeEach(() => {
  vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue('data:image/jpeg;base64,FAKE')
  // jsdom does not implement img.decode or window.print
  Object.defineProperty(HTMLImageElement.prototype, 'decode', {
    configurable: true,
    writable: true,
    value: vi.fn(() => Promise.resolve()),
  })
  window.print = vi.fn(() => {
    window.dispatchEvent(new Event('afterprint'))
  })
})

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

describe('printPdf', () => {
  it('renders one image per page into a print root and opens the print dialog', async () => {
    const { doc, getPage, render } = fakeDoc(3)
    let imgsAtPrintTime = 0
    ;(window.print as ReturnType<typeof vi.fn>).mockImplementation(() => {
      imgsAtPrintTime = document.querySelectorAll('.pdf-print-root img').length
      window.dispatchEvent(new Event('afterprint'))
    })

    await printPdf(doc)

    expect(getPage).toHaveBeenCalledTimes(3)
    expect(getPage).toHaveBeenNthCalledWith(1, 1)
    expect(render).toHaveBeenCalledTimes(3)
    expect(imgsAtPrintTime).toBe(3)
    expect(window.print).toHaveBeenCalledTimes(1)
  })

  it('removes the print root after the dialog closes', async () => {
    const { doc } = fakeDoc(2)
    await printPdf(doc)
    expect(document.querySelector('.pdf-print-root')).toBeNull()
  })

  it('renders only the requested pages, sorted, when given a subset', async () => {
    const { doc, getPage, render } = fakeDoc(5)
    let imgsAtPrintTime = 0
    ;(window.print as ReturnType<typeof vi.fn>).mockImplementation(() => {
      imgsAtPrintTime = document.querySelectorAll('.pdf-print-root img').length
      window.dispatchEvent(new Event('afterprint'))
    })

    await printPdf(doc, [4, 2, 99, 0])

    // Out-of-range entries are dropped; the rest render in document order.
    expect(getPage).toHaveBeenCalledTimes(2)
    expect(getPage).toHaveBeenNthCalledWith(1, 2)
    expect(getPage).toHaveBeenNthCalledWith(2, 4)
    expect(render).toHaveBeenCalledTimes(2)
    expect(imgsAtPrintTime).toBe(2)
    expect(window.print).toHaveBeenCalledTimes(1)
  })

  it('waits for afterprint before resolving', async () => {
    const { doc } = fakeDoc(1)
    let fireAfterPrint: () => void = () => {}
    ;(window.print as ReturnType<typeof vi.fn>).mockImplementation(() => {
      fireAfterPrint = () => window.dispatchEvent(new Event('afterprint'))
    })

    let resolved = false
    const done = printPdf(doc).then(() => {
      resolved = true
    })
    // Let rendering and the print call complete
    await vi.waitFor(() => expect(window.print).toHaveBeenCalled())
    expect(resolved).toBe(false)
    expect(document.querySelector('.pdf-print-root')).not.toBeNull()

    fireAfterPrint()
    await done
    expect(resolved).toBe(true)
    expect(document.querySelector('.pdf-print-root')).toBeNull()
  })

  it('cleans up the print root even when image decoding fails', async () => {
    const { doc } = fakeDoc(1)
    ;(HTMLImageElement.prototype.decode as ReturnType<typeof vi.fn>).mockRejectedValue(
      new Error('decode failed'),
    )
    await expect(printPdf(doc)).rejects.toThrow('decode failed')
    expect(document.querySelector('.pdf-print-root')).toBeNull()
    expect(window.print).not.toHaveBeenCalled()
  })
})

/** US-Letter page area at scale 1 (PDF points²) */
const LETTER_AREA = 612 * 792

describe('printScaleForAreas', () => {
  it('uses the 200 DPI target while the document fits the pixel budget', () => {
    expect(printScaleForAreas([LETTER_AREA])).toBeCloseTo(200 / 72, 10)
    // ~40 letter pages at 200 DPI ≈ the 150 MP budget
    expect(printScaleForAreas(new Array(40).fill(LETTER_AREA))).toBeCloseTo(200 / 72, 10)
  })

  it('falls back to the target for empty or degenerate input', () => {
    expect(printScaleForAreas([])).toBeCloseTo(200 / 72, 10)
    expect(printScaleForAreas([0])).toBeCloseTo(200 / 72, 10)
  })

  it('scales down proportionally over budget, never below the 150 DPI baseline', () => {
    const mid = printScaleForAreas(new Array(60).fill(LETTER_AREA))
    expect(mid).toBeLessThan(200 / 72)
    expect(mid).toBeGreaterThan(150 / 72)
    expect(printScaleForAreas(new Array(500).fill(LETTER_AREA))).toBeCloseTo(150 / 72, 10)
  })
})

describe('printPdf render scale', () => {
  function scalesDoc(numPages: number) {
    const scales: number[] = []
    const getViewport = vi.fn(({ scale }: { scale: number }) => {
      scales.push(scale)
      return { width: 612 * scale, height: 792 * scale }
    })
    const render = vi.fn(() => ({ promise: Promise.resolve() }))
    const getPage = vi.fn(async () => ({ getViewport, render }))
    const doc = { numPages, getPage } as unknown as PDFDocumentProxy
    return { doc, scales, getPage, render }
  }

  it('renders small documents at 200 DPI (one measure call at scale 1 per page)', async () => {
    const { doc, scales, getPage } = scalesDoc(2)
    await printPdf(doc)
    expect(getPage).toHaveBeenCalledTimes(2)
    expect(scales).toEqual([1, 200 / 72, 1, 200 / 72])
  })

  it('floors huge documents at the 150 DPI baseline', async () => {
    const { doc, scales } = scalesDoc(200)
    await printPdf(doc)
    expect(scales.filter((s) => s !== 1)).toHaveLength(200)
    for (const s of scales) {
      expect(s === 1 || s === 150 / 72).toBe(true)
    }
  })
})
