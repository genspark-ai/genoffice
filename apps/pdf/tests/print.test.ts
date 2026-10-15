import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import {
  filterPageParity,
  printPdf,
  printScaleForAreas,
  PRINT_DIALOG_TIMEOUT_MS,
} from '../src/renderer/print'

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

    expect(getPage).toHaveBeenCalledTimes(6)
    expect(getPage).toHaveBeenNthCalledWith(1, 1)
    expect(getPage).toHaveBeenNthCalledWith(2, 2)
    expect(getPage).toHaveBeenNthCalledWith(3, 3)
    expect(getPage).toHaveBeenNthCalledWith(4, 1)
    expect(getPage).toHaveBeenNthCalledWith(5, 2)
    expect(getPage).toHaveBeenNthCalledWith(6, 3)
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
    // Measure pass fetches each target first, then the render pass refetches.
    expect(getPage).toHaveBeenCalledTimes(4)
    expect(getPage).toHaveBeenNthCalledWith(1, 2)
    expect(getPage).toHaveBeenNthCalledWith(2, 4)
    expect(getPage).toHaveBeenNthCalledWith(3, 2)
    expect(getPage).toHaveBeenNthCalledWith(4, 4)
    expect(render).toHaveBeenCalledTimes(2)
    expect(imgsAtPrintTime).toBe(2)
    expect(window.print).toHaveBeenCalledTimes(1)
  })

  it('drops non-integer page targets instead of aborting the job', async () => {
    const { doc, getPage } = fakeDoc(5)
    await printPdf(doc, [1.5, 2, 99])
    // 1.5 would throw inside pdf.js getPage and abort the whole print;
    // only the valid integer target renders (measure + render passes).
    expect(getPage).toHaveBeenCalledTimes(2)
    expect(getPage).toHaveBeenNthCalledWith(1, 2)
    expect(getPage).toHaveBeenNthCalledWith(2, 2)
    expect(window.print).toHaveBeenCalledTimes(1)
  })

  it('refuses to open the print dialog when every requested page is out of range', async () => {
    const { doc, getPage } = fakeDoc(5)

    // Printing an empty document silently is worse than telling the caller
    // their selection no longer exists.
    await expect(printPdf(doc, [0])).rejects.toThrow('none of the requested pages')
    await expect(printPdf(doc, [999])).rejects.toThrow('none of the requested pages')
    expect(getPage).not.toHaveBeenCalled()
    expect(window.print).not.toHaveBeenCalled()
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

  it('gives up and releases the pages when afterprint never fires', async () => {
    vi.useFakeTimers()
    try {
      const { doc } = fakeDoc(2)
      // A suppressed print: the dialog never opens, so no event ever arrives.
      ;(window.print as ReturnType<typeof vi.fn>).mockImplementation(() => {})
      const done = printPdf(doc)
      const settled = done.then(
        () => 'resolved',
        (err: Error) => err.message,
      )
      await vi.advanceTimersByTimeAsync(PRINT_DIALOG_TIMEOUT_MS)

      // Settles instead of hanging, and the rasterized pages are released.
      await expect(settled).resolves.toContain('never reported completion')
      expect(document.querySelector('.pdf-print-root')).toBeNull()
      // Nothing left listening, so a stray event cannot resolve a dead print.
      expect(window.print).toHaveBeenCalledTimes(1)
    } finally {
      vi.useRealTimers()
    }
  })

  it('rejects when window.print throws synchronously', async () => {
    const { doc } = fakeDoc(1)
    ;(window.print as ReturnType<typeof vi.fn>).mockImplementation(() => {
      throw new Error('print blocked')
    })
    await expect(printPdf(doc)).rejects.toThrow('print blocked')
    expect(document.querySelector('.pdf-print-root')).toBeNull()
  })

  it('still resolves on afterprint without waiting out the timeout', async () => {
    vi.useFakeTimers()
    try {
      const { doc } = fakeDoc(1)
      ;(window.print as ReturnType<typeof vi.fn>).mockImplementation(() => {
        window.dispatchEvent(new Event('afterprint'))
      })
      await expect(printPdf(doc)).resolves.toBeUndefined()
      expect(document.querySelector('.pdf-print-root')).toBeNull()
    } finally {
      vi.useRealTimers()
    }
  })
})

/** US-Letter page area at scale 1 (PDF points²) */
const LETTER_AREA = 612 * 792

describe('printPdf with a printer job', () => {
  it('prints through the main process instead of the system dialog', async () => {
    const { doc } = fakeDoc(2)
    let imgsAtPrintTime = 0
    const print = vi.fn(async () => {
      imgsAtPrintTime = document.querySelectorAll('.pdf-print-root img').length
      return { ok: true }
    })
    ;(window as unknown as { pdfApi: unknown }).pdfApi = { print }

    const job = { deviceName: 'HP', copies: 2, collate: true, duplex: 'longEdge' as const }
    await printPdf(doc, undefined, job)

    expect(print).toHaveBeenCalledWith(job)
    expect(imgsAtPrintTime).toBe(2)
    expect(window.print).not.toHaveBeenCalled()
    expect(document.querySelector('.pdf-print-root')).toBeNull()
  })

  it('surfaces a main-process failure and still releases the print root', async () => {
    const { doc } = fakeDoc(1)
    ;(window as unknown as { pdfApi: unknown }).pdfApi = {
      print: vi.fn(async () => ({ ok: false, error: 'printer offline' })),
    }
    await expect(printPdf(doc, undefined, { deviceName: '' })).rejects.toThrow('printer offline')
    expect(document.querySelector('.pdf-print-root')).toBeNull()
  })
})

describe('filterPageParity', () => {
  it('keeps odd or even file pages and passes everything through for all', () => {
    const pages = [1, 2, 3, 4, 5]
    expect(filterPageParity(pages, 'all')).toBe(pages)
    expect(filterPageParity(pages, 'odd')).toEqual([1, 3, 5])
    expect(filterPageParity(pages, 'even')).toEqual([2, 4])
    expect(filterPageParity([3], 'even')).toEqual([])
  })
})

describe('printScaleForAreas', () => {
  it('uses the 200 DPI target while the document fits the pixel budget', () => {
    expect(printScaleForAreas([LETTER_AREA])).toBeCloseTo(200 / 72, 10)
    // ~40 letter pages at 200 DPI ≈ the 150 MP budget
    expect(printScaleForAreas(new Array(40).fill(LETTER_AREA))).toBeCloseTo(200 / 72, 10)
  })

  it('falls back to the target for empty or degenerate input', () => {
    expect(printScaleForAreas([])).toBeCloseTo(200 / 72, 10)
    expect(printScaleForAreas([0])).toBeCloseTo(200 / 72, 10)
    expect(printScaleForAreas([Number.NaN])).toBeCloseTo(200 / 72, 10)
    expect(printScaleForAreas([Number.POSITIVE_INFINITY])).toBeCloseTo(200 / 72, 10)
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

  it('renders small documents at 200 DPI (all pages measured before the render pass)', async () => {
    const { doc, scales, getPage } = scalesDoc(2)
    await printPdf(doc)
    expect(getPage).toHaveBeenCalledTimes(4)
    expect(scales).toEqual([1, 1, 200 / 72, 200 / 72])
  })

  it('floors huge documents at the 150 DPI baseline', async () => {
    const { doc, scales } = scalesDoc(200)
    await printPdf(doc)
    expect(scales.filter((s) => s !== 1)).toHaveLength(200)
    for (const s of scales) {
      expect(s === 1 || s === 150 / 72).toBe(true)
    }
  })

  it('budgets a page subset on its own: two pages out of a huge document print at 200 DPI', async () => {
    const { doc, scales, getPage } = scalesDoc(200)
    await printPdf(doc, [7, 3])
    expect(getPage).toHaveBeenCalledTimes(4)
    expect(getPage).toHaveBeenNthCalledWith(1, 3)
    expect(getPage).toHaveBeenNthCalledWith(2, 7)
    expect(getPage).toHaveBeenNthCalledWith(3, 3)
    expect(getPage).toHaveBeenNthCalledWith(4, 7)
    expect(scales).toEqual([1, 1, 200 / 72, 200 / 72])
  })

  it('streams pages and cleans up each one, bounding a corrupt huge page', async () => {
    const cleanups: number[] = []
    const viewports: Array<{ width: number; height: number }> = [
      { width: 612, height: 792 },
      { width: Number.NaN, height: Number.NaN },
      { width: 1e6, height: 1e6 },
    ]
    const getPage = vi.fn(async (n: number) => {
      const vp = viewports[n - 1]!
      return {
        getViewport: vi.fn(({ scale }: { scale: number }) => ({
          width: vp.width * scale,
          height: vp.height * scale,
        })),
        render: vi.fn(() => ({ promise: Promise.resolve() })),
        cleanup: vi.fn(() => cleanups.push(1)),
      }
    })
    const doc = { numPages: 3, getPage } as unknown as PDFDocumentProxy
    // The print root is removed on afterprint, so snapshot the image count
    // at print time like the other tests do.
    let imgsAtPrintTime = 0
    ;(window.print as ReturnType<typeof vi.fn>).mockImplementation(() => {
      imgsAtPrintTime = document.querySelectorAll('.pdf-print-root img').length
      window.dispatchEvent(new Event('afterprint'))
    })
    await printPdf(doc)
    // each page cleaned up twice: once after measure, once after render/skip
    expect(cleanups).toHaveLength(6)
    // NaN pages are skipped and huge pages render at their own reduced scale
    // rather than aborting the whole print
    expect(imgsAtPrintTime).toBe(2)
  })
})
