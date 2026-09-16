import { describe, expect, it, vi } from 'vitest'
import { handlePdfControl, type PdfControlState } from '../src/renderer/control'

const state = (overrides: Partial<PdfControlState> = {}): PdfControlState => ({
  loaded: true,
  pageCount: 4,
  currentPage: 2,
  scrollToPage: vi.fn(),
  selectedText: () => '',
  ...overrides,
})

describe('pdf control hook', () => {
  it('scrolls to a page inside the document and rejects one outside it', () => {
    const s = state()
    expect(handlePdfControl({ cmd: 'goto', target: { kind: 'page', page: 3 } }, s)).toEqual({
      status: 'ok',
      result: { page: 3 },
    })
    expect(s.scrollToPage).toHaveBeenCalledWith(3)
    expect(handlePdfControl({ cmd: 'goto', target: { kind: 'page', page: 0 } }, s)).toMatchObject({
      status: 'error',
      error: { reason: 'out_of_range', detail: { valid_range: '1-4' } },
    })
  })

  it('reports the current page and any selected text, or not_ready before the file loads', () => {
    expect(handlePdfControl({ cmd: 'selection' }, state({ loaded: false }))).toEqual({
      status: 'not_ready',
    })
    expect(handlePdfControl({ cmd: 'selection' }, state({ selectedText: () => 'quoted' }))).toEqual(
      {
        status: 'ok',
        result: { page: 2, text: 'quoted' },
      },
    )
  })
})
