import { describe, expect, it } from 'vitest'
import { createPdfSkill } from '../src/renderer/ai/pdf-skill'
import { PDF_NAV_SCHEME, parsePdfNavHref } from '../src/renderer/ai/pdf-nav'
import type { PdfAiDeps } from '../src/renderer/ai/tools'

function makeDeps(over: Partial<PdfAiDeps> = {}): PdfAiDeps {
  const stub = new Proxy(
    {
      fileName: () => 'doc.pdf',
      pageCount: () => 3,
      currentPage: () => 2,
      readOnly: () => false,
      ocrText: () => null,
      selection: () => null,
      pendingSummary: () => '',
      annotationSummary: () => '',
      outline: () => null,
      ...over,
    } as Record<string, unknown>,
    {
      get(target, prop: string) {
        if (prop in target) return target[prop]
        return () => {
          throw new Error(`unexpected deps call: ${prop}`)
        }
      },
    },
  )
  return stub as unknown as PdfAiDeps
}

const contextOf = (over: Partial<PdfAiDeps> = {}): string =>
  createPdfSkill(makeDeps(over)).buildContext?.() ?? ''

describe('buildContext', () => {
  it('has only the document line by default', () => {
    const ctx = contextOf()
    expect(ctx).toContain('"doc.pdf", 3 pages')
    expect(ctx).not.toContain('selected')
    expect(ctx).not.toContain('Unsaved')
  })

  it('injects the cached selection with its page number', () => {
    const ctx = contextOf({ selection: () => ({ page: 2, lastPage: 2, text: 'lorem ipsum' }) })
    expect(ctx).toContain('selected the following text on page 2')
    expect(ctx).toContain('lorem ipsum')
  })

  it('labels a cross-page selection with its page span', () => {
    const ctx = contextOf({ selection: () => ({ page: 2, lastPage: 3, text: 'spans pages' }) })
    expect(ctx).toContain('on pages 2-3')
  })

  it('truncates an over-long selection', () => {
    const ctx = contextOf({ selection: () => ({ page: 1, lastPage: 1, text: 'x'.repeat(20_000) }) })
    expect(ctx.length).toBeLessThan(14_000)
    expect(ctx).toContain('…')
  })

  it('skips whitespace-only selections', () => {
    const ctx = contextOf({ selection: () => ({ page: 1, lastPage: 1, text: '  \n ' }) })
    expect(ctx).not.toContain('selected')
  })

  it('injects the pending-edits summary when present', () => {
    const line = 'Unsaved changes queued this session: text edits: 2.'
    expect(contextOf({ pendingSummary: () => line })).toContain(line)
  })

  it('injects the annotation summary when present', () => {
    const line = 'The document has 3 note thread(s); use read_annotations to read them.'
    expect(contextOf({ annotationSummary: () => line })).toContain(line)
  })
})

describe('parsePdfNavHref', () => {
  it('parses page links', () => {
    expect(parsePdfNavHref('pdfnav://page/12')).toBe(12)
    expect(parsePdfNavHref(`${PDF_NAV_SCHEME}page/1`)).toBe(1)
  })

  it('rejects everything else', () => {
    expect(parsePdfNavHref('pdfnav://page/')).toBeNull()
    expect(parsePdfNavHref('pdfnav://page/-1')).toBeNull()
    expect(parsePdfNavHref('pdfnav://block/3')).toBeNull()
    expect(parsePdfNavHref('https://example.com')).toBeNull()
    expect(parsePdfNavHref('mdnav://block/3')).toBeNull()
  })
})
