import { describe, expect, it } from 'vitest'

import { resolveDefaultRowHeightPt } from '../src/renderer/univer-sync'

type Workbook = Parameters<typeof resolveDefaultRowHeightPt>[0]

function workbook(fontFamily: string, fontSize: number, normalFontName?: string): Workbook {
  return {
    styles: [{ fontFamily, fontSize }],
    ...(normalFontName === undefined ? {} : { normalFontName }),
  } as unknown as Workbook
}

const missing = { defaultRowHeight: null, defaultRowHeightFixed: false }

describe('resolveDefaultRowHeightPt', () => {
  it('derives the default from the Normal font when sheetFormatPr is missing', () => {
    expect(resolveDefaultRowHeightPt(workbook('Arial', 11), missing)).toBe(15)
    expect(resolveDefaultRowHeightPt(workbook('Calibri', 11), missing)).toBe(15)
    expect(resolveDefaultRowHeightPt(workbook('Calibri', 12), missing)).toBe(15)
    expect(resolveDefaultRowHeightPt(workbook('Calibri', 10), missing)).toBe(12.75)
  })

  it('ignores a stale default above the Normal font pitch when customHeight is off', () => {
    const sheet = { defaultRowHeight: 21.6, defaultRowHeightFixed: false }
    expect(resolveDefaultRowHeightPt(workbook('Calibri', 11), sheet)).toBe(15)
    expect(resolveDefaultRowHeightPt(workbook('Times New Roman', 14), sheet)).toBe(17.25)
  })

  it('keeps a stored default below the Normal font pitch', () => {
    const sheet = { defaultRowHeight: 10, defaultRowHeightFixed: false }
    expect(resolveDefaultRowHeightPt(workbook('Calibri', 11), sheet)).toBe(10)
  })

  it('honors the stored default when customHeight fixes it', () => {
    const sheet = { defaultRowHeight: 21.6, defaultRowHeightFixed: true }
    expect(resolveDefaultRowHeightPt(workbook('Calibri', 11), sheet)).toBe(21.6)
  })

  it('keeps a stored default within a pixel of the pitch and falls back to the pitch when missing', () => {
    const ja = workbook('Calibri', 11, 'ＭＳ Ｐゴシック')
    const sheet = { defaultRowHeight: 13.5, defaultRowHeightFixed: false }
    expect(resolveDefaultRowHeightPt(ja, sheet)).toBe(13.5)
    expect(resolveDefaultRowHeightPt(ja, missing)).toBe(12.75)
  })

  it('keeps Normal faces without any pitch entry on the stored value or 15pt', () => {
    const unknown = workbook('Calibri', 11, 'Some Unlisted Face')
    const sheet = { defaultRowHeight: 13.5, defaultRowHeightFixed: false }
    expect(resolveDefaultRowHeightPt(unknown, sheet)).toBe(13.5)
    expect(resolveDefaultRowHeightPt(unknown, missing)).toBe(15)
  })

  it('treats an older sidecar without the fixed flag as not fixed', () => {
    const sheet = { defaultRowHeight: 21.6, defaultRowHeightFixed: undefined }
    expect(resolveDefaultRowHeightPt(workbook('Calibri', 11), sheet)).toBe(15)
  })
})
