import { readFileSync } from 'node:fs'

import { describe, expect, it } from 'vitest'

import { workbookFileSchema } from '../src/shared/desktop-api'

// The open path parses the sidecar output with a strict schema, and the
// preload rebuilds sheet metadata from a hand-written whitelist — a wire
// field missing from either side silently drops the saved zoom.
const sheetSchema = workbookFileSchema.shape.sheets.element

const baseSheet = {
  id: 'sheet-1',
  name: 'Sheet1',
  rowCount: 10,
  columnCount: 5,
  columnWidths: [],
  defaultRowHeight: null,
  defaultColumnWidth: null,
  freeze: null,
  hidden: false,
  tabColor: null,
  showGridLines: true,
  tables: [],
  comments: [],
  pivotRanges: [],
}

describe('sheet zoomScale wire field', () => {
  it('accepts the sidecar zoom percent and keeps it optional', () => {
    expect(sheetSchema.parse({ ...baseSheet, zoomScale: 70 }).zoomScale).toBe(70)
    expect(sheetSchema.parse(baseSheet).zoomScale).toBeUndefined()
  })

  it('rejects out-of-range zoom percents', () => {
    expect(() => sheetSchema.parse({ ...baseSheet, zoomScale: 5 })).toThrow()
    expect(() => sheetSchema.parse({ ...baseSheet, zoomScale: 500 })).toThrow()
    expect(() => sheetSchema.parse({ ...baseSheet, zoomScale: 70.5 })).toThrow()
  })

  it('is named by the preload whitelist', () => {
    const preloadSource = readFileSync(new URL('../src/preload/index.ts', import.meta.url), 'utf8')
    expect(preloadSource).toContain('zoomScale')
  })
})
