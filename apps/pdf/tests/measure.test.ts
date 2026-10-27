import { describe, expect, it } from 'vitest'
import { formatLength, parseScale } from '../src/renderer/measure'
import { applySaveRequest } from '../src/main/save-pdf'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { PDFDocument } from 'pdf-lib'
import type { SavePdfRequest } from '../src/shared/ipc'

describe('formatLength', () => {
  const scale = { unitsPerPt: 100, unit: 'm' }
  it('converts points through the scale with two decimals', () => {
    expect(formatLength(235, scale)).toBe('2.35 m')
  })
  it('rounds to two decimals', () => {
    expect(formatLength(100, scale)).toBe('1.00 m')
    expect(formatLength(333.333, scale)).toBe('3.33 m')
  })
})

describe('parseScale', () => {
  it('accepts a positive value with a short unit', () => {
    expect(parseScale('100', 'm')).toEqual({ unitsPerPt: 100, unit: 'm' })
  })
  it('rejects zero, negatives, garbage and overlong units', () => {
    expect(parseScale('0', 'm')).toBeNull()
    expect(parseScale('-5', 'm')).toBeNull()
    expect(parseScale('abc', 'm')).toBeNull()
    expect(parseScale('10', '')).toBeNull()
    expect(parseScale('10', 'x'.repeat(13))).toBeNull()
  })
})

describe('the measure annotation in a save', () => {
  it('writes a Line annotation carrying the label as /Contents', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    const request: SavePdfRequest = {
      path: 'fixture.pdf',
      markups: [],
      drawings: [
        {
          kind: 'measure',
          pageIndex: 0,
          color: [0.2, 0.4, 0.9],
          width: 1.5,
          from: [50, 60],
          to: [250, 60],
          label: '2.00 m',
        },
      ],
      formValues: [],
      stamps: [],
    }
    const { bytes } = await applySaveRequest(await doc.save(), request)
    const pdoc = await getDocument({ data: bytes.slice() }).promise
    const page = await pdoc.getPage(1)
    const annots = await page.getAnnotations()
    const line = annots.find((a) => a.subtype === 'Line')
    expect(line).toBeTruthy()
    expect(line!.lineCoordinates).toEqual([50, 60, 250, 60])
    // pdf.js v6 exposes /Contents as contentsObj; plain  is only set on some paths
    expect(line!.contentsObj?.str).toBe('2.00 m')
  })
})
