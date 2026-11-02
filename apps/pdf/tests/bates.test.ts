import { describe, expect, it } from 'vitest'
import {
  buildBatesInserts,
  formatBates,
  validateBates,
  type BatesConfig,
} from '../src/renderer/bates'

const A4 = { width: 595, height: 842, rotate: 0 }
const base: BatesConfig = {
  prefix: 'ACME-',
  start: 41,
  digits: 6,
  fontSize: 10,
  color: [0.1, 0.1, 0.12],
  position: 'bottom-right',
  firstPage: 1,
  lastPage: null,
}
/** widths keyed to the text, so alignment offsets are assertable */
const measure = (text: string) => text.length * 5

describe('formatBates', () => {
  it('pads to the digit count after the prefix', () => {
    expect(formatBates('ACME-', 41, 6)).toBe('ACME-000041')
  })
  it('never truncates a number that outgrows its field', () => {
    expect(formatBates('', 1234567, 3)).toBe('1234567')
  })
})

describe('validateBates', () => {
  it('accepts the defaults', () => {
    expect(validateBates(base)).toBeNull()
  })
  it('rejects a fractional or negative start, bad digits, bad range', () => {
    expect(validateBates({ ...base, start: -1 })).toBe('batesErrStart')
    expect(validateBates({ ...base, start: 1.5 })).toBe('batesErrStart')
    expect(validateBates({ ...base, digits: 0 })).toBe('batesErrDigits')
    expect(validateBates({ ...base, digits: 16 })).toBe('batesErrDigits')
    expect(validateBates({ ...base, firstPage: 0 })).toBe('batesErrRange')
    expect(validateBates({ ...base, lastPage: 0 })).toBe('batesErrRange')
    expect(validateBates({ ...base, lastPage: 5, firstPage: 6 })).toBe('batesErrRange')
  })
})

describe('buildBatesInserts', () => {
  it('stamps the bottom-right corner of a portrait page, right-aligned', () => {
    const inserts = buildBatesInserts([A4], base, measure)
    expect(inserts).toHaveLength(1)
    const insert = inserts[0]!
    expect(insert.text).toBe('ACME-000041')
    expect(insert.align).toBe('right')
    // display (567, 814) from the top-left flips to user space 28pt above the bottom
    expect(insert.origin).toEqual([595 - 28, 28])
    // right alignment travels as a negative lineXOffset of the measured width
    expect(insert.lineXOffsets).toEqual([-('ACME-000041'.length * 5)])
  })

  it('centers top stamps on the display width', () => {
    const inserts = buildBatesInserts([A4], { ...base, position: 'top-center' }, measure)
    expect(inserts[0]!.origin).toEqual([595 / 2, 842 - 28 - 10])
    expect(inserts[0]!.align).toBe('center')
    expect(inserts[0]!.lineXOffsets).toEqual([-('ACME-000041'.length * 5) / 2])
  })

  it('counts up across pages', () => {
    const inserts = buildBatesInserts([A4, A4, A4], base, measure)
    expect(inserts.map((i) => i.text)).toEqual(['ACME-000041', 'ACME-000042', 'ACME-000043'])
    expect(inserts.map((i) => i.pageIndex)).toEqual([0, 1, 2])
  })

  it('slices by the page range', () => {
    const pages = [A4, A4, A4, A4]
    expect(
      buildBatesInserts(pages, { ...base, firstPage: 2 }, measure).map((i) => i.pageIndex),
    ).toEqual([1, 2, 3])
    expect(
      buildBatesInserts(pages, { ...base, lastPage: 2 }, measure).map((i) => i.pageIndex),
    ).toEqual([0, 1])
    expect(
      buildBatesInserts(pages, { ...base, firstPage: 2, lastPage: 3 }, measure).map(
        (i) => i.pageIndex,
      ),
    ).toEqual([1, 2])
  })

  it('keeps the number upright on a rotated page by mapping the display corner back', () => {
    // /Rotate 90 displays the 595x842 page as 842x595: the page's +y axis points
    // right on screen and +x points down, so display (814, 567) is user space
    // (y_d, x_d) = (567, 814). This is viewToPdf's mapping, the one every other
    // layer uses; a hand-rolled copy used h - x_d and landed in the wrong corner.
    const rotated = { ...A4, rotate: 90 }
    const inserts = buildBatesInserts([rotated], base, measure)
    expect(inserts[0]!.rotate).toBe(90)
    expect(inserts[0]!.origin).toEqual([567, 814])
  })

  it('scales the display corner by /UserUnit, as every other insert does', () => {
    // a page with /UserUnit 2 lays out at 2x: the display corner is halved
    // back into raw PDF units before it lands in the file
    const inserts = buildBatesInserts([{ ...A4, userUnit: 2 }], base, measure)
    const [plain] = buildBatesInserts([A4], base, measure)
    expect(inserts[0]!.origin).toEqual(plain!.origin.map((v) => v / 2))
  })
})
