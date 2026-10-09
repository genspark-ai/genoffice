import { describe, expect, it } from 'vitest'
import { detectFormFields, nameFromLabel } from '../src/renderer/form-detect'
import type { DetectTextItem } from '../src/renderer/form-detect'

const item = (str: string, x: number, y: number, w: number, h = 10): DetectTextItem => ({
  str,
  x,
  y,
  w,
  h,
})

const base = {
  pageIndex: 0,
  pageBox: [0, 0, 612, 792] as [number, number, number, number],
  occupied: [] as [number, number, number, number][],
  taken: new Set<string>(),
}

describe('detectFormFields', () => {
  it('turns underscore runs, label gaps, boxes, underlines and check glyphs into fields', () => {
    const fields = detectFormFields({
      ...base,
      items: [
        item('Full name: ', 72, 700, 60),
        item('____________', 132, 700, 120),
        item('Date:', 72, 660, 30),
        item('Signature:', 72, 560, 55),
        item('[ ]', 72, 500, 12),
        item('I agree', 90, 500, 40),
        item('Country', 300, 620, 40),
      ],
      shapes: [
        [130, 548, 330, 549], // underline after Signature:
        [300, 590, 500, 612], // empty box under "Country"
        [72, 400, 500, 401.2], // underline with nothing near: still a text field
      ],
    })
    const byName = Object.fromEntries(fields.map((f) => [f.name, f]))
    expect(Object.keys(byName).sort()).toEqual([
      'country',
      'date',
      'field',
      'full_name',
      'i_agree',
      'signature',
    ])
    expect(byName.full_name!.kind).toBe('text')
    expect(byName.full_name!.rect[0]).toBeCloseTo(132, 0)
    expect(byName.date!.dateFormat).toBe('yyyy-mm-dd')
    expect(byName.date!.rect[2]).toBeCloseTo(612 - 36 - 4, 0)
    expect(byName.signature!.kind).toBe('signature')
    expect(byName.signature!.rect[3] - byName.signature!.rect[1]).toBeGreaterThanOrEqual(36)
    expect(byName.i_agree!.kind).toBe('checkbox')
    expect(byName.country!.rect).toEqual([301, 591, 499, 611])
    // Reading order: top of the page first
    expect(fields[0]!.name).toBe('full_name')
  })

  it('skips boxes that already hold text or a widget and keeps names unique', () => {
    const fields = detectFormFields({
      ...base,
      items: [item('Total', 110, 304, 30), item('Name:', 72, 700, 30), item('Name:', 72, 650, 30)],
      shapes: [
        [100, 300, 300, 320],
        [100, 200, 300, 220],
      ],
      occupied: [[100, 200, 300, 220]],
      taken: new Set(['name']),
    })
    expect(fields.map((f) => f.name)).toEqual(['name_1', 'name_2'])
  })

  it('keeps a signature grown from a line at the page bottom on the page and off its neighbours', () => {
    const fields = detectFormFields({
      ...base,
      items: [item('Signature:', 72, 20, 55), item('Date:', 300, 20, 30)],
      shapes: [[130, 16, 290, 17]],
    })
    const sig = fields.find((f) => f.kind === 'signature')!
    expect(sig.rect[1]).toBeGreaterThanOrEqual(0)
    expect(sig.rect[3] - sig.rect[1]).toBeGreaterThanOrEqual(36)
    for (const f of fields) {
      if (f === sig) continue
      const overlap = Math.min(sig.rect[2], f.rect[2]) - Math.max(sig.rect[0], f.rect[0])
      expect(
        overlap <= 0 || Math.min(sig.rect[3], f.rect[3]) - Math.max(sig.rect[1], f.rect[1]) <= 0,
      ).toBe(true)
    }
  })

  it('derives identifiers from labels', () => {
    expect(nameFromLabel('Full name:')).toBe('full_name')
    expect(nameFromLabel('E-mail address *')).toBe('e_mail_address')
    expect(nameFromLabel('\u59d3\u540d\uff1a')).toBe('')
  })
})
