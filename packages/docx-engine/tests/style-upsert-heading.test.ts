import { describe, expect, it } from 'vitest'
import { pendingHeadingLevel, type StyleHeadingInfo, type StyleUpsert } from '../src/index'

const existing = new Map<string, StyleHeadingInfo>([
  ['Normal', { name: 'Normal' }],
  ['Heading1', { name: 'heading 1', basedOn: 'Normal', headingLevel: 1 }],
  ['Chapter', { name: 'Chapter', basedOn: 'Normal', headingLevel: 2 }],
  ['Sub', { name: 'Sub', basedOn: 'Heading1', headingLevel: 1, headingLevelInherited: true }],
  ['Body', { name: 'Body', basedOn: 'Heading1' }],
  ['TOCHeading', { name: 'TOC Heading', basedOn: 'Heading1', headingOutlineOff: true }],
])
const level = (pending: StyleUpsert[], id: string) =>
  pendingHeadingLevel(
    id,
    (k) => pending.find((u) => u.styleId === k),
    (k) => existing.get(k),
  )

describe('pendingHeadingLevel', () => {
  it('inherits the parent heading level the way parse does on reopen', () => {
    expect(level([{ styleId: 'Callout', basedOn: 'Heading1' }], 'Callout')).toBe(1)
    expect(level([{ styleId: 'Callout', basedOn: 'Normal' }], 'Callout')).toBeUndefined()
    expect(level([{ styleId: 'Callout' }], 'Callout')).toBeUndefined()
  })

  it('own outline level wins and null switches the heading off', () => {
    expect(
      level([{ styleId: 'Callout', basedOn: 'Heading1', pPr: { outlineLevel: 3 } }], 'Callout'),
    ).toBe(3)
    expect(
      level([{ styleId: 'Body', basedOn: 'Heading1', pPr: { outlineLevel: null } }], 'Body'),
    ).toBeUndefined()
  })

  it('follows chains through other pending styles and survives cycles', () => {
    const chain: StyleUpsert[] = [
      { styleId: 'A', basedOn: 'B' },
      { styleId: 'B', basedOn: 'Heading1' },
    ]
    expect(level(chain, 'A')).toBe(1)
    const cycle: StyleUpsert[] = [
      { styleId: 'A', basedOn: 'B' },
      { styleId: 'B', basedOn: 'A' },
    ]
    expect(level(cycle, 'A')).toBeUndefined()
  })

  it('an explicit outline-off style stays body text under any parent, like parse does', () => {
    expect(level([{ styleId: 'TOCHeading', basedOn: 'Heading1' }], 'TOCHeading')).toBeUndefined()
    expect(level([{ styleId: 'TOCHeading', basedOn: 'Chapter' }], 'TOCHeading')).toBeUndefined()
    expect(level([{ styleId: 'TOCHeading', rPr: { bold: true } }], 'TOCHeading')).toBeUndefined()
    expect(level([{ styleId: 'Note', basedOn: 'TOCHeading' }], 'Note')).toBeUndefined()
    expect(
      level(
        [{ styleId: 'TOCHeading', basedOn: 'Chapter', pPr: { outlineLevel: 3 } }],
        'TOCHeading',
      ),
    ).toBe(3)
    expect(level([{ styleId: 'TOCHeading', name: 'heading 2' }], 'TOCHeading')).toBe(2)
  })

  it('heading names and ids count as that level, like parse does', () => {
    expect(level([{ styleId: 'Heading3', basedOn: 'Normal' }], 'Heading3')).toBe(3)
    expect(level([{ styleId: 'Chap', name: 'heading 4', basedOn: 'Normal' }], 'Chap')).toBe(4)
    expect(level([{ styleId: 'Heading1', basedOn: 'Normal' }], 'Heading1')).toBe(1)
  })

  it('an existing style keeps its level; only an inherited level follows a new parent', () => {
    expect(level([{ styleId: 'Sub', rPr: { bold: true } }], 'Sub')).toBe(1)
    expect(level([{ styleId: 'Sub', basedOn: 'Normal' }], 'Sub')).toBeUndefined()
    expect(level([{ styleId: 'Sub', basedOn: 'Chapter' }], 'Sub')).toBe(2)
    expect(level([{ styleId: 'Chapter', basedOn: 'Normal' }], 'Chapter')).toBe(2)
    expect(level([{ styleId: 'Chapter', basedOn: null }], 'Chapter')).toBe(2)
    expect(level([{ styleId: 'Body', rPr: { bold: true } }], 'Body')).toBeUndefined()
    expect(level([], 'Sub')).toBe(1)
  })
})
