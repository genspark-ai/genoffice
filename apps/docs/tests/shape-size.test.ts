import { describe, expect, it } from 'vitest'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { SectionSettings, TextboxDisplay } from '@genoffice/docx-engine'
import {
  boxWithSize,
  boxesFittedToSection,
  pageGeometryKey,
  relSidePx,
} from '../src/renderer/editor/shape-size'

const LETTER = {
  pageWidth: 12240,
  pageHeight: 15840,
  marginTop: 1440,
  marginRight: 1440,
  marginBottom: 1440,
  marginLeft: 1440,
} as SectionSettings

const box = (over: Partial<TextboxDisplay> = {}): TextboxDisplay =>
  ({ paras: [], widthPx: 312, heightPx: 96, minHeightPx: 96, ...over }) as TextboxDisplay

const nodeOf = (textboxes: TextboxDisplay[]) => ({ attrs: { textboxes } }) as unknown as PmNode

describe('a shape sized as a share of the page', () => {
  it('comes to the px its share gives on the page of the section', () => {
    // half of a 6.5 in text area
    expect(relSidePx(LETTER, 'H', { from: 'margin', pct: 50 })).toBe(312)
    expect(relSidePx(LETTER, 'V', { from: 'page', pct: 10 })).toBe(106)
  })

  it('a side set without a share is absolute; the other keeps what it had', () => {
    const rel = box({
      sizeRel: { w: { from: 'margin', pct: 50 }, h: { from: 'page', pct: 10 } },
    })
    const dragged = boxWithSize(rel, { w: { px: 200.4, rel: null } })
    expect(dragged.widthPx).toBe(200)
    expect(dragged.sizeRel).toEqual({ h: { from: 'page', pct: 10 } })
    const both = boxWithSize(dragged, { h: { px: 50, rel: null } })
    expect(both.sizeRel).toBeUndefined()
    expect(both.minHeightPx).toBe(50)
    expect(
      boxWithSize(box(), { w: { px: 100, rel: { from: 'page', pct: 12.3 } } }).sizeRel,
    ).toEqual({
      w: { from: 'page', pct: 12.3 },
    })
  })

  it('follows a new page setup, and only where it states a share', () => {
    const wide = { ...LETTER, marginLeft: 720, marginRight: 720 } as SectionSettings
    const rel = box({ sizeRel: { w: { from: 'margin', pct: 50 } } })
    const plain = box()
    const fitted = boxesFittedToSection(nodeOf([plain, rel]), wide)
    expect(fitted?.[0]).toBe(plain)
    expect(fitted?.[1]).toMatchObject({ widthPx: 360, heightPx: 96, sizeRel: rel.sizeRel })
    expect(boxesFittedToSection(nodeOf([plain, rel]), LETTER)).toBeNull()
    expect(boxesFittedToSection(nodeOf([plain]), wide)).toBeNull()
  })

  it('leaves the height of a shape that is as tall as its text', () => {
    const auto = box({ heightPx: undefined, sizeRel: { h: { from: 'margin', pct: 20 } } })
    expect(boxesFittedToSection(nodeOf([auto]), LETTER)).toBeNull()
  })

  it('tells one page geometry from another', () => {
    expect(pageGeometryKey(LETTER)).toBe('12240,15840,1440,1440,1440,1440')
    expect(pageGeometryKey(null)).toBe('')
  })
})
