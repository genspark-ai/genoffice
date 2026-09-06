import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import type { HfPartInfo, SectionSettings } from '@genoffice/docx-engine'
import { hfFromPart } from '../src/renderer/doc-state'
import { hfHasVisibleContent } from '../src/renderer/editor/hf-dom'
import { estimateHfHeight } from '../src/renderer/line-metrics'
import { effectiveBottomPx } from '../src/renderer/pagination-sections'

const twipsToPx = (twips: number) => (twips / 1440) * 96

// A4, 1022-twip margins, footer 706 twips from the edge (316 twips of slack)
const settings: SectionSettings = {
  pageWidth: 11906,
  pageHeight: 16838,
  orientation: 'portrait',
  marginTop: 1022,
  marginRight: 1411,
  marginBottom: 1022,
  marginLeft: 1699,
  footerDist: 706,
  pageBorder: false,
  columns: 1,
}

// footer part made of two blank paragraphs: 11pt with line 288 / after 200,
// then the docDefaults 11pt line 259 / after 160
const blankFooter: HfPartInfo = {
  text: '',
  hasPageNumber: false,
  paras: [
    { runs: [], lineRule: 'auto', lineRawTwips: 288, spaceAfter: 200 },
    { runs: [], lineRule: 'auto', lineRawTwips: 259, spaceAfter: 160 },
  ],
}

describe('all-blank footer part still pushes the body bottom up (Word)', () => {
  it('keeps the blank paragraphs as a footer value without any visible content', () => {
    const value = hfFromPart(blankFooter)
    expect(value).not.toBeNull()
    expect(value!.paras).toHaveLength(2)
    expect(hfHasVisibleContent(value)).toBe(false)
  })

  it('reserves the blank lines so footerDist + height exceeds the bottom margin', () => {
    const value = hfFromPart(blankFooter)!
    const contentW = twipsToPx(settings.pageWidth - settings.marginLeft - settings.marginRight)
    const h = estimateHfHeight(value, contentW)
    // two 11pt lines plus 18pt of paragraph spacing: well over the 316-twip slack
    expect(h).toBeGreaterThan(twipsToPx(700))
    expect(effectiveBottomPx(settings, h)).toBeGreaterThan(twipsToPx(settings.marginBottom))
    expect(effectiveBottomPx(settings, 0)).toBe(twipsToPx(settings.marginBottom))
  })
})

describe('direct w:contextualSpacing on unstyled list items', () => {
  const css = readFileSync(
    join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
    'utf8',
  )
  it('covers .doc-li blocks, not only <p>', () => {
    expect(css).toContain(
      '.doc-page :is(p, .doc-li).ctx-sp:not([data-style]):has(+ :is(p, .doc-li):not([data-style]))',
    )
    expect(css).toContain(
      '.doc-page :is(p, .doc-li):not([data-style]) + :is(p, .doc-li).ctx-sp:not([data-style])',
    )
  })
})
