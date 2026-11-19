import { describe, expect, it } from 'vitest'
import type { SectionInfo } from '@genoffice/docx-engine'
import { canvasContentTopPx, measureBlocks } from '../src/renderer/pagination'

const twipsToPx = (twips: number) => (twips / 1440) * 96

const sec = (marginTop: number): SectionInfo => ({
  settings: {
    pageWidth: 11906,
    pageHeight: 16838,
    orientation: 'portrait',
    marginTop,
    marginRight: 1440,
    marginBottom: 1440,
    marginLeft: 1440,
    pageBorder: false,
    columns: 1,
  },
  startType: 'nextPage',
  firstBlockIndex: 0,
  lastBlockIndex: 0,
  sectPrXml: '',
  titlePg: false,
  headerRefs: {},
  footerRefs: {},
})

const rectAt = (top: number, height: number): DOMRect =>
  ({ top, height, bottom: top + height, left: 0, right: 100, width: 100, x: 0, y: top }) as DOMRect

describe('canvas content-area top', () => {
  // cover section with a 120-twip top margin followed by 720-twip body sections
  const sections = [sec(120), sec(720), sec(720)]
  const body = sections[sections.length - 1].settings

  it('follows the first section (the shared paper padding), not the body sectPr', () => {
    expect(canvasContentTopPx(sections, body, 0)).toBe(twipsToPx(120))
    expect(canvasContentTopPx([], body, 0)).toBe(twipsToPx(720))
  })

  it('header push-down applies to the first section too', () => {
    const headerPx = 60
    const headerDist = twipsToPx(720)
    expect(canvasContentTopPx(sections, body, headerPx)).toBe(headerDist + headerPx)
  })

  it('measured against that origin, the first block starts at virtual 0 like the clone', () => {
    const padTop = canvasContentTopPx(sections, body, 0)
    const pm = document.createElement('div')
    pm.getBoundingClientRect = () => rectAt(0, 1000)
    const p = document.createElement('p')
    p.textContent = 'x'
    p.getBoundingClientRect = () => rectAt(padTop, 20)
    pm.appendChild(p)
    expect(measureBlocks(pm, padTop, 1).blocks[0].top).toBe(0)
    // the old body-section origin put it 40px above the page top
    const bodyTop = canvasContentTopPx([], body, 0)
    expect(measureBlocks(pm, bodyTop, 1).blocks[0].top).toBe(padTop - bodyTop)
  })
})
