import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  getSlideHidden,
  openPptx,
  patchSlideShowMasterSpXml,
  savePptx,
  setSlideHidden,
} from '../src/index'
import { parseSlide } from '../src/parse'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

describe('hidden slide (<p:sld show="0">)', () => {
  it('defaults to visible', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    for (const s of opened.deck.slides) expect(getSlideHidden(s)).toBe(false)
  })

  it('hides a slide, survives save/reopen, and can be un-hidden', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    setSlideHidden(slide, true)
    expect(getSlideHidden(slide)).toBe(true)

    // Root element carries show="0", patched only onto the <p:sld> opening tag
    const reopened = await openPptx(await savePptx(opened))
    const xml = reopened.deck.slides[0]!.originalXml
    const sldOpen = /<p:sld\b[^>]*>/.exec(xml)![0]
    expect(sldOpen).toContain('show="0"')
    expect(getSlideHidden(reopened.deck.slides[0]!)).toBe(true)
    expect(getSlideHidden(reopened.deck.slides[1]!)).toBe(false)

    // Unhide: attribute is removed cleanly
    setSlideHidden(reopened.deck.slides[0]!, false)
    expect(getSlideHidden(reopened.deck.slides[0]!)).toBe(false)
    const again = await openPptx(await savePptx(reopened))
    expect(/<p:sld\b[^>]*>/.exec(again.deck.slides[0]!.originalXml)![0]).not.toContain('show=')
  })

  it('hide is idempotent and does not stack attributes', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    setSlideHidden(slide, true)
    setSlideHidden(slide, true)
    const sldOpen = /<p:sld\b[^>]*>/.exec(slide.bodyPrefix)![0]
    expect(sldOpen.match(/show="0"/g)!.length).toBe(1)
  })
})

describe('showMasterSp quote variants (<p:sld showMasterSp>)', () => {
  const slideWithFlag = (flag: string) =>
    '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"' +
    `${flag}><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>`

  it('parses single-quoted showMasterSp as hidden', () => {
    for (const flag of [` showMasterSp="0"`, ` showMasterSp='0'`, ` showMasterSp='false'`]) {
      const slide = parseSlide({
        path: 'ppt/slides/slide1.xml',
        slideXml: slideWithFlag(flag),
        ctx: {},
      })
      expect(slide.masterSpHidden).toBe(true)
    }
    const shown = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWithFlag(''),
      ctx: {},
    })
    expect(shown.masterSpHidden).toBeUndefined()
  })

  it('toggling visibility strips a single-quoted flag without duplicating', () => {
    const prefix = "<p:sld showMasterSp='0'><p:cSld>"
    expect(patchSlideShowMasterSpXml(prefix, false)).toBe('<p:sld><p:cSld>')
    expect(patchSlideShowMasterSpXml(prefix, true)).toBe('<p:sld showMasterSp="0"><p:cSld>')
  })
})
