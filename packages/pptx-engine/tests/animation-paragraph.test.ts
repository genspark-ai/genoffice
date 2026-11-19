/**
 * Per-paragraph animation (pgRg + build="p"): write/read round-trip, mixing with whole-shape animations, incremental path.
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  buildTimingXml,
  elementSpid,
  getSlideAnimations,
  openPptx,
  patchSlideTimingIncrementalXml,
  readSlideTimingXml,
  savePptx,
  setSlideAnimations,
  type SlideAnimation,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const para = (spid: number, paragraph: number, o?: Partial<SlideAnimation>): SlideAnimation => ({
  spid,
  effect: 'fade',
  trigger: 'onClick',
  durationMs: 500,
  delayMs: 0,
  paragraph,
  ...o,
})

describe('paragraph builds (pgRg + build="p")', () => {
  it('two-paragraph fade: pgRg on every behavior target + a single build="p" bldP', () => {
    const anims = [para(4, 0), para(4, 1)]
    const xml = buildTimingXml(anims)
    // fade = set + animEffect; the two animations have 4 behavior targets total, all with txEl/pgRg
    expect(xml.match(/<p:pgRg st="0" end="0"\/>/g)!.length).toBe(2)
    expect(xml.match(/<p:pgRg st="1" end="1"\/>/g)!.length).toBe(2)
    expect(xml.match(/<p:spTgt spid="4"><p:txEl><p:pgRg/g)!.length).toBe(4)
    // bldP: one build="p" for the whole shape; no longer writes a plain bldP per grpId
    expect(xml.match(/<p:bldP\b/g)!.length).toBe(1)
    expect(xml).toContain('<p:bldP spid="4" grpId="0" uiExpand="1" build="p"/>')
    expect(readSlideTimingXml(`</p:cSld>${xml}</p:sld>`)).toEqual(anims)
  })

  it('round-trips through save/reopen, mixed with whole-shape animations', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const spids = slide.elements.map((el) => elementSpid(el)!)
    const anims: SlideAnimation[] = [
      para(spids[0]!, 0),
      para(spids[0]!, 1, { trigger: 'afterPrev', durationMs: 800 }),
      { spid: spids[1]!, effect: 'flyIn', trigger: 'onClick', durationMs: 700, delayMs: 100 },
      para(spids[0]!, 2, { effect: 'wipe' }),
    ]
    setSlideAnimations(slide, anims)
    expect(getSlideAnimations(slide)).toEqual(anims)
    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideAnimations(reopened.deck.slides[0]!)).toEqual(anims)
    // Paragraph shape build="p" and whole-shape plain bldP coexist
    const suffix = reopened.deck.slides[0]!.bodySuffix
    expect(suffix).toContain(`<p:bldP spid="${spids[0]}" grpId="0" uiExpand="1" build="p"/>`)
    expect(suffix).toContain(`<p:bldP spid="${spids[1]}" grpId="0"/>`)
  })

  it('paragraph targets survive behaviors of movement/scale/motion effects', () => {
    const anims: SlideAnimation[] = [
      para(6, 0, { effect: 'flyIn' }),
      para(6, 1, { effect: 'zoom' }),
      para(6, 2, { effect: 'motionPath', motionPath: 'M 0 0 L 0.25 0' }),
    ]
    const xml = buildTimingXml(anims)
    // Each behavior's target (set/anim/animScale/animMotion) carries the paragraph
    expect(xml.match(/<p:tgtEl><p:spTgt spid="6"\/><\/p:tgtEl>/g)).toBeNull()
    expect(readSlideTimingXml(`</p:cSld>${xml}</p:sld>`)).toEqual(anims)
  })

  it('incremental append of paragraph animations keeps bytes and adds one build="p" bldP', () => {
    const body = `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${buildTimingXml([para(4, 0)])}</p:sld>`
    const out = patchSlideTimingIncrementalXml(body, [para(4, 0), para(4, 1), para(4, 2)])
    expect(out).not.toBeNull()
    expect(readSlideTimingXml(out!)).toEqual([para(4, 0), para(4, 1), para(4, 2)])
    // Appending does not duplicate the build="p" bldP
    expect(out!.match(/<p:bldP\b/g)!.length).toBe(1)
  })

  it('incremental tail edit of a paragraph animation patches in place', () => {
    const anims = [para(4, 0), para(4, 1)]
    const body = `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${buildTimingXml(anims)}</p:sld>`
    const edited = [para(4, 0), para(4, 1, { durationMs: 1500, delayMs: 200 })]
    const out = patchSlideTimingIncrementalXml(body, edited)
    expect(out).not.toBeNull()
    expect(readSlideTimingXml(out!)).toEqual(edited)
  })

  it('falls back to full rebuild when paragraph shape changes bldP form', () => {
    const whole: SlideAnimation = { spid: 4, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 }
    const wholeBody = `</p:cSld>${buildTimingXml([whole])}</p:sld>`
    // Appending a paragraph animation to a whole-shape spid → needs a switch to build="p"; falls back
    expect(patchSlideTimingIncrementalXml(wholeBody, [whole, para(4, 0)])).toBeNull()
    // Tail edit changing the paragraph → falls back
    const paraBody = `</p:cSld>${buildTimingXml([para(4, 0), para(4, 1)])}</p:sld>`
    expect(patchSlideTimingIncrementalXml(paraBody, [para(4, 0), para(4, 2)])).toBeNull()
  })
})
