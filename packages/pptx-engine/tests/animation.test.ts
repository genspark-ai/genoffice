import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  animClassOf,
  applyHeaderFooter,
  buildTimingXml,
  deleteElement,
  elementSpid,
  pruneTimingForSpids,
  getSlideAnimations,
  openPptx,
  readSlideTimingXml,
  savePptx,
  setSlideAnimations,
  setSlideTransition,
  type SlideAnimation,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

describe('shape animations (<p:timing>)', () => {
  it('elementSpid extracts the cNvPr id of each element', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const spids = slide.elements.map((el) => elementSpid(el))
    expect(spids.every((s) => typeof s === 'number' && s! > 0)).toBe(true)
    expect(new Set(spids).size).toBe(spids.length) // unique within the slide
  })

  it('classifies effects', () => {
    expect(animClassOf('fade')).toBe('entrance')
    expect(animClassOf('spin')).toBe('emphasis')
    expect(animClassOf('flyOut')).toBe('exit')
  })

  it('builds a PowerPoint-shaped timing tree', () => {
    const xml = buildTimingXml([
      { spid: 4, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ])
    expect(xml).toContain('nodeType="tmRoot"')
    expect(xml).toContain('nodeType="mainSeq"')
    expect(xml).toContain('presetID="10" presetClass="entr"')
    expect(xml).toContain('<p:spTgt spid="4"/>')
    expect(xml).toContain('<p:bldP spid="4" grpId="0"/>')
    // cTn ids are globally unique
    const ids = [...xml.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => Number(m[1]))
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('round-trips a mixed animation list through save/reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const spids = slide.elements.map((el) => elementSpid(el)!)
    const anims: SlideAnimation[] = [
      { spid: spids[0]!, effect: 'appear', trigger: 'onClick', durationMs: 0, delayMs: 0 },
      { spid: spids[1]!, effect: 'flyIn', trigger: 'withPrev', durationMs: 800, delayMs: 200 },
      { spid: spids[0]!, effect: 'spin', trigger: 'afterPrev', durationMs: 1000, delayMs: 0 },
      { spid: spids[1]!, effect: 'fadeOut', trigger: 'onClick', durationMs: 500, delayMs: 100 },
    ]
    setSlideAnimations(slide, anims)
    expect(getSlideAnimations(slide)).toEqual(anims)

    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideAnimations(reopened.deck.slides[0]!)).toEqual(anims)
    // Other slides are unaffected
    expect(getSlideAnimations(reopened.deck.slides[1]!)).toEqual([])
  })

  it('replaces existing timing instead of stacking, and can clear', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const spid = elementSpid(slide.elements[0]!)!
    setSlideAnimations(slide, [
      { spid, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ])
    setSlideAnimations(slide, [
      { spid, effect: 'zoom', trigger: 'onClick', durationMs: 700, delayMs: 0 },
    ])
    expect(slide.bodySuffix.match(/<p:timing>/g)!.length).toBe(1)
    expect(getSlideAnimations(slide)).toEqual([
      { spid, effect: 'zoom', trigger: 'onClick', durationMs: 700, delayMs: 0 },
    ])

    setSlideAnimations(slide, [])
    expect(slide.bodySuffix).not.toContain('<p:timing')
    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideAnimations(reopened.deck.slides[0]!)).toEqual([])
  })

  it('coexists with a slide transition (transition before timing)', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const spid = elementSpid(slide.elements[0]!)!
    setSlideTransition(slide, 'fade')
    setSlideAnimations(slide, [
      { spid, effect: 'wipe', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ])
    const reopened = await openPptx(await savePptx(opened))
    const suffix = reopened.deck.slides[0]!.bodySuffix
    expect(suffix.indexOf('<p:transition')).toBeGreaterThanOrEqual(0)
    expect(suffix.indexOf('<p:transition')).toBeLessThan(suffix.indexOf('<p:timing'))
    expect(getSlideAnimations(reopened.deck.slides[0]!)).toHaveLength(1)
  })

  it('reads foreign presets with best-effort fallback', () => {
    // Unmodeled presetIDs (e.g. entr box 4) degrade to a similar effect of the same class
    const xml = buildTimingXml([
      { spid: 7, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ]).replace('presetID="10"', 'presetID="4"')
    const anims = readSlideTimingXml(`</p:cSld>${xml}</p:sld>`)
    expect(anims).toHaveLength(1)
    expect(anims[0]).toMatchObject({
      spid: 7,
      effect: 'fade',
      trigger: 'onClick',
      durationMs: 500,
      delayMs: 0,
      preset: { id: 4, cls: 'entr', sub: 0 },
    })
    // the original par is kept so a rewrite never downgrades it
    expect(anims[0]!.presetXml).toMatch(/^<p:par>.*presetID="4".*<\/p:par>$/)
  })
})

describe('extended effects + motion paths', () => {
  it('classifies the new effects', () => {
    expect(animClassOf('bounce')).toBe('entrance')
    expect(animClassOf('flipIn')).toBe('entrance')
    expect(animClassOf('teeter')).toBe('emphasis')
    expect(animClassOf('shrink')).toBe('exit')
    expect(animClassOf('motionPath')).toBe('path')
  })

  it('round-trips every new effect kind through save/reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const spid = elementSpid(slide.elements[0]!)!
    const anims: SlideAnimation[] = (
      ['wipeDown', 'splitIn', 'bounce', 'flipIn', 'teeter', 'wipeOut', 'shrink'] as const
    ).map((effect, i) => ({
      spid,
      effect,
      trigger: i === 0 ? ('onClick' as const) : ('afterPrev' as const),
      durationMs: 500 + i * 100,
      delayMs: 0,
    }))
    setSlideAnimations(slide, anims)
    expect(getSlideAnimations(slide)).toEqual(anims)
    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideAnimations(reopened.deck.slides[0]!)).toEqual(anims)
  })

  it('distinguishes wipe directions by presetSubtype on read-back', () => {
    const anims: SlideAnimation[] = [
      { spid: 3, effect: 'wipe', trigger: 'onClick', durationMs: 500, delayMs: 0 },
      { spid: 3, effect: 'wipeDown', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ]
    const xml = buildTimingXml(anims)
    expect(xml).toContain('filter="wipe(up)"')
    expect(xml).toContain('filter="wipe(down)"')
    expect(readSlideTimingXml(`</p:cSld>${xml}</p:sld>`)).toEqual(anims)
  })

  it('writes motion paths as p:animMotion with a relative path and E terminator', () => {
    const xml = buildTimingXml([
      {
        spid: 5,
        effect: 'motionPath',
        trigger: 'onClick',
        durationMs: 2000,
        delayMs: 0,
        motionPath: 'M 0 0 L 0.3 0',
      },
    ])
    expect(xml).toContain('presetID="0" presetClass="path"')
    expect(xml).toContain(
      '<p:animMotion origin="layout" path="M 0 0 L 0.3 0 E" pathEditMode="relative">',
    )
    expect(xml).toContain('<p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName>')
  })

  it('round-trips motion paths (line + bezier circle) through save/reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const spids = slide.elements.map((el) => elementSpid(el)!)
    const circle =
      'M 0 0 C 0.069 0 0.125 0.056 0.125 0.125 C 0.125 0.194 0.069 0.25 0 0.25 ' +
      'C -0.069 0.25 -0.125 0.194 -0.125 0.125 C -0.125 0.056 -0.069 0 0 0 Z'
    const anims: SlideAnimation[] = [
      {
        spid: spids[0]!,
        effect: 'motionPath',
        trigger: 'onClick',
        durationMs: 2000,
        delayMs: 0,
        motionPath: 'M 0 0 L 0.25 0.25',
      },
      {
        spid: spids[1]!,
        effect: 'motionPath',
        trigger: 'withPrev',
        durationMs: 3000,
        delayMs: 100,
        motionPath: circle,
      },
      // Mixing new and old: existing effects are unaffected
      { spid: spids[0]!, effect: 'fadeOut', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ]
    setSlideAnimations(slide, anims)
    expect(getSlideAnimations(slide)).toEqual(anims)
    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideAnimations(reopened.deck.slides[0]!)).toEqual(anims)
  })

  it('falls back to the default path when motionPath is missing (legacy data safety)', () => {
    const xml = buildTimingXml([
      { spid: 9, effect: 'motionPath', trigger: 'onClick', durationMs: 2000, delayMs: 0 },
    ])
    expect(xml).toContain('path="M 0 0 L 0.25 0 E"')
    const back = readSlideTimingXml(`</p:cSld>${xml}</p:sld>`)
    expect(back[0]!.motionPath).toBe('M 0 0 L 0.25 0')
  })
})

describe('pruneTimingForSpids (element removal must not leave dangling spTgt refs)', () => {
  async function seeded() {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const spids = slide.elements.map((el) => elementSpid(el)!)
    setSlideAnimations(slide, [
      { spid: spids[0]!, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
      { spid: spids[1]!, effect: 'flyIn', trigger: 'withPrev', durationMs: 800, delayMs: 0 },
    ])
    return { opened, slide, spids }
  }

  it('removes only the targeted effect, keeps the rest byte-preserved', async () => {
    const { slide, spids } = await seeded()
    expect(pruneTimingForSpids(slide, new Set([spids[0]!]))).toBe(true)
    const left = getSlideAnimations(slide)
    expect(left.map((a) => a.spid)).toEqual([spids[1]!])
    expect(slide.bodySuffix).not.toContain(`<p:spTgt spid="${spids[0]}"`)
  })

  it('drops the whole <p:timing> when no target survives', async () => {
    const { slide, spids } = await seeded()
    pruneTimingForSpids(slide, new Set([spids[0]!, spids[1]!]))
    expect(slide.bodySuffix).not.toContain('<p:timing>')
  })

  it('a trigger reference removes only its interactiveSeq, never the whole timing', async () => {
    const { slide, spids } = await seeded()
    const TRIGGER_SEQ =
      '<p:seq concurrent="1" nextAc="none"><p:cTn id="900" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">' +
      '<p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="99"/></p:tgtEl></p:cond></p:stCondLst>' +
      '<p:childTnLst><p:par><p:cTn id="901" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">' +
      '<p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
      `<p:childTnLst><p:set><p:cBhvr><p:cTn id="902" dur="1"/><p:tgtEl><p:spTgt spid="${spids[1]}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst>` +
      '</p:cTn></p:par></p:childTnLst></p:cTn></p:seq>'
    const at = slide.bodySuffix.indexOf('</p:seq>') + '</p:seq>'.length
    slide.bodySuffix = slide.bodySuffix.slice(0, at) + TRIGGER_SEQ + slide.bodySuffix.slice(at)

    // Deleting the trigger shape (99): its seq goes — including the triggered
    // effect on a surviving shape — but the main sequence stays intact
    expect(pruneTimingForSpids(slide, new Set([99]))).toBe(true)
    expect(slide.bodySuffix).toContain('<p:timing>')
    expect(slide.bodySuffix).not.toContain('interactiveSeq')
    expect(getSlideAnimations(slide).map((a) => a.spid)).toEqual([spids[0]!, spids[1]!])
  })

  it('pruning an effect target inside an interactiveSeq removes only that effect par', async () => {
    const { slide, spids } = await seeded()
    // interactiveSeq triggered by shape 99 with TWO effects: one on spids[0]
    // (about to be deleted) and one on spids[1] (must survive). The effect's
    // own closed delay cond must not be mistaken for a trigger wrapper.
    const effectPar = (id: number, spid: number) =>
      `<p:par><p:cTn id="${id}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">` +
      '<p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
      `<p:childTnLst><p:set><p:cBhvr><p:cTn id="${id + 1}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>` +
      `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
      '<p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par>'
    const TRIGGER_SEQ =
      '<p:seq concurrent="1" nextAc="none"><p:cTn id="900" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">' +
      '<p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="99"/></p:tgtEl></p:cond></p:stCondLst>' +
      `<p:childTnLst>${effectPar(901, spids[0]!)}${effectPar(910, spids[1]!)}</p:childTnLst></p:cTn></p:seq>`
    const at = slide.bodySuffix.indexOf('</p:seq>') + '</p:seq>'.length
    slide.bodySuffix = slide.bodySuffix.slice(0, at) + TRIGGER_SEQ + slide.bodySuffix.slice(at)

    pruneTimingForSpids(slide, new Set([spids[0]!]))
    expect(slide.bodySuffix).toContain('interactiveSeq') // seq survives
    const spidsLeft = [...slide.bodySuffix.matchAll(/<p:spTgt spid="(\d+)"/g)].map((m) =>
      Number(m[1]),
    )
    expect(spidsLeft).not.toContain(spids[0]!)
    expect(spidsLeft.filter((x) => x === spids[1]!).length).toBeGreaterThanOrEqual(2) // mainSeq + trigger effect
    expect(spidsLeft).toContain(99)
  })

  it('prunes the matching bldP build refs and keeps the survivors', async () => {
    const { slide, spids } = await seeded()
    expect(slide.bodySuffix).toContain(`<p:bldP spid="${spids[0]}"`)
    pruneTimingForSpids(slide, new Set([spids[0]!]))
    expect(slide.bodySuffix).not.toContain(`<p:bldP spid="${spids[0]}"`)
    expect(slide.bodySuffix).toContain(`<p:bldP spid="${spids[1]}"`)
  })

  it('a wrapper shared with a surviving target is left alone (no collateral destruction)', () => {
    // One par carrying behaviors for two shapes: pruning one of them must not
    // take the shared par (and the survivor's animation) with it
    const sharedPar =
      '<p:timing><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst>' +
      '<p:par><p:cTn id="2"><p:childTnLst>' +
      '<p:set><p:cBhvr><p:cTn id="3" dur="1"/><p:tgtEl><p:spTgt spid="5"/></p:tgtEl></p:cBhvr></p:set>' +
      '<p:set><p:cBhvr><p:cTn id="4" dur="1"/><p:tgtEl><p:spTgt spid="6"/></p:tgtEl></p:cBhvr></p:set>' +
      '</p:childTnLst></p:cTn></p:par>' +
      '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
    const slide = {
      bodySuffix: `${sharedPar}</p:sld>`,
      structureDirty: false,
    } as never as Parameters<typeof pruneTimingForSpids>[0]
    expect(pruneTimingForSpids(slide, new Set([5]))).toBe(false)
    expect((slide as { bodySuffix: string }).bodySuffix).toContain('<p:spTgt spid="6"')
    expect((slide as { bodySuffix: string }).bodySuffix).toContain('<p:spTgt spid="5"')
  })

  it('deleteElement prunes the deleted shape from the timing', async () => {
    const { opened, slide, spids } = await seeded()
    const victim = slide.elements.find((el) => elementSpid(el) === spids[0])!
    expect(deleteElement(opened, slide, victim.id)).toBe(true)
    expect(slide.bodySuffix).not.toContain(`<p:spTgt spid="${spids[0]}"`)
    expect(getSlideAnimations(slide).map((a) => a.spid)).toEqual([spids[1]!])
    // survives save→reopen
    const saved = await openPptx(await savePptx(opened))
    const spidsAfter = new Set(
      [...saved.deck.slides[0]!.bodySuffix.matchAll(/<p:spTgt spid="(\d+)"/g)].map((m) =>
        Number(m[1]),
      ),
    )
    expect(spidsAfter.has(spids[0]!)).toBe(false)
    expect(spidsAfter.has(spids[1]!)).toBe(true)
  })

  it('applyHeaderFooter prunes animations of the placeholders it replaces', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    // First install footers, animate one, then re-apply: the old placeholder
    // is removed and its animation must go with it
    expect(applyHeaderFooter(opened, { footer: 'v1', slideNum: true })).toBe(true)
    const slide = opened.deck.slides[0]!
    const ftr = slide.elements.find((el) => el.placeholder === 'ftr')!
    const spid = elementSpid(ftr)!
    setSlideAnimations(slide, [
      { spid, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ])
    expect(applyHeaderFooter(opened, { footer: 'v2', slideNum: true })).toBe(true)
    const fresh = opened.deck.slides[0]!
    const ids = new Set(fresh.elements.map((el) => elementSpid(el)))
    for (const m of fresh.bodySuffix.matchAll(/<p:spTgt spid="(\d+)"/g)) {
      expect(ids.has(Number(m[1]))).toBe(true)
    }
    expect(fresh.bodySuffix).not.toContain(`<p:spTgt spid="${spid}"`)
  })
})

describe('directions and verbatim presets', () => {
  const fx1 = () => openPptx(fx('01_standard_business.pptx'))

  it('writes fly/wipe sides as PowerPoint subtypes and reads them back', async () => {
    const opened = await fx1()
    const slide = opened.deck.slides[0]!
    const spid = elementSpid(slide.elements[0]!)!
    const anims: SlideAnimation[] = [
      { spid, effect: 'flyIn', trigger: 'onClick', durationMs: 500, delayMs: 0, direction: 'left' },
      { spid, effect: 'wipe', trigger: 'onClick', durationMs: 500, delayMs: 0, direction: 'right' },
      { spid, effect: 'flyOut', trigger: 'onClick', durationMs: 500, delayMs: 0, direction: 'top' },
      { spid, effect: 'wipeOut', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ]
    setSlideAnimations(slide, anims)
    const xml = slide.bodySuffix
    expect(xml).toContain('presetID="2" presetClass="entr" presetSubtype="8"')
    expect(xml).toContain('filter="wipe(left)"')
    expect(xml).toContain('presetID="22" presetClass="entr" presetSubtype="2"')
    expect(xml).toContain('0-#ppt_h/2')
    const reopened = await openPptx(await savePptx(opened))
    const back = getSlideAnimations(reopened.deck.slides[0]!)
    expect(back.map((a) => [a.effect, a.direction])).toEqual([
      ['flyIn', 'left'],
      ['wipe', 'right'],
      ['flyOut', 'top'],
      ['wipeOut', undefined],
    ])
    expect(back[0]!.preset).toBeUndefined()
  })

  it('keeps an unmodeled effect byte-for-byte across a full rebuild', async () => {
    const opened = await fx1()
    const slide = opened.deck.slides[0]!
    const spid = elementSpid(slide.elements[0]!)!
    const other = elementSpid(slide.elements[1]!)!
    const foreign =
      `<p:par><p:cTn id="7" presetID="5" presetClass="entr" presetSubtype="10" fill="hold" grpId="0" nodeType="clickEffect">` +
      '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
      `<p:animEffect transition="in" filter="randombar(horizontal)"><p:cBhvr><p:cTn id="8" dur="500"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect>` +
      '</p:childTnLst></p:cTn></p:par>'
    const timing =
      '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
      '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' +
      '<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>' +
      '<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
      foreign +
      '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>' +
      '</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
    slide.bodySuffix = slide.bodySuffix.replace('</p:sld>', `${timing}</p:sld>`)
    const read = readSlideTimingXml(slide.bodySuffix)
    expect(read).toHaveLength(1)
    expect(read[0]!.presetXml).toBe(foreign)
    expect(read[0]!.effect).toBe('fade')

    // deleting/reordering forces the full rebuild path; the foreign par must survive it
    setSlideAnimations(slide, [
      { spid: other, effect: 'appear', trigger: 'onClick', durationMs: 0, delayMs: 0 },
      ...read,
    ])
    const out = slide.bodySuffix
    expect(out).toContain('filter="randombar(horizontal)"')
    expect(out).toContain('presetID="5" presetClass="entr" presetSubtype="10"')
    const ids = [...out.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => Number(m[1]))
    expect(new Set(ids).size).toBe(ids.length)
    const again = readSlideTimingXml(out)
    expect(again.map((a) => a.effect)).toEqual(['appear', 'fade'])
    expect(again[1]!.presetXml).toContain('randombar')
  })
})
