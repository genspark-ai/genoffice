import { describe, it, expect } from 'vitest'
import {
  animClassOf,
  buildTimingXml,
  isMediaEffect,
  patchSlideTimingIncrementalXml,
  patchSlideTimingXml,
  readSlideTimingXml,
  type SlideAnimation,
} from '../src/index'

/** <p:timing> PowerPoint writes for a video whose Start = "In Click Sequence". */
function pptVideoTiming(cmd = 'playFrom(0.0)', presetId = 2): string {
  return (
    '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
    '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' +
    '<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>' +
    '<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
    `<p:par><p:cTn id="5" presetID="${presetId}" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:cmd type="call" cmd="${cmd}"><p:cBhvr><p:cTn id="6" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="4"/></p:tgtEl></p:cBhvr></p:cmd>` +
    '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>' +
    '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>' +
    '<p:video><p:cMediaNode vol="80000"><p:cTn id="7" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:endCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn><p:tgtEl><p:spTgt spid="4"/></p:tgtEl></p:cMediaNode></p:video>' +
    '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
  )
}
const wrap = (timing: string) => `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${timing}</p:sld>`

describe('media command animations', () => {
  it('classifies play/pause/stop as media effects', () => {
    for (const e of ['mediaPlay', 'mediaPause', 'mediaStop'] as const) {
      expect(isMediaEffect(e)).toBe(true)
      expect(animClassOf(e)).toBe('media')
    }
    expect(isMediaEffect('fade')).toBe(false)
  })

  it('reads a PowerPoint click-sequence video as mediaPlay, not as an entrance', () => {
    const [a, ...rest] = readSlideTimingXml(wrap(pptVideoTiming()))
    expect(rest).toEqual([])
    expect(a).toMatchObject({ spid: 4, effect: 'mediaPlay', trigger: 'onClick', durationMs: 0 })
    expect(a!.presetXml).toBeUndefined()
  })

  it('maps the command string over the presetID', () => {
    expect(readSlideTimingXml(wrap(pptVideoTiming('togglePause', 1)))[0]!.effect).toBe('mediaPause')
    expect(readSlideTimingXml(wrap(pptVideoTiming('stop', 3)))[0]!.effect).toBe('mediaStop')
    // Mislabeled presetID still follows the command
    expect(readSlideTimingXml(wrap(pptVideoTiming('stop', 2)))[0]!.effect).toBe('mediaStop')
  })

  it('keeps a playFrom offset verbatim (no model for bookmark/trim starts)', () => {
    const [a] = readSlideTimingXml(wrap(pptVideoTiming('playFrom(3.5)')))
    expect(a!.effect).toBe('mediaPlay')
    expect(a!.presetXml).toContain('playFrom(3.5)')
    const out = patchSlideTimingXml(wrap(pptVideoTiming('playFrom(3.5)')), [a!])
    expect(out).toContain('cmd="playFrom(3.5)"')
  })

  it('rewrites the main sequence with a media cmd and keeps the existing <p:video> node', () => {
    const body = wrap(pptVideoTiming())
    const anims = readSlideTimingXml(body)
    const out = patchSlideTimingXml(body, [{ ...anims[0]!, mediaKind: 'video' }])
    expect(out).toContain('<p:cmd type="call" cmd="playFrom(0.0)">')
    expect(out).toContain('presetClass="mediacall"')
    expect(out.match(/<p:video>/g)).toHaveLength(1)
    expect(out).not.toContain('<p:bldP')
    expect(readSlideTimingXml(out)[0]).toMatchObject({ effect: 'mediaPlay', durationMs: 0 })
  })

  it('a fresh build adds the media timeline node for each media shape and no bldP for media', () => {
    const anims: SlideAnimation[] = [
      { spid: 3, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
      {
        spid: 4,
        effect: 'mediaPlay',
        trigger: 'onClick',
        durationMs: 0,
        delayMs: 0,
        mediaKind: 'video',
      },
      {
        spid: 4,
        effect: 'mediaPause',
        trigger: 'onClick',
        durationMs: 0,
        delayMs: 0,
        mediaKind: 'video',
      },
      {
        spid: 5,
        effect: 'mediaPlay',
        trigger: 'withPrev',
        durationMs: 0,
        delayMs: 0,
        mediaKind: 'audio',
      },
    ]
    const xml = buildTimingXml(anims)
    expect(xml.match(/<p:video>/g)).toHaveLength(1)
    expect(xml.match(/<p:audio>/g)).toHaveLength(1)
    expect(xml).toContain('cmd="togglePause"')
    expect(xml.match(/<p:bldP /g)).toHaveLength(1)
    expect(xml).toContain('<p:bldP spid="3"')
    const back = readSlideTimingXml(wrap(xml))
    expect(back.map((a) => a.effect)).toEqual(['fade', 'mediaPlay', 'mediaPause', 'mediaPlay'])
    expect(back[3]!.trigger).toBe('withPrev')
  })

  it('media steps do not consume a build group id: the visual effect after them keeps grpId 0', () => {
    const xml = buildTimingXml([
      {
        spid: 4,
        effect: 'mediaPlay',
        trigger: 'onClick',
        durationMs: 0,
        delayMs: 0,
        mediaKind: 'video',
      },
      { spid: 4, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
      {
        spid: 4,
        effect: 'mediaStop',
        trigger: 'onClick',
        durationMs: 0,
        delayMs: 0,
        mediaKind: 'video',
      },
      { spid: 4, effect: 'fadeOut', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ])
    const grpIds = [...xml.matchAll(/presetClass="(entr|exit)"[^>]*\bgrpId="(\d+)"/g)].map(
      (m) => m[2],
    )
    expect(grpIds).toEqual(['0', '1'])
    expect(xml).toContain('<p:bldP spid="4" grpId="0"/>')
    expect(xml).toContain('<p:bldP spid="4" grpId="1"/>')
    expect(xml.match(/<p:bldP /g)).toHaveLength(2)
    // Incremental append after a media step numbers the same way
    const base = wrap(xml)
    const next = patchSlideTimingIncrementalXml(base, [
      ...readSlideTimingXml(base),
      { spid: 4, effect: 'zoom', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ])
    expect(next).not.toBeNull()
    expect(next!).toContain('<p:bldP spid="4" grpId="2"/>')
    expect(/presetClass="entr" presetSubtype="16"[^>]*\bgrpId="2"/.test(next!)).toBe(true)
  })

  it('a media-only build has no bldLst', () => {
    const xml = buildTimingXml([
      {
        spid: 4,
        effect: 'mediaPlay',
        trigger: 'onClick',
        durationMs: 0,
        delayMs: 0,
        mediaKind: 'video',
      },
    ])
    expect(xml).not.toContain('<p:bldLst')
    expect(xml).toContain('<p:video>')
  })

  it('incremental append of a media effect adds its media node once', () => {
    const base = wrap(
      buildTimingXml([
        { spid: 3, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
      ]),
    )
    const old = readSlideTimingXml(base)
    const play: SlideAnimation = {
      spid: 4,
      effect: 'mediaPlay',
      trigger: 'onClick',
      durationMs: 0,
      delayMs: 0,
      mediaKind: 'video',
    }
    const step1 = patchSlideTimingIncrementalXml(base, [...old, play])
    expect(step1).not.toBeNull()
    expect(step1!.match(/<p:video>/g)).toHaveLength(1)
    const step2 = patchSlideTimingIncrementalXml(step1!, [
      ...readSlideTimingXml(step1!),
      { ...play, effect: 'mediaStop' },
    ])
    expect(step2).not.toBeNull()
    expect(step2!.match(/<p:video>/g)).toHaveLength(1)
    expect(step2!).toContain('cmd="stop"')
    expect(step2!.match(/<p:bldP /g)).toHaveLength(1)
  })
})
