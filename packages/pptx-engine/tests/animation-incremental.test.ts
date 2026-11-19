/**
 * Incremental animation write-back: appends and single edits at the tail use surgical patches,
 * keeping original bytes of unmodeled timing nodes; delete/reorder/trigger changes fall back to full rebuild.
 */
import { describe, it, expect } from 'vitest'
import {
  buildTimingXml,
  patchSlideTimingIncrementalXml,
  patchSlideTimingXml,
  readSlideTimingXml,
  type SlideAnimation,
} from '../src/animation'

const A1: SlideAnimation = {
  spid: 4,
  effect: 'fade',
  trigger: 'onClick',
  durationMs: 500,
  delayMs: 0,
}
const A2: SlideAnimation = {
  spid: 5,
  effect: 'flyIn',
  trigger: 'onClick',
  durationMs: 1000,
  delayMs: 0,
}

/** Injects an unmodeled interactive sequence into our generated timing (skeleton of a PowerPoint trigger animation). */
const EXOTIC =
  '<p:seq concurrent="1" nextAc="none"><p:cTn id="900" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">' +
  '<p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="99"/></p:tgtEl></p:cond></p:stCondLst>' +
  '<p:childTnLst/></p:cTn></p:seq>'

/** bodySuffix skeleton: clrMapOvr + timing. */
function bodySuffixWith(timing: string): string {
  return `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${timing}</p:sld>`
}

/** Inject an exotic node into timing (after mainSeq, inside tmRoot childTnLst). */
function withExotic(timing: string): string {
  const at = timing.indexOf('</p:seq>') + '</p:seq>'.length
  return timing.slice(0, at) + EXOTIC + timing.slice(at)
}

describe('patchSlideTimingIncrementalXml', () => {
  it('append animation: original effects and unmodeled nodes keep their bytes, new group attached at the end of mainSeq', () => {
    const body = bodySuffixWith(withExotic(buildTimingXml([A1])))
    const out = patchSlideTimingIncrementalXml(body, [A1, A2])
    expect(out).not.toBeNull()
    expect(out!).toContain(EXOTIC) // unmodeled node kept as-is
    const anims = readSlideTimingXml(out!)
    expect(anims.length).toBe(2)
    expect(anims[1]).toMatchObject({ spid: 5, effect: 'flyIn' })
    // Original A1 effect cTn bytes untouched (the par buildTimingXml generated separately is still there)
    const a1Par = /<p:cTn[^>]*presetClass="entr"[^>]*>/.exec(buildTimingXml([A1]))![0]
    expect(out!).toContain(a1Par)
    // bldP appended
    expect(out!).toContain('<p:bldP spid="5" grpId="0"/>')
  })

  it('appended cTn ids do not clash with existing ones', () => {
    const body = bodySuffixWith(withExotic(buildTimingXml([A1]))) // exotic contains id="900"
    const out = patchSlideTimingIncrementalXml(body, [A1, A2])!
    const ids = [...out.matchAll(/\bid="(\d+)"/g)].map((m) => Number(m[1]))
    expect(new Set(ids).size).toBe(ids.length)
    expect(Math.max(...ids)).toBeGreaterThan(900)
  })

  it('single tail edit: only that effect par replaced, siblings and unmodeled nodes preserved', () => {
    const body = bodySuffixWith(withExotic(buildTimingXml([A1, A2])))
    const edited = { ...A2, durationMs: 2000, delayMs: 250 }
    const out = patchSlideTimingIncrementalXml(body, [A1, edited])
    expect(out).not.toBeNull()
    expect(out!).toContain(EXOTIC)
    const anims = readSlideTimingXml(out!)
    expect(anims[1]).toMatchObject({ durationMs: 2000, delayMs: 250, effect: 'flyIn' })
    expect(anims[0]).toMatchObject({ spid: 4, effect: 'fade', durationMs: 500 })
  })

  it('delete/reorder/trigger change → returns null (falls back to rebuild)', () => {
    const body = bodySuffixWith(buildTimingXml([A1, A2]))
    expect(patchSlideTimingIncrementalXml(body, [A2])).toBeNull() // delete
    expect(patchSlideTimingIncrementalXml(body, [A2, A1])).toBeNull() // reorder
    expect(patchSlideTimingIncrementalXml(body, [A1, { ...A2, trigger: 'withPrev' }])).toBeNull() // trigger change
    // Append whose first entry is not onClick (would merge into the old group) → fall back
    expect(
      patchSlideTimingIncrementalXml(body, [A1, A2, { ...A1, trigger: 'withPrev' }]),
    ).toBeNull()
  })

  it('round-trip: appending after an append still takes the incremental path', () => {
    let body = bodySuffixWith(withExotic(buildTimingXml([A1])))
    body = patchSlideTimingIncrementalXml(body, [A1, A2])!
    const A3: SlideAnimation = {
      spid: 4,
      effect: 'zoomOut',
      trigger: 'onClick',
      durationMs: 300,
      delayMs: 100,
    }
    const out = patchSlideTimingIncrementalXml(body, [A1, A2, A3])
    expect(out).not.toBeNull()
    expect(out!).toContain(EXOTIC)
    const anims = readSlideTimingXml(out!)
    expect(anims.length).toBe(3)
    expect(anims[2]).toMatchObject({ spid: 4, effect: 'zoomOut', delayMs: 100 })
    // Second animation on the same spid continues with grpId 1
    expect(out!).toContain('<p:bldP spid="4" grpId="1"/>')
  })
})

/** Interactive sequence with a real trigger effect (presetClass cTn inside interactiveSeq). */
const EXOTIC_WITH_EFFECT =
  '<p:seq concurrent="1" nextAc="none"><p:cTn id="900" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">' +
  '<p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="99"/></p:tgtEl></p:cond></p:stCondLst>' +
  '<p:childTnLst><p:par><p:cTn id="901" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">' +
  '<p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
  '<p:childTnLst><p:set><p:cBhvr><p:cTn id="902" dur="1"/><p:tgtEl><p:spTgt spid="88"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst>' +
  '</p:cTn></p:par></p:childTnLst></p:cTn></p:seq>'

describe('trigger animations are not mistakenly absorbed', () => {
  it('readSlideTimingXml reads only the main sequence; effects inside interactiveSeq do not appear in the list', () => {
    const body = bodySuffixWith(
      withExotic(buildTimingXml([A1])).replace(EXOTIC, EXOTIC_WITH_EFFECT),
    )
    const anims = readSlideTimingXml(body)
    expect(anims.length).toBe(1)
    expect(anims[0]).toMatchObject({ spid: 4, effect: 'fade' })
  })

  it('no main sequence (trigger-only timing) -> reads back an empty list', () => {
    const timing = `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${EXOTIC_WITH_EFFECT}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`
    expect(readSlideTimingXml(bodySuffixWith(timing))).toEqual([])
  })
})

describe('patchSlideTimingXml full rebuild preserves unmodeled nodes', () => {
  it('deleting an animation (reorder/removal falls back to rebuild): interactiveSeq keeps original bytes', () => {
    const body = bodySuffixWith(withExotic(buildTimingXml([A1, A2])))
    const out = patchSlideTimingXml(body, [A2]) // remove A1
    expect(out).toContain(EXOTIC)
    const anims = readSlideTimingXml(out)
    expect(anims.length).toBe(1)
    expect(anims[0]).toMatchObject({ spid: 5, effect: 'flyIn' })
  })

  it('reordering animations: trigger effects are kept and not absorbed', () => {
    const body = bodySuffixWith(
      withExotic(buildTimingXml([A1, A2])).replace(EXOTIC, EXOTIC_WITH_EFFECT),
    )
    const out = patchSlideTimingXml(body, [A2, A1])
    expect(out).toContain(EXOTIC_WITH_EFFECT)
    const anims = readSlideTimingXml(out)
    expect(anims.map((a) => a.spid)).toEqual([5, 4])
  })

  it('rebuilt cTn ids do not clash with preserved nodes', () => {
    const body = bodySuffixWith(withExotic(buildTimingXml([A1, A2]))) // exotic id=900
    const out = patchSlideTimingXml(body, [A2, A1])
    const ids = [...out.matchAll(/\bid="(\d+)"/g)].map((m) => Number(m[1]))
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('clearing animations: main sequence removed, interactiveSeq kept, timing stays', () => {
    const body = bodySuffixWith(withExotic(buildTimingXml([A1])))
    const out = patchSlideTimingXml(body, [])
    expect(out).toContain('<p:timing>')
    expect(out).toContain(EXOTIC)
    expect(out).not.toContain('nodeType="mainSeq"')
    expect(readSlideTimingXml(out)).toEqual([])
  })

  it('clearing animations with no unmodeled nodes: the whole timing is removed', () => {
    const body = bodySuffixWith(buildTimingXml([A1]))
    const out = patchSlideTimingXml(body, [])
    expect(out).not.toContain('<p:timing>')
  })

  it('bldLst: bldP rewritten, unmodeled builds like bldDgm kept', () => {
    let timing = withExotic(buildTimingXml([A1, A2]))
    timing = timing.replace('</p:bldLst>', '<p:bldDgm spid="77" grpId="0"/></p:bldLst>')
    const out = patchSlideTimingXml(bodySuffixWith(timing), [A2])
    expect(out).toContain('<p:bldDgm spid="77" grpId="0"/>')
    expect(out).toContain('<p:bldP spid="5" grpId="0"/>')
    expect(out).not.toContain('<p:bldP spid="4"')
  })
})
