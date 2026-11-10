import { describe, expect, it } from 'vitest'
import { pruneTimingForSpids, type Slide } from '../src/index'

const slideOf = (bodySuffix: string): Slide => ({ bodySuffix, structureDirty: false }) as never

const INNERMOST_PAR = /<p:par\b[^>]*>(?:(?!<p:par\b|<\/p:par>)[\s\S])*?<\/p:par>/g

const countPars = (xml: string): number => [...xml.matchAll(/<p:par\b/g)].length

const countEmptyPars = (xml: string): number =>
  [...xml.matchAll(INNERMOST_PAR)].filter((m) => !m[0].includes('<p:spTgt')).length

const setFor = (id: number, spid: number): string =>
  `<p:set><p:cBhvr><p:cTn id="${id}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>` +
  `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>` +
  '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>' +
  '</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'

const effectPar = (id: number, spid: number): string =>
  `<p:par><p:cTn id="${id}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">` +
  '<p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
  `<p:childTnLst>${setFor(id + 1, spid)}</p:childTnLst></p:cTn></p:par>`

const wrapChain = (levels: number, idBase: number, inner: string): string => {
  let out = inner
  for (let i = levels - 1; i >= 0; i--) {
    out = `<p:par><p:cTn id="${idBase + i}"><p:childTnLst>${out}</p:childTnLst></p:cTn></p:par>`
  }
  return out
}

const timingOf = (body: string): string =>
  '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
  body +
  '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'

describe('bounded timing cleanup (deep par nesting, many targets)', () => {
  it('collapses a realistic nesting depth completely', () => {
    const body = wrapChain(12, 100, effectPar(900, 5)) + effectPar(800, 7)
    const slide = slideOf(`${timingOf(body)}</p:sld>`)

    expect(pruneTimingForSpids(slide, new Set([5]))).toBe(true)
    expect(slide.bodySuffix).toContain('<p:spTgt spid="7"')
    expect(slide.bodySuffix).not.toContain('<p:spTgt spid="5"')
    expect(countEmptyPars(slide.bodySuffix)).toBe(0)
    expect(countPars(slide.bodySuffix)).toBe(2)
  })

  it('stays bounded past the collapse-pass cap and keeps surviving targets', () => {
    const body = wrapChain(200, 1000, effectPar(5000, 5)) + effectPar(800, 7)
    const slide = slideOf(`${timingOf(body)}</p:sld>`)

    expect(pruneTimingForSpids(slide, new Set([5]))).toBe(true)
    expect(slide.bodySuffix).toContain('<p:spTgt spid="7"')
    expect(slide.bodySuffix).not.toContain('<p:spTgt spid="5"')
    expect(countEmptyPars(slide.bodySuffix)).toBeLessThanOrEqual(64)
  })

  it('drops a deep timing whose every target is removed', () => {
    const slide = slideOf(`${timingOf(wrapChain(200, 1000, effectPar(5000, 5)))}</p:sld>`)

    expect(pruneTimingForSpids(slide, new Set([5]))).toBe(true)
    expect(slide.bodySuffix).not.toContain('<p:timing>')
  })

  it('prunes every one of many targets in a single call', () => {
    const N = 300
    let body = effectPar(90_000, 99_999)
    for (let i = 0; i < N; i++) body += effectPar(10_000 + i * 2, 1000 + i)
    const slide = slideOf(`${timingOf(body)}</p:sld>`)

    const spids = new Set<number>()
    for (let i = 0; i < N; i++) spids.add(1000 + i)

    expect(pruneTimingForSpids(slide, spids)).toBe(true)
    expect(slide.bodySuffix).toContain('<p:spTgt spid="99999"')
    for (let i = 0; i < N; i++)
      expect(slide.bodySuffix).not.toContain(`<p:spTgt spid="${1000 + i}"`)
  })

  it('leaves a shared wrapper alone while pruning many other targets', () => {
    const shared = `<p:par><p:cTn id="7000"><p:childTnLst>${setFor(7001, 5)}${setFor(7002, 6)}</p:childTnLst></p:cTn></p:par>`
    let body = shared
    for (let i = 0; i < 50; i++) body += effectPar(20_000 + i * 2, 1000 + i)
    const slide = slideOf(`${timingOf(body)}</p:sld>`)

    const spids = new Set<number>()
    for (let i = 0; i < 50; i++) spids.add(1000 + i)
    spids.add(5)

    expect(pruneTimingForSpids(slide, spids)).toBe(true)
    expect(slide.bodySuffix).toContain('<p:spTgt spid="6"')
    expect(slide.bodySuffix).toContain('<p:spTgt spid="5"')
  })
})
