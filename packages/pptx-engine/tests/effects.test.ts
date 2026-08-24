/** setElementEffects: <a:effectLst> shadow/glow/softEdge byte surgery. */
import { describe, it, expect } from 'vitest'
import { addElement, createBlankPptx, openPptx, setElementEffects } from '../src/index'
import type { ShadowEffect, GlowEffect, TextElement } from '../src/types'

type EffectsElement = TextElement & { shadow?: ShadowEffect; glow?: GlowEffect; softEdge?: number }

async function shapeSlide() {
  const opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  const el = addElement(slide, {
    kind: 'rect',
    offset: { x: 0, y: 0, cx: 1000, cy: 1000 },
  })
  return { slide, el: el as EffectsElement }
}

describe('setElementEffects', () => {
  it('writes an outer shadow with alpha color and mirrors the model', async () => {
    const { slide, el } = await shapeSlide()
    expect(
      setElementEffects(slide, el.id, {
        shadow: { color: '#00000066', blurRad: 50800, dist: 38100, dirDeg: 45 },
      }),
    ).toBe(true)
    const xml = el.anchor.originalXml
    expect(xml).toMatch(
      /<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"\/><\/a:srgbClr><\/a:outerShdw><\/a:effectLst>/,
    )
    expect(el.shadow).toEqual({ color: '#00000066', blurRad: 50800, dist: 38100, dirDeg: 45 })
  })

  it('glow and softEdge coexist with the shadow in schema order', async () => {
    const { slide, el } = await shapeSlide()
    setElementEffects(slide, el.id, {
      shadow: { color: '#000000', blurRad: 0, dist: 0, dirDeg: 0 },
    })
    setElementEffects(slide, el.id, { glow: { color: '#FF0000', radius: 63500 } })
    setElementEffects(slide, el.id, { softEdge: 127000 })
    const xml = el.anchor.originalXml
    const lst = /<a:effectLst>([\s\S]*?)<\/a:effectLst>/.exec(xml)?.[1] ?? ''
    // schema order: glow < outerShdw < softEdge
    expect(lst.indexOf('<a:glow')).toBeGreaterThanOrEqual(0)
    expect(lst.indexOf('<a:glow')).toBeLessThan(lst.indexOf('<a:outerShdw'))
    expect(lst.indexOf('<a:outerShdw')).toBeLessThan(lst.indexOf('<a:softEdge'))
    expect(el.glow).toEqual({ color: '#FF0000', radius: 63500 })
    expect((el as { softEdge?: number }).softEdge).toBe(127000)
  })

  it('null clears one effect and removes an emptied effectLst entirely', async () => {
    const { slide, el } = await shapeSlide()
    setElementEffects(slide, el.id, {
      shadow: { color: '#000000', blurRad: 1000, dist: 1000, dirDeg: 90 },
      glow: { color: '#00FF00', radius: 1000 },
    })
    expect(setElementEffects(slide, el.id, { glow: null })).toBe(true)
    expect(el.anchor.originalXml).not.toMatch(/<a:glow/)
    expect(el.anchor.originalXml).toMatch(/<a:outerShdw/)
    expect(el.glow).toBeUndefined()
    expect(setElementEffects(slide, el.id, { shadow: null })).toBe(true)
    expect(el.anchor.originalXml).not.toMatch(/<a:effectLst/)
    expect(el.shadow).toBeUndefined()
  })

  it('inner: true writes <a:innerShdw> and switching kinds replaces the other', async () => {
    const { slide, el } = await shapeSlide()
    setElementEffects(slide, el.id, {
      shadow: { color: '#00000080', blurRad: 63500, dist: 50800, dirDeg: 45, inner: true },
    })
    expect(el.anchor.originalXml).toMatch(
      /<a:innerShdw blurRad="63500" dist="50800" dir="2700000">/,
    )
    expect(el.anchor.originalXml).not.toMatch(/outerShdw|rotWithShape/)
    expect(el.shadow).toMatchObject({ inner: true, dirDeg: 45 })
    setElementEffects(slide, el.id, {
      shadow: { color: '#000000', blurRad: 0, dist: 0, dirDeg: 0 },
    })
    expect(el.anchor.originalXml).toMatch(/<a:outerShdw/)
    expect(el.anchor.originalXml).not.toMatch(/innerShdw/)
    expect(el.shadow?.inner).toBeUndefined()
  })

  it('perspective attributes serialize in schema order with scale/skew units', async () => {
    const { slide, el } = await shapeSlide()
    setElementEffects(slide, el.id, {
      shadow: {
        color: '#0000003b',
        blurRad: 152400,
        dist: 317500,
        dirDeg: 90,
        sx: 0.9,
        sy: -0.19,
        kxDeg: -20,
        algn: 'bl',
      },
    })
    expect(el.anchor.originalXml).toMatch(
      /<a:outerShdw blurRad="152400" dist="317500" dir="5400000" sx="90000" sy="-19000" kx="-1200000" algn="bl" rotWithShape="0">/,
    )
    expect(el.shadow).toMatchObject({ sx: 0.9, sy: -0.19, kxDeg: -20, algn: 'bl' })
  })

  it('a full-opacity color writes no alpha child; dir normalizes into 0..360°', async () => {
    const { slide, el } = await shapeSlide()
    setElementEffects(slide, el.id, {
      shadow: { color: '#123456', blurRad: 0, dist: 0, dirDeg: -90 },
    })
    const xml = el.anchor.originalXml
    expect(xml).toMatch(/<a:srgbClr val="123456"\/>/)
    expect(xml).not.toMatch(/<a:alpha/)
    expect(xml).toMatch(/dir="16200000"/) // -90° → 270°
  })
})

describe('setElementEffects reflection', () => {
  it('writes <a:reflection> with flip/fade attrs, mirrors the model, and clears', async () => {
    const { slide, el } = await shapeSlide()
    setElementEffects(slide, el.id, {
      reflection: { blurRad: 6350, startA: 0.52, endPos: 0.35, dist: 50800 },
    })
    expect(el.anchor.originalXml).toMatch(
      /<a:reflection blurRad="6350" stA="52000" endA="300" endPos="35000" dist="50800" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"\/>/,
    )
    expect((el as { reflection?: object }).reflection).toEqual({
      blurRad: 6350,
      startA: 0.52,
      endPos: 0.35,
      dist: 50800,
    })
    // Coexists with a shadow in schema order (glow < innerShdw/outerShdw < reflection < softEdge)
    setElementEffects(slide, el.id, {
      shadow: { color: '#000000', blurRad: 0, dist: 0, dirDeg: 0 },
    })
    const lst = /<a:effectLst>([\s\S]*?)<\/a:effectLst>/.exec(el.anchor.originalXml)?.[1] ?? ''
    expect(lst.indexOf('<a:outerShdw')).toBeLessThan(lst.indexOf('<a:reflection'))
    setElementEffects(slide, el.id, { reflection: null })
    expect(el.anchor.originalXml).not.toMatch(/<a:reflection/)
    expect((el as { reflection?: object }).reflection).toBeUndefined()
  })
})
