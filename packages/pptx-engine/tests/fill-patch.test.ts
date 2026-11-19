import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { openPptx, savePptx, addElement, patchElementFill } from '../src/index'
import type { TextElement } from '../src/types'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

describe('patchElementFill', () => {
  it('replaces an existing solidFill without touching the outline fill', () => {
    const xml =
      '<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="112233"/></a:solidFill>' +
      '<a:ln><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:ln>' +
      '</p:spPr></p:sp>'
    const out = patchElementFill(xml, '#C43E1C')
    expect(out).toContain('<a:solidFill><a:srgbClr val="C43E1C"/></a:solidFill><a:ln>')
    expect(out).toContain('<a:ln><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:ln>')
    expect(out).not.toContain('112233')
  })

  it('injects fill after geometry when the sp has none', () => {
    const xml = '<p:sp><p:spPr><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr></p:sp>'
    const out = patchElementFill(xml, '#00FF00')
    expect(out).toContain('</a:prstGeom><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>')
  })

  it("'none' writes <a:noFill/> replacing gradient fill", () => {
    const xml =
      '<p:sp><p:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst></a:gradFill></p:spPr></p:sp>'
    const out = patchElementFill(xml, 'none')
    expect(out).toContain('<a:noFill/>')
    expect(out).not.toContain('gradFill')
  })

  it('dirtyFill round-trips through save → reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'rect',
      offset: { x: 914400, y: 914400, cx: 914400, cy: 914400 },
      fillColor: '#112233',
    })
    // Simulate the edit-fill handler
    el.fill = { type: 'solid', color: '#C43E1C' }
    el.dirtyFill = true

    const reopened = await openPptx(await savePptx(opened))
    const el2 = reopened.deck.slides[0]!.elements.at(-1) as TextElement
    expect(el2.fill).toEqual({ type: 'solid', color: '#C43E1C' })
  })
})

describe('gradient fill write-back', () => {
  it('linear gradient gradFill + lin ang', async () => {
    const { patchElementFill } = await import('../src/generate')
    const xml =
      '<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>'
    const out = patchElementFill(xml, {
      stops: [
        { pos: 0, color: '#112233' },
        { pos: 1, color: '#445566' },
      ],
      angle: 5400000,
    })
    expect(out).toContain('<a:gradFill rotWithShape="1">')
    expect(out).toContain('<a:gs pos="0"><a:srgbClr val="112233"/></a:gs>')
    expect(out).toContain('<a:gs pos="100000"><a:srgbClr val="445566"/></a:gs>')
    expect(out).toContain('<a:lin ang="5400000" scaled="1"/>')
    expect(out).not.toContain('FF0000')
  })

  it('radial gradient writes path circle', async () => {
    const { patchElementFill } = await import('../src/generate')
    const xml = '<p:sp><p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>'
    const out = patchElementFill(xml, {
      stops: [
        { pos: 0, color: '#000000' },
        { pos: 1, color: '#ffffff' },
      ],
      radial: true,
    })
    expect(out).toContain('<a:path path="circle">')
    expect(out).not.toContain('<a:lin')
  })
})

describe('same-kind fill in-place patch (unmodeled parameters preserved)', () => {
  it('solid->solid keeps alpha', () => {
    const xml =
      '<p:sp><p:spPr><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="60000"/><a:alpha val="50000"/></a:schemeClr></a:solidFill></p:spPr></p:sp>'
    const out = patchElementFill(xml, '#C43E1C')
    expect(out).toContain(
      '<a:solidFill><a:srgbClr val="C43E1C"><a:alpha val="50000"/></a:srgbClr></a:solidFill>',
    )
  })

  it('an 8-digit color sets an explicit alpha', () => {
    const xml = '<p:sp><p:spPr><a:solidFill><a:srgbClr val="C43E1C"/></a:solidFill></p:spPr></p:sp>'
    // 0x33 = 51/255 → 20% opacity
    const out = patchElementFill(xml, '#C0000033')
    expect(out).toContain(
      '<a:solidFill><a:srgbClr val="C00000"><a:alpha val="20000"/></a:srgbClr></a:solidFill>',
    )
  })

  it('an explicit FF alpha removes the existing alpha', () => {
    const xml =
      '<p:sp><p:spPr><a:solidFill><a:srgbClr val="C43E1C"><a:alpha val="50000"/></a:srgbClr></a:solidFill></p:spPr></p:sp>'
    const out = patchElementFill(xml, '#C43E1CFF')
    expect(out).toContain('<a:solidFill><a:srgbClr val="C43E1C"/></a:solidFill>')
    expect(out).not.toContain('a:alpha')
  })

  it('inserting a fresh fill honors the alpha', () => {
    const xml = '<p:sp><p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>'
    const out = patchElementFill(xml, '#C0000080')
    expect(out).toContain(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="C00000"><a:alpha val="50196"/></a:srgbClr></a:solidFill>',
    )
  })

  it('grad->grad keeps gradFill attributes / tileRect / off-center path, replaces gsLst', () => {
    const xml =
      '<p:sp><p:spPr><a:gradFill flip="xy" rotWithShape="0">' +
      '<a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst>' +
      '<a:path path="rect"><a:fillToRect l="20000" t="30000" r="80000" b="70000"/></a:path>' +
      '<a:tileRect l="-10000"/>' +
      '</a:gradFill></p:spPr></p:sp>'
    const out = patchElementFill(xml, {
      stops: [
        { pos: 0, color: '#111111' },
        { pos: 1, color: '#222222' },
      ],
      radial: true,
    })
    expect(out).toContain('<a:gradFill flip="xy" rotWithShape="0">')
    expect(out).toContain(
      '<a:gs pos="0"><a:srgbClr val="111111"/></a:gs><a:gs pos="100000"><a:srgbClr val="222222"/></a:gs>',
    )
    // The original off-center path keeps its bytes, not overwritten by the default centered circle
    expect(out).toContain(
      '<a:path path="rect"><a:fillToRect l="20000" t="30000" r="80000" b="70000"/></a:path>',
    )
    expect(out).toContain('<a:tileRect l="-10000"/>')
    expect(out).not.toContain('FF0000')
  })

  it('grad->grad linear angle change keeps scaled, radial path replaced by a:lin', () => {
    const xml =
      '<p:sp><p:spPr><a:gradFill rotWithShape="1">' +
      '<a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst>' +
      '<a:lin ang="2700000" scaled="0"/>' +
      '</a:gradFill></p:spPr></p:sp>'
    const out = patchElementFill(xml, {
      stops: [
        { pos: 0, color: '#111111' },
        { pos: 1, color: '#222222' },
      ],
      angle: 5400000,
    })
    expect(out).toContain('<a:lin ang="5400000" scaled="0"/>')
    expect(out).not.toContain('<a:path')
  })

  it('lin->radial with no existing path adds a default circle path right after gsLst', () => {
    const xml =
      '<p:sp><p:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/><a:tileRect/></a:gradFill></p:spPr></p:sp>'
    const out = patchElementFill(xml, {
      stops: [
        { pos: 0, color: '#111111' },
        { pos: 1, color: '#222222' },
      ],
      radial: true,
    })
    expect(out).toContain('</a:gsLst><a:path path="circle">')
    expect(out).toContain('</a:path><a:tileRect/>')
    expect(out).not.toContain('<a:lin')
  })

  it('cross-type (grad->solid) replaces the whole block', () => {
    const xml =
      '<p:sp><p:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill></p:spPr></p:sp>'
    const out = patchElementFill(xml, '#123456')
    expect(out).toContain('<a:solidFill><a:srgbClr val="123456"/></a:solidFill>')
    expect(out).not.toContain('gradFill')
  })
})
