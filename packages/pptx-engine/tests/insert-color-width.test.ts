import { describe, expect, it } from 'vitest'
import { addElement, buildSpXml } from '../src/insert'
import type { Slide } from '../src/types'

const slide = { originalXml: '<p:spTree/>', elements: [] } as unknown as Slide

const base = { kind: 'textbox', offset: { x: 0, y: 0, cx: 914400, cy: 914400 } }

describe('inserted shape colors and stroke widths', () => {
  it('normalizes 3-digit and invalid fill colors to ST_HexColorRGB', () => {
    const xml = buildSpXml(slide, { ...base, fillColor: '#F0A' })
    expect(xml).toContain('<a:srgbClr val="FF00AA"/>')
    const bad = buildSpXml(slide, { ...base, fillColor: 'not-a-color' })
    // ST_HexColorRGB needs 6 hex digits — anything else falls back to black
    expect(bad).toContain('val="000000"')
    expect(bad).not.toContain('val="NOT"')
  })

  it('normalizes stroke colors the same way', () => {
    const xml = buildSpXml(slide, { ...base, stroke: { color: '#0f0', widthEmu: 12700 } })
    expect(xml).toContain('<a:srgbClr val="00FF00"/>')
  })

  it('falls back and clamps non-finite stroke widths into ST_LineWidth', () => {
    const inf = buildSpXml(slide, { ...base, stroke: { color: '#000000', widthEmu: Infinity } })
    expect(inf).toContain('<a:ln w="12700">')
    const huge = buildSpXml(slide, { ...base, stroke: { color: '#000000', widthEmu: 1e12 } })
    expect(huge).toContain('<a:ln w="20116800">')
    const nan = buildSpXml(slide, { ...base, stroke: { color: '#000000', widthEmu: NaN } })
    expect(nan).toContain('<a:ln w="12700">')
  })

  it('applies the same clamps to connector lines', () => {
    const el = addElement(slide, {
      ...base,
      kind: 'line',
      stroke: { color: '#abc', widthEmu: Infinity },
    })
    expect(el.anchor.originalXml).toContain('<a:ln w="12700" cap="flat">')
    expect(el.anchor.originalXml).toContain('val="AABBCC"')
  })
})
