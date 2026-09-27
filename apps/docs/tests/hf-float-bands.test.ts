import { describe, expect, it } from 'vitest'
import type { HfImage, HfTextBox } from '@genoffice/docx-engine'
import { hfFloatPagePos, hfTextBoxStyle, type HfStripGeom } from '../src/renderer/editor/hf-dom'

/** Letter page, 1in margins, header/footer strips 0.5in from the edges (px) */
const box = {
  pageW: 816,
  pageH: 1056,
  marginLeft: 96,
  marginRight: 96,
  marginTop: 96,
  marginBottom: 96,
  headerDist: 48,
  sectMarginTop: 96,
}

const geom: HfStripGeom = { ...box, headerStripTop: 48, footerDist: 48 }

describe('anchored objects relative to the margin bands', () => {
  it('centres a bottomMargin anchor inside the bottom margin, not the page', () => {
    const img: HfImage = {
      dataUrl: '',
      posH: 'center',
      posHRel: 'margin',
      posV: 'center',
      posVRel: 'bottomMargin',
    }
    const p = hfFloatPagePos(img, box)
    expect(p.x).toBe(408)
    expect(p.y).toBe(1008)
    expect(p.translateY).toBe(-50)
  })

  it('measures the bottomMargin band from the raw sectPr margin, not the footer-pushed one', () => {
    const img: HfImage = { dataUrl: '', posV: 'center', posVRel: 'bottomMargin' }
    const p = hfFloatPagePos(img, { ...box, marginBottom: 200, sectMarginBottom: 96 })
    expect(p.y).toBe(1008)
  })

  it('measures topMargin / leftMargin / rightMargin offsets from the band edges', () => {
    const top = hfFloatPagePos(
      { dataUrl: '', posXPx: 10, posHRel: 'leftMargin', posYPx: 20, posVRel: 'topMargin' },
      box,
    )
    expect([top.x, top.y]).toEqual([10, 20])
    const right = hfFloatPagePos(
      {
        dataUrl: '',
        posXPx: 10,
        posHRel: 'rightMargin',
        posH: 'left',
        posV: 'bottom',
        posVRel: 'topMargin',
      },
      box,
    )
    expect([right.x, right.y, right.translateY]).toEqual([730, 96, -100])
  })

  it('keeps the legacy margin/page placement', () => {
    const margin = hfFloatPagePos(
      { dataUrl: '', posXPx: 10, posHRel: 'margin', posV: 'bottom', posVRel: 'margin' },
      box,
    )
    expect([margin.x, margin.y]).toEqual([106, 960])
    const page = hfFloatPagePos(
      { dataUrl: '', posH: 'center', posV: 'center', posHRel: 'page', posVRel: 'page' },
      box,
    )
    expect([page.x, page.y]).toEqual([408, 528])
  })

  it('pins a bottomMargin-centred footer box by its bottom edge inside the margin band', () => {
    const tb: HfTextBox = {
      id: 1,
      widthPx: 66,
      heightPx: 66,
      posH: 'center',
      posHRel: 'margin',
      posV: 'center',
      posVRel: 'bottomMargin',
    }
    const style = hfTextBoxStyle(tb, 'footer', geom)!
    // centre 1008, half height 33: box spans 975..1041; bottom edge is 15 px above the footer strip bottom (1008 = pageH - footerDist)
    expect(style.left).toBe(`${408 - 33 - 96}px`)
    expect(style.bottom).toBe(`${1056 - 48 - 1041}px`)
  })
})
