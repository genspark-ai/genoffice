import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { resolveTableStyle } from '../src/table-style'
import type { Theme } from '../src/theme'

const theme = {
  colors: {
    dk1: '#000000',
    lt1: '#FFFFFF',
    accent1: '#4472C4',
    accent2: '#ED7D31',
  },
} as unknown as Theme

const MEDIUM2_A1 = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'

function tableSlideXml(tblPrAttrs: string, styleId: string): string {
  const tc = (t: string) =>
    `<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>${t}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`
  return (
    '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>' +
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="T"/></p:nvGraphicFramePr>' +
    '<p:xfrm><a:off x="0" y="0"/><a:ext cx="3657600" cy="2743200"/></p:xfrm>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>' +
    `<a:tblPr ${tblPrAttrs}><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr>` +
    '<a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>' +
    `<a:tr h="914400">${tc('H1')}${tc('H2')}</a:tr>` +
    `<a:tr h="914400">${tc('A1')}${tc('A2')}</a:tr>` +
    `<a:tr h="914400">${tc('B1')}${tc('B2')}</a:tr>` +
    '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>' +
    '</p:spTree></p:cSld></p:sld>'
  )
}

describe('built-in table style Medium Style 2 (PowerPoint default)', () => {
  const slide = parseSlide({
    path: 'ppt/slides/slide1.xml',
    slideXml: tableSlideXml('firstRow="1" bandRow="1"', MEDIUM2_A1),
    ctx: { theme },
  })
  const tbl = slide.elements[0] as any

  it('header row: solid accent1 fill + bold white text', () => {
    const h = tbl.rows[0][0]
    expect(h.fill).toEqual({ type: 'solid', color: '#4472C4' })
    const run = h.text.paragraphs[0].runs[0]
    expect(run.bold).toBe(true)
    expect(run.color).toBe('#FFFFFF')
  })

  it('banded rows: odd data rows 40% tint, even 20% tint (wholeTbl)', () => {
    // accent1 #4472C4: 40% tint = 255*0.6 + c*0.4
    expect(tbl.rows[1][0].fill.color).toBe('#B4C7E7')
    expect(tbl.rows[2][0].fill.color).toBe('#DAE3F3')
  })

  it('inner separator lines + outer border: lt1 white lines (last-row bottom edge comes from the outer border)', () => {
    expect(tbl.rows[0][0].borders.b.fill.color).toBe('#FFFFFF')
    expect(tbl.rows[0][0].borders.r.fill.color).toBe('#FFFFFF')
    expect(tbl.rows[2][0].borders.b.fill.color).toBe('#FFFFFF')
    expect(tbl.rows[0][1].borders.r.fill.color).toBe('#FFFFFF')
    expect(tbl.rows[0][0].borders.t.fill.color).toBe('#FFFFFF')
    expect(tbl.rows[0][0].borders.l.fill.color).toBe('#FFFFFF')
  })

  it('data row text not bold, color dk1', () => {
    const run = tbl.rows[1][0].text.paragraphs[0].runs[0]
    expect(run.bold).toBe(false)
    expect(run.color).toBe('#000000')
  })
})

describe('tableStyles.xml custom styles', () => {
  const xml =
    '<?xml version="1.0"?><a:tblStyleLst xmlns:a="a" def="{X}">' +
    '<a:tblStyle styleId="{X}" styleName="Custom">' +
    '<a:wholeTbl><a:tcTxStyle><a:schemeClr val="dk1"/></a:tcTxStyle>' +
    '<a:tcStyle><a:fill><a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>' +
    '<a:firstRow><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle>' +
    '<a:tcStyle><a:fill><a:solidFill><a:srgbClr val="AA0000"/></a:solidFill></a:fill></a:tcStyle></a:firstRow>' +
    '</a:tblStyle></a:tblStyleLst>'

  it('parses wholeTbl/firstRow fill and text styles', () => {
    const def = resolveTableStyle('{X}', xml, theme)!
    expect(def.wholeTbl?.fill).toEqual({ type: 'solid', color: '#EEEEEE' })
    expect(def.wholeTbl?.textColor).toBe('#000000')
    expect(def.firstRow?.fill).toEqual({ type: 'solid', color: '#AA0000' })
    expect(def.firstRow?.bold).toBe(true)
    expect(def.firstRow?.textColor).toBe('#FFFFFF')
  })

  it('unknown styleId that is not built-in → undefined (stays unstyled)', () => {
    expect(resolveTableStyle('{DEAD-BEEF}', xml, theme)).toBeUndefined()
    expect(resolveTableStyle(undefined, xml, theme)).toBeUndefined()
  })

  it('tblBg fillRef + alpha band fills + lnRef borders (bnc480256)', () => {
    const themed = {
      ...theme,
      fillStyles: [
        {},
        {},
        { 'a:gradFill': {} }, // placeholder; parse.ts instantiates the ref, table-style only records it
      ],
      lnStyles: [
        { 'a:ln': { '@_w': '25400', 'a:solidFill': { 'a:schemeClr': { '@_val': 'phClr' } } } },
      ],
    } as unknown as Theme
    const bgXml =
      '<?xml version="1.0"?><a:tblStyleLst xmlns:a="a" def="{Y}">' +
      '<a:tblStyle styleId="{Y}" styleName="Bg">' +
      '<a:tblBg><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef></a:tblBg>' +
      '<a:wholeTbl><a:tcStyle><a:tcBdr>' +
      '<a:insideH><a:lnRef idx="1"><a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr></a:lnRef></a:insideH>' +
      '<a:insideV><a:ln><a:noFill/></a:ln></a:insideV>' +
      '</a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl>' +
      '<a:band1H><a:tcStyle><a:tcBdr/><a:fill><a:solidFill>' +
      '<a:schemeClr val="lt1"><a:alpha val="20000"/></a:schemeClr>' +
      '</a:solidFill></a:fill></a:tcStyle></a:band1H>' +
      '</a:tblStyle></a:tblStyleLst>'
    const def = resolveTableStyle('{Y}', bgXml, themed)!
    expect(def.tblBgRef).toEqual({ idx: 3, phClr: '#4472C4' })
    // alpha 20% → #RRGGBBAA suffix (composites over tblBg at render time)
    expect(def.band1H?.fill).toEqual({ type: 'solid', color: '#FFFFFF33' })
    // lnRef: phClr from the ref child (tint 50% of accent1), width from the theme lnStyle template
    expect(def.insideH).toEqual({ fill: { type: 'solid', color: '#A2B9E2' }, width: 25400 })
    // explicit <a:ln><a:noFill/> → no inside-vertical line
    expect(def.insideV).toBeUndefined()
  })
})

describe('built-in style family expansion (full GUID table)', () => {
  it('Light Style 1 - Accent1: no whole-table fill, 20% pre-blended bands, bold first row + accent separator line', () => {
    const def = resolveTableStyle('{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}', undefined, theme)!
    expect(def.wholeTbl?.fill).toBeUndefined()
    expect(def.firstRow?.bold).toBe(true)
    // accent1 20% alpha over white = #DAE3F3
    expect(def.band1H?.fill).toEqual({ type: 'solid', color: '#DAE3F3' })
    expect(def.firstRowBottom?.fill).toEqual({ type: 'solid', color: '#4472C4' })
    expect(def.outer?.t?.fill).toEqual({ type: 'solid', color: '#4472C4' })
    expect(def.outer?.l).toBeUndefined() // Light1 only has top/bottom outer borders
  })

  it('Dark Style 1 - Accent1: shade dark background + light text; no-accent variant tint light background dark text', () => {
    const dark = resolveTableStyle('{125E5076-3810-47DD-B79F-674D7AD40C01}', undefined, theme)!
    // accent1 shade 20% = #0E1727
    expect(dark.wholeTbl?.fill).toEqual({ type: 'solid', color: '#0E1727' })
    expect(dark.wholeTbl?.textColor).toBe('#FFFFFF')
    expect(dark.firstRow?.fill).toEqual({ type: 'solid', color: '#000000' })
    const base = resolveTableStyle('{E8034E78-7F5D-4C2E-B375-FC64B27BC917}', undefined, theme)!
    expect(base.wholeTbl?.fill).toEqual({ type: 'solid', color: '#CCCCCC' }) // dk1 tint 20%
    expect(base.wholeTbl?.textColor).toBe('#000000')
  })

  it('Dark Style 2 - Accent1: header uses the paired color accent2', () => {
    const def = resolveTableStyle('{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}', undefined, theme)!
    expect(def.firstRow?.fill).toEqual({ type: 'solid', color: '#ED7D31' })
  })

  it('No Style, Table Grid (themed2 without accent): all-dk1 grid including outer border', () => {
    const def = resolveTableStyle('{5940675A-B579-460E-94D1-54222C63F5DA}', undefined, theme)!
    expect(def.insideH?.fill).toEqual({ type: 'solid', color: '#000000' })
    expect(def.outer?.t?.fill).toEqual({ type: 'solid', color: '#000000' })
    expect(def.wholeTbl?.fill).toBeUndefined()
  })

  it('No Style, No Grid (official 307C) and the legacy made-up 307D both resolve to no style', () => {
    const noGrid = resolveTableStyle('{2D5ABB26-0587-4C30-8999-92F81FD0307C}', undefined, theme)!
    expect(noGrid.insideH).toBeUndefined()
    expect(noGrid.firstRow).toBeUndefined()
    expect(resolveTableStyle('{2D5ABB26-0587-4C30-8999-92F81FD0307D}', undefined, theme)).toEqual(
      {},
    )
  })

  it('Medium Style 3 - Accent2: first/last columns solid accent fill, outer border only top/bottom dk1', () => {
    const def = resolveTableStyle('{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}', undefined, theme)!
    expect(def.firstCol?.fill).toEqual({ type: 'solid', color: '#ED7D31' })
    expect(def.lastCol?.textColor).toBe('#FFFFFF')
    expect(def.outer?.t?.fill).toEqual({ type: 'solid', color: '#000000' })
    expect(def.outer?.l).toBeUndefined()
  })
})
