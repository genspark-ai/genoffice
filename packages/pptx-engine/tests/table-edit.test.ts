import { describe, it, expect } from 'vitest'
import { patchTableStyleXml, ensureTableStyleXml, TABLE_STYLE_PRESETS } from '../src/table-edit'
import { resolveTableStyle } from '../src/table-style'

/** Minimal table graphicFrame XML (with tblPr + 2 tc) */
const MINIMAL_TABLE_XML =
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="T"/></p:nvGraphicFramePr>' +
  '<a:graphic><a:graphicData>' +
  '<a:tbl>' +
  '<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{2D5ABB26-0587-4C30-8999-92F81FD0307C}</a:tableStyleId></a:tblPr>' +
  '<a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>' +
  '<a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody>' +
  '<a:tcPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:tcPr></a:tc></a:tr>' +
  '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>'

describe('patchTableStyleXml', () => {
  it('replaces the whole tblPr block (preset style)', () => {
    const preset = TABLE_STYLE_PRESETS['zebraBlue']!
    const result = patchTableStyleXml(MINIMAL_TABLE_XML, { tblPrXml: preset.tblPrXml })
    expect(result).toContain(preset.styleId)
    expect(result).not.toContain('{2D5ABB26-0587-4C30-8999-92F81FD0307C}')
  })

  it('changes only the firstRow flag, keeps bandRow and styleId', () => {
    const result = patchTableStyleXml(MINIMAL_TABLE_XML, { firstRow: false })
    expect(result).not.toContain('firstRow="1"')
    expect(result).toContain('bandRow="1"')
    expect(result).toContain('{2D5ABB26-0587-4C30-8999-92F81FD0307C}')
  })

  it('set shading color → solidFill inserted in every tcPr, replacing existing fills', () => {
    const result = patchTableStyleXml(MINIMAL_TABLE_XML, { shadingColor: '#AABBCC' })
    expect(result).toContain('<a:srgbClr val="AABBCC"/>')
    // The original FF0000 is replaced
    expect(result).not.toContain('FF0000')
  })

  it('clear shading (shadingColor=none) → noFill', () => {
    const result = patchTableStyleXml(MINIMAL_TABLE_XML, { shadingColor: 'none' })
    expect(result).toContain('<a:noFill/>')
    expect(result).not.toContain('FF0000')
  })

  it('all-borders preset → lnL/lnR/lnT/lnB inserted into tcPr', () => {
    const result = patchTableStyleXml(MINIMAL_TABLE_XML, {
      borderPreset: 'all',
      borderColor: '#000000',
      borderWidthEmu: 12700,
    })
    expect(result).toContain('<a:lnL')
    expect(result).toContain('<a:lnR')
    expect(result).toContain('<a:lnT')
    expect(result).toContain('<a:lnB')
  })

  it('clear-borders preset → tcPr gains w=0 noFill borders', () => {
    const result = patchTableStyleXml(MINIMAL_TABLE_XML, { borderPreset: 'none' })
    expect(result).toContain('<a:lnL w="0"><a:noFill/></a:lnL>')
    expect(result).toContain('<a:lnT w="0"><a:noFill/></a:lnT>')
  })

  it('clearDirectFormatting → removes direct tcPr fills and borders, keeps text and margins', () => {
    const xml =
      '<a:tbl><a:tblPr/><a:tr><a:tc><a:txBody><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="F4EFE6"/></a:solidFill></a:rPr><a:t>H</a:t></a:r></a:p></a:txBody>' +
      '<a:tcPr marL="171907" anchor="ctr"><a:lnL w="0"><a:solidFill><a:srgbClr val="C4B79E"/></a:solidFill></a:lnL><a:lnB w="0"/><a:solidFill><a:srgbClr val="0F3B3A"/></a:solidFill></a:tcPr></a:tc></a:tr></a:tbl>'
    const result = patchTableStyleXml(xml, {
      tblPrXml: TABLE_STYLE_PRESETS['zebraGray']!.tblPrXml,
      clearDirectFormatting: true,
    })
    expect(result).toContain(TABLE_STYLE_PRESETS['zebraGray']!.styleId!)
    expect(result).toContain('<a:tcPr marL="171907" anchor="ctr"></a:tcPr>')
    expect(result).not.toContain('0F3B3A')
    expect(result).not.toContain('lnL')
    // Run-level text color untouched (table style edits never override run-level colors)
    expect(result).toContain('F4EFE6')
  })

  it('cells scope: only the targeted cell is patched, siblings untouched', () => {
    const tc = (fill: string) =>
      `<a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></a:tcPr></a:tc>`
    const xml =
      '<p:graphicFrame><a:graphic><a:graphicData><a:tbl><a:tblPr/>' +
      `<a:tr h="1">${tc('111111')}${tc('222222')}</a:tr>` +
      `<a:tr h="1">${tc('333333')}${tc('444444')}</a:tr>` +
      '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
    const result = patchTableStyleXml(xml, {
      shadingColor: '#AABBCC',
      cells: [{ row: 1, col: 0 }],
    })
    expect(result).toContain('111111')
    expect(result).toContain('222222')
    expect(result).toContain('444444')
    expect(result).not.toContain('333333')
    expect(result).toContain('<a:srgbClr val="AABBCC"/>')
  })

  it('cells scope: a whole row and a cell without tcPr both work', () => {
    const bareTc = '<a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc>'
    const xml =
      '<a:tbl><a:tblPr/>' +
      `<a:tr h="1">${bareTc}${bareTc}</a:tr>` +
      `<a:tr h="1">${bareTc}${bareTc}</a:tr>` +
      '</a:tbl>'
    const result = patchTableStyleXml(xml, {
      shadingColor: '#EFF2FC',
      cells: [
        { row: 0, col: 0 },
        { row: 0, col: 1 },
      ],
    })
    // Row 0: both cells gain a created tcPr with the fill; row 1 untouched
    const row1 = result.slice(result.indexOf('<a:tr'), result.indexOf('</a:tr>'))
    expect((row1.match(/EFF2FC/g) ?? []).length).toBe(2)
    expect(result.match(/EFF2FC/g)?.length).toBe(2)
    expect(result).toContain(
      '<a:tcPr><a:solidFill><a:srgbClr val="EFF2FC"/></a:solidFill></a:tcPr></a:tc>',
    )
  })

  it('cells scope: shadingColor none writes noFill only in the target cell', () => {
    const tc = (fill: string) =>
      `<a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></a:tcPr></a:tc>`
    const xml = `<a:tbl><a:tblPr/><a:tr h="1">${tc('111111')}${tc('222222')}</a:tr></a:tbl>`
    const result = patchTableStyleXml(xml, {
      shadingColor: 'none',
      cells: [{ row: 0, col: 1 }],
    })
    expect(result).toContain('111111')
    expect(result).not.toContain('222222')
    expect((result.match(/<a:noFill\/>/g) ?? []).length).toBe(1)
  })

  it('all 8 preset styles have valid tblPrXml', () => {
    for (const [key, preset] of Object.entries(TABLE_STYLE_PRESETS)) {
      expect(preset.tblPrXml).toContain('<a:tblPr')
      expect(preset.tblPrXml).toContain('</a:tblPr>')
      expect(preset.description).toBeTruthy()
      // Check key coverage
      expect([
        'none',
        'lightGrid',
        'zebraBlue',
        'zebraGray',
        'headerDarkBlue',
        'headerOrange',
        'noBorder',
        'fullBorder',
      ]).toContain(key)
    }
    expect(Object.keys(TABLE_STYLE_PRESETS)).toHaveLength(8)
  })

  it('fixed-color presets: styleDefXml uses literal srgbClr, tblPr references the same styleId', () => {
    for (const key of ['zebraBlue', 'zebraGray', 'headerDarkBlue', 'headerOrange', 'noBorder']) {
      const p = TABLE_STYLE_PRESETS[key]!
      expect(p.styleId).toBeTruthy()
      expect(p.styleDefXml).toContain(`styleId="${p.styleId}"`)
      expect(p.tblPrXml).toContain(p.styleId!)
      expect(p.styleDefXml).not.toContain('schemeClr')
    }
    // Dark-blue/orange presets: header color matches the ribbon swatch
    expect(TABLE_STYLE_PRESETS['headerDarkBlue']!.styleDefXml).toContain('val="1F3864"')
    expect(TABLE_STYLE_PRESETS['headerOrange']!.styleDefXml).toContain('val="ED7D31"')
  })
})

describe('ensureTableStyleXml', () => {
  const DEF = TABLE_STYLE_PRESETS['headerDarkBlue']!.styleDefXml!
  const ID = TABLE_STYLE_PRESETS['headerDarkBlue']!.styleId!

  it('self-closing tblStyleLst → expanded with the definition inserted', () => {
    const src =
      '<?xml version="1.0"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'
    const out = ensureTableStyleXml(src, ID, DEF)
    expect(out).toContain(`styleId="${ID}"`)
    expect(out).toContain('</a:tblStyleLst>')
    expect(out).toContain('def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"')
  })

  it('idempotent when the definition already exists', () => {
    const once = ensureTableStyleXml(null, ID, DEF)
    expect(ensureTableStyleXml(once, ID, DEF)).toBe(once)
  })

  it('missing part (null) → created from the empty template', () => {
    const out = ensureTableStyleXml(null, ID, DEF)
    expect(out).toContain('<a:tblStyleLst')
    expect(out).toContain(`styleId="${ID}"`)
  })

  it('non-self-closing list → appended at the end', () => {
    const zebra = TABLE_STYLE_PRESETS['zebraBlue']!
    const withOne = ensureTableStyleXml(null, ID, DEF)
    const withTwo = ensureTableStyleXml(withOne, zebra.styleId!, zebra.styleDefXml!)
    expect(withTwo).toContain(`styleId="${ID}"`)
    expect(withTwo).toContain(`styleId="${zebra.styleId}"`)
  })

  it('loop back: resolveTableStyle parses the preset literal colors', () => {
    const xml = ensureTableStyleXml(null, ID, DEF)
    const def = resolveTableStyle(ID, xml, undefined)
    expect(def?.firstRow?.fill).toEqual({ type: 'solid', color: '#1F3864' })
    expect(def?.firstRow?.textColor).toBe('#FFFFFF')
    expect(def?.firstRow?.bold).toBe(true)
    expect(def?.band1H?.fill).toEqual({ type: 'solid', color: '#E9EDF5' })
    expect(def?.wholeTbl?.fill).toEqual({ type: 'solid', color: '#FFFFFF' })
    expect(def?.insideH?.fill).toEqual({ type: 'solid', color: '#D9D9D9' })
  })
})

describe('custom table style outer borders (tcBdr left/right/top/bottom)', () => {
  const XML = `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{X}">
    <a:tblStyle styleId="{X}" styleName="T"><a:wholeTbl><a:tcStyle><a:tcBdr>
      <a:left><a:ln w="9525"><a:solidFill><a:srgbClr val="9E9E9E"/></a:solidFill></a:ln></a:left>
      <a:right><a:ln w="9525"><a:solidFill><a:srgbClr val="9E9E9E"/></a:solidFill></a:ln></a:right>
      <a:top><a:ln w="9525"><a:solidFill><a:srgbClr val="9E9E9E"/></a:solidFill></a:ln></a:top>
      <a:bottom><a:ln w="9525"><a:solidFill><a:srgbClr val="9E9E9E"/></a:solidFill></a:ln></a:bottom>
      <a:insideH><a:ln w="9525"><a:solidFill><a:srgbClr val="9E9E9E"/></a:solidFill></a:ln></a:insideH>
      <a:insideV><a:ln w="9525"><a:solidFill><a:srgbClr val="9E9E9E"/></a:solidFill></a:ln></a:insideV>
    </a:tcBdr></a:tcStyle></a:wholeTbl></a:tblStyle></a:tblStyleLst>`
  it('reads the whole-table outer edges alongside inside lines', () => {
    const def = resolveTableStyle('{X}', XML, undefined)!
    for (const k of ['l', 'r', 't', 'b'] as const) {
      expect(def.outer?.[k]?.fill).toEqual({ type: 'solid', color: '#9E9E9E' })
      expect(def.outer?.[k]?.width).toBe(9525)
    }
    expect(def.insideH?.fill).toEqual({ type: 'solid', color: '#9E9E9E' })
  })
})
