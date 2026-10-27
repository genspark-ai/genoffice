import { describe, expect, it } from 'vitest'

import { applyDvRules } from '../src/gateway/xlsx-dv'

const SHEET =
  '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
  '<dataValidations count="1"><dataValidation type="whole" operator="equal" sqref="B1">' +
  '<formula1>7</formula1></dataValidation></dataValidations>' +
  '<hyperlinks><hyperlink ref="G1" r:id="rId1"/></hyperlinks>' +
  '<pageMargins left="0.7"/></worksheet>'

/// Excel's x14 data-validation extension. The literal is spelled out here on
/// purpose: a test that imports the constant from the source cannot tell the
/// right uri from the conditional-formatting one.
const X14_DV_EXT_URI = '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}'
/// Excel's x14 conditional-formatting extension — a different one, whose
/// validation rules Excel ignores.
const X14_CF_EXT_URI = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}'
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main'

const range = (column = 0) => ({ startRow: 0, endRow: 4, startColumn: column, endColumn: column })

/// The content of the `<ext>` element carrying `uri`, or null when the
/// worksheet has no such extension.
function extContent(xml: string, uri: string): string | null {
  const open = new RegExp(`<ext\\b[^>]*\\buri="${uri.replace(/[{}]/g, '\\$&')}"[^>]*>`).exec(xml)
  if (open === null) return null
  const close = xml.indexOf('</ext>', open.index)
  return close === -1 ? null : xml.slice(open.index + open[0].length, close)
}

describe('applyDvRules x14 extension writes', () => {
  it('routes cross-sheet list rules into the x14 block and omits the legacy twin', () => {
    const xml = applyDvRules(SHEET, [
      { ranges: [range()], rule: { type: 'whole', operator: 'equal', formula1: '1' } },
      {
        ranges: [range(3)],
        rule: { type: 'list', formula1: '=Sheet2!$A$1:$A$5', showDropDown: true },
      },
    ])
    // Legacy section: only the legacy-expressible rule, before the extLst.
    expect(xml).toContain(
      '<dataValidations count="1"><dataValidation type="whole" operator="equal" sqref="A1:A5">' +
        '<formula1>1</formula1></dataValidation></dataValidations><hyperlinks',
    )
    // x14 block: uri, both namespaces, the wrapped formula, and the sqref child.
    expect(xml).toContain(
      `<extLst><ext uri="${X14_DV_EXT_URI}" xmlns:x14="${X14_NS}">` +
        `<x14:dataValidations xmlns:x14="${X14_NS}" xmlns:xm="${XM_NS}" count="1">` +
        '<x14:dataValidation type="list">' +
        '<x14:formula1><xm:f>Sheet2!$A$1:$A$5</xm:f></x14:formula1>' +
        '<xm:sqref>D1:D5</xm:sqref></x14:dataValidation>' +
        '</x14:dataValidations></ext></extLst></worksheet>',
    )
  })

  it('puts the block in the data-validation extension, not the conditional-formatting one', () => {
    const xml = applyDvRules(SHEET, [
      { ranges: [range()], rule: { type: 'list', formula1: '=Sheet2!$A$1:$A$5' } },
    ])
    // The literal uri of Excel's x14 data-validation extension.
    expect(xml).toContain('<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"')
    // The conditional-formatting extension is a different one; Excel ignores
    // validation rules parked in it, so it must not appear at all.
    expect(xml).not.toContain('{78C0D931-6437-407d-A8EE-F0AAD7539E65}')
    expect(extContent(xml, '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}')).toContain(
      '<x14:dataValidations',
    )
  })

  it('routes cross-sheet custom and numeric-bound rules the same way', () => {
    const xml = applyDvRules(SHEET, [
      { ranges: [range()], rule: { type: 'custom', formula1: "=AND('Q 1'!A1>0)" } },
      {
        ranges: [range(1)],
        rule: { type: 'whole', operator: 'between', formula1: '=Sheet3!$A$1', formula2: '9' },
      },
    ])
    expect(xml).toContain('<x14:dataValidations xmlns:x14')
    expect(xml).toContain("<xm:f>AND('Q 1'!A1&gt;0)</xm:f>")
    expect(xml).toContain('<x14:formula1><xm:f>Sheet3!$A$1</xm:f></x14:formula1>')
    expect(xml).toContain('<x14:formula2><xm:f>9</xm:f></x14:formula2>')
    expect(xml).not.toContain('<dataValidations count=')
  })

  it('keeps a list literal that merely contains ! in the legacy section', () => {
    const xml = applyDvRules(SHEET, [
      { ranges: [range()], rule: { type: 'list', formula1: 'Yes! No,Maybe' } },
    ])
    expect(xml).toContain('<formula1>"Yes! No,Maybe"</formula1>')
    expect(xml).not.toContain('x14:dataValidations')
  })

  it('rewrites the existing x14 block, keeping the conditional-formatting extension verbatim', () => {
    const x14Sheet = SHEET.replace(
      '<pageMargins left="0.7"/>',
      `<pageMargins left="0.7"/><extLst>` +
        `<ext uri="${X14_CF_EXT_URI}" xmlns:x14="${X14_NS}">` +
        '<x14:conditionalFormattings><x14:conditionalFormatting ' +
        `xmlns:xm="${XM_NS}"><x14:cfRule type="dataBar" id="{DB1}"/><xm:sqref>F1:F4</xm:sqref>` +
        '</x14:conditionalFormatting></x14:conditionalFormattings></ext>' +
        `<ext uri="${X14_DV_EXT_URI}" xmlns:x14="${X14_NS}">` +
        `<x14:dataValidations xmlns:xm="${XM_NS}" count="1">` +
        '<x14:dataValidation type="list"><x14:formula1><xm:f>Old!$A$1</xm:f></x14:formula1>' +
        '<xm:sqref>C1:C9</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>',
    )
    const xml = applyDvRules(x14Sheet, [
      { ranges: [range(2)], rule: { type: 'list', formula1: '=Other!$B$2:$B$4' } },
    ])
    expect(xml).not.toContain('Old!$A$1')
    expect(xml).toContain('<xm:f>Other!$B$2:$B$4</xm:f>')
    expect(xml).toContain('<xm:sqref>C1:C5</xm:sqref>')
    // The conditional-formatting extension stays byte-identical.
    expect(extContent(xml, X14_CF_EXT_URI)).toBe(
      '<x14:conditionalFormattings><x14:conditionalFormatting ' +
        `xmlns:xm="${XM_NS}"><x14:cfRule type="dataBar" id="{DB1}"/><xm:sqref>F1:F4</xm:sqref>` +
        '</x14:conditionalFormatting></x14:conditionalFormattings>',
    )
    // The block lives in the data-validation extension, and nowhere else.
    expect(extContent(xml, X14_DV_EXT_URI)).toContain(
      `<x14:dataValidations xmlns:x14="${X14_NS}" xmlns:xm="${XM_NS}" count="1">` +
        '<x14:dataValidation type="list"><x14:formula1><xm:f>Other!$B$2:$B$4</xm:f></x14:formula1>' +
        '<xm:sqref>C1:C5</xm:sqref></x14:dataValidation>',
    )
    // The legacy whole rule from the file is gone (full snapshot rewrite).
    expect(xml).not.toContain('type="whole"')
  })

  it('moves a block an older build wrote into the conditional-formatting extension', () => {
    const wrongSheet = SHEET.replace(
      '<pageMargins left="0.7"/>',
      `<pageMargins left="0.7"/><extLst><ext uri="${X14_CF_EXT_URI}" xmlns:x14="${X14_NS}">` +
        `<x14:dataValidations xmlns:xm="${XM_NS}" count="1">` +
        '<x14:dataValidation type="list"><x14:formula1><xm:f>Old!$A$1</xm:f></x14:formula1>' +
        '<xm:sqref>C1:C9</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>',
    )
    const xml = applyDvRules(wrongSheet, [
      { ranges: [range()], rule: { type: 'list', formula1: '=New!$A$1:$A$2' } },
    ])
    // The rules end up where Excel looks for them.
    expect(extContent(xml, X14_DV_EXT_URI)).toContain('<xm:f>New!$A$1:$A$2</xm:f>')
    // The conditional-formatting extension lost its only element, so it goes
    // with the block instead of lingering as an empty husk Excel repairs.
    expect(extContent(xml, X14_CF_EXT_URI)).toBeNull()
    expect(xml).not.toContain(X14_CF_EXT_URI)
    // The extLst survives to hold the data-validation extension.
    expect(extContent(xml, X14_DV_EXT_URI)).toContain('<x14:dataValidations')
  })

  it('moves a rule back to legacy once its source is same-sheet', () => {
    const x14Sheet = SHEET.replace(
      '<pageMargins left="0.7"/>',
      `<pageMargins left="0.7"/><extLst><ext uri="${X14_DV_EXT_URI}" xmlns:x14="${X14_NS}">` +
        `<x14:dataValidations xmlns:xm="${XM_NS}" count="1">` +
        '<x14:dataValidation type="list"><x14:formula1><xm:f>Sheet9!$A$1:$A$2</xm:f></x14:formula1>' +
        '<xm:sqref>B1:B9</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>',
    )
    const xml = applyDvRules(x14Sheet, [
      { ranges: [range(1)], rule: { type: 'list', formula1: '=$E$1:$E$4' } },
    ])
    expect(xml).toContain('<formula1>$E$1:$E$4</formula1>')
    expect(xml).not.toContain('x14:dataValidations')
    expect(xml).not.toContain('<extLst')
  })

  it('removes the x14 block, the emptied extension, and the emptied extLst', () => {
    const x14Sheet = SHEET.replace(
      '<pageMargins left="0.7"/>',
      `<pageMargins left="0.7"/><extLst><ext uri="${X14_DV_EXT_URI}" xmlns:x14="${X14_NS}">` +
        `<x14:dataValidations xmlns:xm="${XM_NS}" count="2">` +
        '<x14:dataValidation type="list"><x14:formula1><xm:f>Old!$A$1</xm:f></x14:formula1>' +
        '<xm:sqref>C1:C9</xm:sqref></x14:dataValidation>' +
        '<x14:dataValidation type="whole"><x14:formula1><xm:f>1</xm:f></x14:formula1>' +
        '<xm:sqref>D1:D9</xm:sqref></x14:dataValidation>' +
        '</x14:dataValidations></ext></extLst>',
    )
    const xml = applyDvRules(x14Sheet, [])
    expect(xml).not.toContain('x14:dataValidations')
    expect(xml).not.toContain('<extLst')
    expect(xml).toContain('<hyperlinks>')
  })

  it('keeps the conditional-formatting extension when the data-validation one goes away', () => {
    const x14Sheet = SHEET.replace(
      '<pageMargins left="0.7"/>',
      `<pageMargins left="0.7"/><extLst>` +
        `<ext uri="${X14_CF_EXT_URI}" xmlns:x14="${X14_NS}"><x14:conditionalFormattings/></ext>` +
        `<ext uri="${X14_DV_EXT_URI}" xmlns:x14="${X14_NS}">` +
        `<x14:dataValidations xmlns:xm="${XM_NS}" count="1">` +
        '<x14:dataValidation type="list"><x14:formula1><xm:f>Old!$A$1</xm:f></x14:formula1>' +
        '<xm:sqref>C1:C9</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>',
    )
    const xml = applyDvRules(x14Sheet, [])
    expect(extContent(xml, X14_CF_EXT_URI)).toBe('<x14:conditionalFormattings/>')
    expect(extContent(xml, X14_DV_EXT_URI)).toBeNull()
    expect(xml).not.toContain('x14:dataValidations')
  })

  it('replaces a self-closing x14 dataValidations element', () => {
    const x14Sheet = SHEET.replace(
      '<pageMargins left="0.7"/>',
      `<pageMargins left="0.7"/><extLst><ext uri="${X14_DV_EXT_URI}" xmlns:x14="${X14_NS}">` +
        '<x14:dataValidations/></ext></extLst>',
    )
    const xml = applyDvRules(x14Sheet, [
      { ranges: [range()], rule: { type: 'list', formula1: '=Sheet2!$A$1:$A$5' } },
    ])
    expect(xml.match(/<x14:dataValidations\b/g)).toHaveLength(1)
    expect(xml).toContain('<xm:f>Sheet2!$A$1:$A$5</xm:f>')
  })

  it('leaves a legacy-only save byte-identical to the legacy path', () => {
    const xml = applyDvRules(SHEET, [
      { ranges: [range()], rule: { type: 'decimal', operator: 'greaterThan', formula1: '3' } },
    ])
    expect(xml).not.toContain('x14:')
    expect(xml).not.toContain('<extLst')
    expect(xml).toContain(
      '<dataValidations count="1"><dataValidation type="decimal" operator="greaterThan" ' +
        'sqref="A1:A5"><formula1>3</formula1></dataValidation></dataValidations><hyperlinks',
    )
  })

  it('appends x14 rules while trimming the areas the batch takes over', () => {
    const x14Sheet =
      '<worksheet><sheetData/>' +
      `<extLst><ext uri="${X14_DV_EXT_URI}" xmlns:x14="${X14_NS}">` +
      `<x14:dataValidations xmlns:xm="${XM_NS}" count="2">` +
      '<x14:dataValidation type="list"><x14:formula1><xm:f>Old!$A$1</xm:f></x14:formula1>' +
      '<xm:sqref>C1:C9 D2</xm:sqref></x14:dataValidation>' +
      '<x14:dataValidation type="whole"><x14:formula1><xm:f>2</xm:f></x14:formula1>' +
      '<xm:sqref>E1:E9</xm:sqref></x14:dataValidation>' +
      '</x14:dataValidations></ext></extLst></worksheet>'
    const c = { startRow: 0, endRow: 8, startColumn: 2, endColumn: 2 }
    const xml = applyDvRules(
      x14Sheet,
      [{ ranges: [c], rule: { type: 'list', formula1: '=New!$B$1:$B$3' } }],
      { append: true, remove: [{ startRow: 1, endRow: 1, startColumn: 3, endColumn: 3 }] },
    )
    expect(xml).toContain('<xm:f>New!$B$1:$B$3</xm:f>')
    // D2 was removed; the first entry keeps nothing, so it is gone.
    expect(xml).not.toContain('Old!$A$1')
    expect(xml).not.toContain('<xm:sqref>D2</xm:sqref>')
    // The untouched E1:E9 entry survives verbatim.
    expect(xml).toContain('<x14:dataValidation type="whole">')
    expect(xml).toContain('<xm:sqref>E1:E9</xm:sqref>')
    expect(xml.match(/<x14:dataValidation\b/g)).toHaveLength(2)
  })

  it('writes multiple x14 entries with a matching count', () => {
    const xml = applyDvRules(SHEET, [
      { ranges: [range()], rule: { type: 'list', formula1: '=Sheet2!$A$1:$A$5' } },
      {
        ranges: [range(1), range(2)],
        rule: { type: 'custom', formula1: '=S3!Z9' },
      },
    ])
    expect(xml).toContain('<x14:dataValidations xmlns:x14')
    expect(xml).toContain('count="2">')
    expect(xml).toContain('<xm:sqref>B1:B5 C1:C5</xm:sqref>')
  })
})
