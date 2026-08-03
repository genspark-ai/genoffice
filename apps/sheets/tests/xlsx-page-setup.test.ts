import { describe, expect, it } from 'vitest'

import {
  applyPageSetupState,
  applyPrintAreas,
  buildHeaderFooterXml,
  PageSetupError,
} from '../src/gateway/xlsx-page-setup'

const BARE = '<worksheet><sheetData/></worksheet>'
const WITH_VIEW =
  '<worksheet><sheetViews><sheetView workbookViewId="0"/></sheetViews>' + '<sheetData/></worksheet>'

describe('applyPageSetupState', () => {
  it('creates pageSetup with orientation and paper size before </worksheet>', () => {
    const xml = applyPageSetupState(BARE, {
      sheetName: 'Sheet1',
      orientation: 'landscape',
      paperSize: 9,
    })
    expect(xml).toBe(
      '<worksheet><sheetData/>' + '<pageSetup orientation="landscape" paperSize="9"/></worksheet>',
    )
  })

  it('merges attributes into an existing pageSetup, keeping the rest', () => {
    const xml =
      '<worksheet><sheetData/>' +
      '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
      '<pageSetup paperSize="1" orientation="portrait" r:id="rId1"/></worksheet>'
    const patched = applyPageSetupState(xml, { sheetName: 'S', orientation: 'landscape' })
    expect(patched).toContain('<pageSetup paperSize="1" orientation="landscape" r:id="rId1"/>')
    expect(patched).toContain('<pageMargins left="0.7"')
  })

  it('inserts pageSetup before headerFooter to keep schema order', () => {
    const xml =
      '<worksheet><sheetData/><headerFooter><oddHeader>x</oddHeader></headerFooter>' +
      '</worksheet>'
    const patched = applyPageSetupState(xml, { sheetName: 'S', orientation: 'portrait' })
    expect(patched.indexOf('<pageSetup')).toBeLessThan(patched.indexOf('<headerFooter'))
  })

  it('writes margin presets as a complete pageMargins element', () => {
    const xml = applyPageSetupState(BARE, { sheetName: 'S', margins: 'narrow' })
    expect(xml).toContain(
      '<pageMargins left="0.25" right="0.25" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
    )
  })

  it('orders printOptions before an existing pageMargins', () => {
    const xml = applyPageSetupState(
      applyPageSetupState(BARE, { sheetName: 'S', margins: 'wide' }),
      { sheetName: 'S', printGridlines: true, printHeadings: true },
    )
    expect(xml).toContain('<printOptions gridLines="1" headings="1"/>')
    expect(xml.indexOf('<printOptions')).toBeLessThan(xml.indexOf('<pageMargins'))
  })

  it('removes print options by dropping the attribute (schema default)', () => {
    const xml = '<worksheet><sheetData/><printOptions gridLines="1" headings="1"/></worksheet>'
    const patched = applyPageSetupState(xml, { sheetName: 'S', printGridlines: false })
    expect(patched).toContain('<printOptions headings="1"/>')
  })

  it('writes fit-to-page: pageSetUpPr plus fitToWidth/fitToHeight', () => {
    const xml = applyPageSetupState(BARE, {
      sheetName: 'S',
      fitToWidth: 1,
      fitToHeight: 0,
      fitToPage: true,
    })
    expect(xml).toContain('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>')
    // fitToWidth="1" is the schema default and is dropped; 0 = automatic.
    expect(xml).toContain('<pageSetup fitToHeight="0"/>')
  })

  it('clears fitToPage when the user switches to a fixed scale', () => {
    const fitted = applyPageSetupState(BARE, {
      sheetName: 'S',
      fitToWidth: 2,
      fitToPage: true,
    })
    const scaled = applyPageSetupState(fitted, { sheetName: 'S', scale: 80, fitToPage: false })
    expect(scaled).not.toContain('fitToPage="1"')
    expect(scaled).toContain('scale="80"')
  })

  it('toggles sheetView@showFormulas, creating sheetViews when missing', () => {
    const shown = applyPageSetupState(WITH_VIEW, { sheetName: 'S', showFormulas: true })
    expect(shown).toContain('showFormulas="1"')
    const restored = applyPageSetupState(shown, { sheetName: 'S', showFormulas: false })
    expect(restored).toBe(WITH_VIEW)
    const created = applyPageSetupState(BARE, { sheetName: 'S', showFormulas: true })
    expect(created).toContain(
      '<sheetViews><sheetView workbookViewId="0" showFormulas="1"/></sheetViews>',
    )
  })

  it('toggles sheetView@showGridLines, creating sheetViews when missing', () => {
    const hidden = applyPageSetupState(WITH_VIEW, { sheetName: 'S', showGridlines: false })
    expect(hidden).toContain('<sheetView showGridLines="0" workbookViewId="0"/>')
    const restored = applyPageSetupState(hidden, { sheetName: 'S', showGridlines: true })
    expect(restored).toBe(WITH_VIEW)
    const created = applyPageSetupState(BARE, { sheetName: 'S', showGridlines: false })
    expect(created).toContain(
      '<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>',
    )
    expect(created.indexOf('<sheetViews>')).toBeLessThan(created.indexOf('<sheetData'))
  })
})

describe('buildHeaderFooterXml', () => {
  it('joins left/center/right with their section markers', () => {
    expect(
      buildHeaderFooterXml(
        { left: 'L1', center: 'C1', right: 'R1' },
        { left: 'L2', center: 'C2', right: 'R2' },
      ),
    ).toBe(
      '<headerFooter><oddHeader>&amp;LL1&amp;CC1&amp;RR1</oddHeader>' +
        '<oddFooter>&amp;LL2&amp;CC2&amp;RR2</oddFooter></headerFooter>',
    )
  })

  it('omits empty sections and an empty half', () => {
    expect(buildHeaderFooterXml({ center: 'Report' }, null)).toBe(
      '<headerFooter><oddHeader>&amp;CReport</oddHeader></headerFooter>',
    )
    expect(buildHeaderFooterXml(null, { left: 'Q1', right: 'Page &P of &N' })).toBe(
      '<headerFooter><oddFooter>&amp;LQ1&amp;RPage &amp;P of &amp;N</oddFooter></headerFooter>',
    )
  })

  it('returns an empty string when everything is empty', () => {
    expect(buildHeaderFooterXml(null, null)).toBe('')
    expect(buildHeaderFooterXml(undefined, undefined)).toBe('')
    expect(buildHeaderFooterXml({}, { left: '', center: '', right: '' })).toBe('')
  })

  it('XML-escapes the content, literal ampersands included', () => {
    expect(buildHeaderFooterXml({ left: 'P&&L <2026> "Q1"' }, null)).toBe(
      '<headerFooter><oddHeader>&amp;LP&amp;&amp;L &lt;2026&gt; &quot;Q1&quot;</oddHeader>' +
        '</headerFooter>',
    )
  })
})

describe('applyPageSetupState headerFooter', () => {
  it('creates headerFooter after pageMargins and before rowBreaks', () => {
    const xml = applyPageSetupState(
      '<worksheet><sheetData/>' +
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
        '<rowBreaks count="1"/></worksheet>',
      { sheetName: 'S', header: { center: 'Title' }, footer: null },
    )
    expect(xml).toContain('<headerFooter><oddHeader>&amp;CTitle</oddHeader></headerFooter>')
    expect(xml.indexOf('<pageMargins')).toBeLessThan(xml.indexOf('<headerFooter'))
    expect(xml.indexOf('<headerFooter')).toBeLessThan(xml.indexOf('<rowBreaks'))
  })

  it('replaces only the touched half of an existing headerFooter', () => {
    const xml =
      '<worksheet><sheetData/><headerFooter alignWithMargins="0">' +
      '<oddHeader>&amp;LOld</oddHeader><oddFooter>&amp;RKeep</oddFooter>' +
      '</headerFooter></worksheet>'
    const patched = applyPageSetupState(xml, { sheetName: 'S', header: { center: 'New' } })
    expect(patched).toContain('<oddHeader>&amp;CNew</oddHeader>')
    expect(patched).toContain('<oddFooter>&amp;RKeep</oddFooter>')
    expect(patched).not.toContain('Old')
  })

  it('removes the element when header and footer both clear', () => {
    const xml =
      '<worksheet><sheetData/>' +
      '<headerFooter><oddHeader>&amp;CGone</oddHeader></headerFooter></worksheet>'
    const cleared = applyPageSetupState(xml, { sheetName: 'S', header: null, footer: null })
    expect(cleared).toBe('<worksheet><sheetData/></worksheet>')
    // No element to remove is a no-op, not an insertion of an empty one.
    expect(applyPageSetupState(BARE, { sheetName: 'S', header: null, footer: null })).toBe(BARE)
  })
})

const WORKBOOK =
  '<workbook><sheets>' +
  '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>' +
  '<sheet name="P&amp;L" sheetId="2" r:id="rId2"/>' +
  '</sheets></workbook>'

describe('applyPrintAreas', () => {
  it('creates the sheet-scoped _xlnm.Print_Area with absolute references', () => {
    const xml = applyPrintAreas(WORKBOOK, [{ sheetName: 'Sheet1', printArea: 'A1:C10' }])
    expect(xml).toContain(
      '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">' +
        "'Sheet1'!$A$1:$C$10</definedName></definedNames>",
    )
  })

  it('uses the workbook.xml sheet position for localSheetId and escapes names', () => {
    const xml = applyPrintAreas(WORKBOOK, [{ sheetName: 'P&L', printArea: 'B2:D4' }])
    expect(xml).toContain('localSheetId="1">\'P&amp;L\'!$B$2:$D$4')
  })

  it('replaces an existing print area and clears on null', () => {
    const withArea = applyPrintAreas(WORKBOOK, [{ sheetName: 'Sheet1', printArea: 'A1:C10' }])
    const updated = applyPrintAreas(withArea, [{ sheetName: 'Sheet1', printArea: 'A1:B2' }])
    expect(updated).toContain('$A$1:$B$2')
    expect(updated).not.toContain('$C$10')
    const cleared = applyPrintAreas(updated, [{ sheetName: 'Sheet1', printArea: null }])
    expect(cleared).toBe(WORKBOOK)
  })

  it("keeps other sheets' print areas when clearing one", () => {
    const both = applyPrintAreas(WORKBOOK, [
      { sheetName: 'Sheet1', printArea: 'A1:C10' },
      { sheetName: 'P&L', printArea: 'B2:D4' },
    ])
    const cleared = applyPrintAreas(both, [{ sheetName: 'Sheet1', printArea: null }])
    expect(cleared).not.toContain('localSheetId="0"')
    expect(cleared).toContain('localSheetId="1"')
  })

  it('rejects unknown sheets and malformed ranges', () => {
    expect(() => applyPrintAreas(WORKBOOK, [{ sheetName: 'Nope', printArea: 'A1:B2' }])).toThrow(
      PageSetupError,
    )
    expect(() => applyPrintAreas(WORKBOOK, [{ sheetName: 'Sheet1', printArea: 'A1;B2' }])).toThrow(
      PageSetupError,
    )
  })

  it('writes and clears _xlnm.Print_Titles as absolute row spans', () => {
    const xml = applyPrintAreas(WORKBOOK, [{ sheetName: 'Sheet1', printTitles: '1:2' }])
    expect(xml).toContain(
      '<definedName name="_xlnm.Print_Titles" localSheetId="0">' + "'Sheet1'!$1:$2</definedName>",
    )
    const cleared = applyPrintAreas(xml, [{ sheetName: 'Sheet1', printTitles: null }])
    expect(cleared).toBe(WORKBOOK)
  })

  it('keeps print area and titles independent on one sheet', () => {
    const both = applyPrintAreas(WORKBOOK, [
      { sheetName: 'Sheet1', printArea: 'A1:C10', printTitles: '1:1' },
    ])
    expect(both).toContain('_xlnm.Print_Area')
    expect(both).toContain('_xlnm.Print_Titles')
    const titlesOnly = applyPrintAreas(both, [{ sheetName: 'Sheet1', printArea: null }])
    expect(titlesOnly).not.toContain('_xlnm.Print_Area')
    expect(titlesOnly).toContain("'Sheet1'!$1:$1")
    expect(() => applyPrintAreas(WORKBOOK, [{ sheetName: 'Sheet1', printTitles: '3:1' }])).toThrow(
      PageSetupError,
    )
  })
})
