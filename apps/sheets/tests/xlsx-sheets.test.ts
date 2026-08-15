import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, assertOnlyTouchedEntriesChanged } from '../src/gateway/xlsx-gateway'
import { relsPathFor, resolveRelTarget } from '../src/gateway/xlsx-drawing-add'
import {
  classifyRemovedSheetRels,
  definedNamesUseToken,
  formatSheetQualifier,
  partPathForRels,
  pivotCacheReadsFromSheet,
  prepareClonedSheetRels,
  renameSheetInFormula,
  renameSheetInPivotCacheSource,
  sanitizeClonedWorksheetXml,
  SheetEditError,
  stripPageSetupRelIds,
  tableDisplayName,
  validateSheetName,
} from '../src/gateway/xlsx-sheets'
import {
  buildCompatibilityFixture,
  buildSatelliteSheetFixture,
  buildSheetsFixture,
} from './fixture-builder'

async function entryText(buffer: Buffer, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(buffer)
  return zip.file(path)?.async('text') ?? ''
}

describe('sheet name rules', () => {
  it('accepts ordinary names and rejects forbidden shapes', () => {
    expect(() => validateSheetName('Budget 2026')).not.toThrow()
    expect(() => validateSheetName('')).toThrow(SheetEditError)
    expect(() => validateSheetName('x'.repeat(32))).toThrow(SheetEditError)
    expect(() => validateSheetName('a[b')).toThrow(SheetEditError)
    expect(() => validateSheetName('a:b')).toThrow(SheetEditError)
    expect(() => validateSheetName("'edge")).toThrow(SheetEditError)
  })

  it('quotes qualifiers only when Excel requires it', () => {
    expect(formatSheetQualifier('Data')).toBe('Data')
    expect(formatSheetQualifier('My Sheet')).toBe("'My Sheet'")
    expect(formatSheetQualifier('A1')).toBe("'A1'")
    expect(formatSheetQualifier('TRUE')).toBe("'TRUE'")
    expect(formatSheetQualifier("O'Brien")).toBe("'O''Brien'")
  })
})

describe('renameSheetInFormula', () => {
  it('rewrites qualified references and leaves the rest alone', () => {
    expect(renameSheetInFormula('Data!A1+SUM(Data!B1:B9)+A1', 'Data', 'Numbers')).toBe(
      'Numbers!A1+SUM(Numbers!B1:B9)+A1',
    )
    expect(renameSheetInFormula('\'My Sheet\'!$A$1&"Data!A1"', 'My Sheet', 'Plain')).toBe(
      'Plain!$A$1&"Data!A1"',
    )
    expect(renameSheetInFormula('Other!A1', 'Data', 'Numbers')).toBe('Other!A1')
  })
})

describe('sheet visibility and order save', () => {
  it('writes state="hidden", reorders tabs, and drops calcChain on reorder', async () => {
    const mutation = await applyCellEditsToXlsx(await buildSheetsFixture(), [], [], [], {
      renames: [],
      additions: [],
      removals: [],
      hiddenChanges: [{ sheetName: 'Other', hidden: true }],
      orderChanged: true,
      order: ['My Sheet', 'Data', 'Other'],
    })
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const workbook = await entryText(mutation.buffer, 'xl/workbook.xml')
    expect(workbook).toContain('<sheet state="hidden" name="Other"')
    const sheetNames = [...workbook.matchAll(/<sheet[^>]*\bname="([^"]+)"/g)].map((m) => m[1])
    expect(sheetNames).toEqual(['My Sheet', 'Data', 'Other'])
  })

  it('unhides by removing the state attribute and rejects hiding every sheet', async () => {
    const source = await buildSheetsFixture()
    await expect(
      applyCellEditsToXlsx(source, [], [], [], {
        renames: [],
        additions: [],
        removals: [],
        hiddenChanges: [
          { sheetName: 'Data', hidden: true },
          { sheetName: 'Other', hidden: true },
          { sheetName: 'My Sheet', hidden: true },
        ],
        order: ['Data', 'Other', 'My Sheet'],
      }),
    ).rejects.toThrow(/at least one visible sheet/)
  })
})

describe('sheet rename save', () => {
  it('renames in workbook.xml and rewrites formulas, chart series, defined names, and hyperlink anchors', async () => {
    const mutation = await applyCellEditsToXlsx(await buildSheetsFixture(), [], [], [], {
      renames: [{ sheetName: 'Data', newName: 'Numbers' }],
      additions: [],
      removals: [],
      order: ['Numbers', 'Other', 'My Sheet'],
    })
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const workbook = await entryText(mutation.buffer, 'xl/workbook.xml')
    expect(workbook).toContain('<sheet name="Numbers" sheetId="1" r:id="rId1"/>')
    expect(workbook).toContain('<definedName name="GlobalTotal">Numbers!$A$1</definedName>')
    const other = await entryText(mutation.buffer, 'xl/worksheets/sheet2.xml')
    expect(other).toContain("<f>Numbers!A1+SUM('My Sheet'!A1:A2)</f>")
    expect(other).toContain('location="Numbers!A1"')
    const chart = await entryText(mutation.buffer, 'xl/charts/chart1.xml')
    expect(chart).toContain('<c:f>Numbers!$A$1:$A$2</c:f>')
    // The renamed sheet's own content had no self-references; untouched.
    expect(mutation.touchedEntries).not.toContain('xl/worksheets/sheet1.xml')
  })

  it('quotes the new name when needed and unquotes when not', async () => {
    const mutation = await applyCellEditsToXlsx(await buildSheetsFixture(), [], [], [], {
      renames: [{ sheetName: 'My Sheet', newName: 'Plain' }],
      additions: [],
      removals: [],
      order: ['Data', 'Other', 'Plain'],
    })
    const data = await entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(data).toContain('<f>Plain!A1*2</f>')
    const other = await entryText(mutation.buffer, 'xl/worksheets/sheet2.xml')
    expect(other).toContain('<f>Data!A1+SUM(Plain!A1:A2)</f>')
    const workbook = await entryText(mutation.buffer, 'xl/workbook.xml')
    expect(workbook).toContain(
      '<definedName name="LocalMy" localSheetId="2">Plain!$A$1</definedName>',
    )

    const spaced = await applyCellEditsToXlsx(await buildSheetsFixture(), [], [], [], {
      renames: [{ sheetName: 'Data', newName: 'New Data' }],
      additions: [],
      removals: [],
      order: ['New Data', 'Other', 'My Sheet'],
    })
    const spacedOther = await entryText(spaced.buffer, 'xl/worksheets/sheet2.xml')
    expect(spacedOther).toContain("<f>'New Data'!A1+SUM('My Sheet'!A1:A2)</f>")
  })

  it('rejects duplicate and invalid names', async () => {
    await expect(
      applyCellEditsToXlsx(await buildSheetsFixture(), [], [], [], {
        renames: [{ sheetName: 'Data', newName: 'Other' }],
        additions: [],
        removals: [],
        order: ['Other', 'Other', 'My Sheet'],
      }),
    ).rejects.toThrow('same name')
    await expect(
      applyCellEditsToXlsx(await buildSheetsFixture(), [], [], [], {
        renames: [{ sheetName: 'Data', newName: 'bad[name' }],
        additions: [],
        removals: [],
        order: ['bad[name', 'Other', 'My Sheet'],
      }),
    ).rejects.toThrow('forbidden character')
  })
})

describe('sheet add save', () => {
  it('creates the part with its override, relationship, and tab entry, and accepts cell edits', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildSheetsFixture(),
      [{ sheetName: 'Extra', row: 0, column: 0, writeValue: true, cell: { value: 'hello' } }],
      [],
      [],
      {
        renames: [],
        additions: [{ name: 'Extra' }],
        removals: [],
        order: ['Data', 'Other', 'My Sheet', 'Extra'],
      },
    )
    expect(mutation.addedEntries).toEqual(['xl/worksheets/sheet4.xml'])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const workbook = await entryText(mutation.buffer, 'xl/workbook.xml')
    expect(workbook).toContain('<sheet name="Extra" sheetId="4" r:id="rId6"/>')
    const contentTypes = await entryText(mutation.buffer, '[Content_Types].xml')
    expect(contentTypes).toContain('PartName="/xl/worksheets/sheet4.xml"')
    const rels = await entryText(mutation.buffer, 'xl/_rels/workbook.xml.rels')
    expect(rels).toContain('Id="rId6"')
    expect(rels).toContain('Target="worksheets/sheet4.xml"')
    const added = await entryText(mutation.buffer, 'xl/worksheets/sheet4.xml')
    expect(added).toContain('<is><t xml:space="preserve">hello</t></is>')
    // Adding a sheet shifts tab indexes, so calcChain is dropped for rebuild.
    expect(mutation.removedEntries).toContain('xl/calcChain.xml')
  })
})

describe('sheet remove save', () => {
  it('removes the part, its override and relationship, renumbers scoped names, and fixes the active tab', async () => {
    const mutation = await applyCellEditsToXlsx(await buildSheetsFixture(), [], [], [], {
      renames: [],
      additions: [],
      removals: ['Other'],
      order: ['Data', 'My Sheet'],
    })
    expect(mutation.removedEntries).toEqual(['xl/calcChain.xml', 'xl/worksheets/sheet2.xml'])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const zip = await JSZip.loadAsync(mutation.buffer)
    expect(zip.file('xl/worksheets/sheet2.xml')).toBeNull()
    const workbook = await entryText(mutation.buffer, 'xl/workbook.xml')
    expect(workbook).not.toContain('name="Other"')
    // The Print_Area scoped to the removed sheet dies with it; the survivor
    // is renumbered from tab 2 to tab 1, as is the active tab.
    expect(workbook).not.toContain('Print_Area')
    expect(workbook).toContain('<definedName name="LocalMy" localSheetId="1">')
    expect(workbook).toContain('activeTab="1"')
    const contentTypes = await entryText(mutation.buffer, '[Content_Types].xml')
    expect(contentTypes).not.toContain('sheet2.xml')
    const rels = await entryText(mutation.buffer, 'xl/_rels/workbook.xml.rels')
    expect(rels).not.toContain('sheet2.xml')
  })

  it('fails closed when formulas, charts, or defined names still reference the sheet', async () => {
    await expect(
      applyCellEditsToXlsx(await buildSheetsFixture(), [], [], [], {
        renames: [],
        additions: [],
        removals: ['My Sheet'],
        order: ['Data', 'Other'],
      }),
    ).rejects.toThrow('reference "My Sheet"')
    await expect(
      applyCellEditsToXlsx(await buildSheetsFixture(), [], [], [], {
        renames: [],
        additions: [],
        removals: ['Data'],
        order: ['Other', 'My Sheet'],
      }),
    ).rejects.toThrow('"Data"')
  })

  it('refuses to remove the last visible sheet', async () => {
    await expect(
      applyCellEditsToXlsx(await buildCompatibilityFixture(), [], [], [], {
        renames: [],
        additions: [],
        removals: ['Sheet1'],
        order: [],
      }),
    ).rejects.toThrow('visible sheet')
  })

  it('classifies removed-sheet relationships into owned, ignorable, and blocking', () => {
    const rels =
      '<Relationships>' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>' +
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/>' +
      '</Relationships>'
    expect(classifyRemovedSheetRels(rels, 'Data')).toEqual([
      '../drawings/drawing1.xml',
      '../comments1.xml',
    ])
    const pivotRels =
      '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>'
    expect(() => classifyRemovedSheetRels(pivotRels, 'Data')).toThrow(/pivot table/)
    const slicerRels =
      '<Relationships><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/></Relationships>'
    expect(() => classifyRemovedSheetRels(slicerRels, 'Data')).toThrow(SheetEditError)
  })

  it('rel-path helpers resolve targets and rels locations', () => {
    expect(resolveRelTarget('xl/worksheets/sheet2.xml', '../drawings/drawing1.xml')).toBe(
      'xl/drawings/drawing1.xml',
    )
    expect(resolveRelTarget('xl/worksheets/sheet2.xml', '../comments1.xml')).toBe(
      'xl/comments1.xml',
    )
    expect(resolveRelTarget('xl/charts/chart1.xml', 'colors1.xml')).toBe('xl/charts/colors1.xml')
    expect(relsPathFor('xl/drawings/drawing1.xml')).toBe('xl/drawings/_rels/drawing1.xml.rels')
    expect(partPathForRels('xl/drawings/_rels/drawing1.xml.rels')).toBe('xl/drawings/drawing1.xml')
    expect(partPathForRels('_rels/.rels')).toBe('')
    expect(tableDisplayName('<table id="1" name="T1" displayName="DecoTable" ref="A1:A2"/>')).toBe(
      'DecoTable',
    )
  })

  it('definedNamesUseToken sees surviving names and skips names scoped to removed sheets', () => {
    const scoped =
      '<workbook><definedNames>' +
      '<definedName name="Scoped" localSheetId="1">SUM(DecoTable[Amt])</definedName>' +
      '</definedNames></workbook>'
    expect(definedNamesUseToken(scoped, 'DecoTable[', new Set([1]))).toBe(false)
    expect(definedNamesUseToken(scoped, 'DecoTable[', new Set())).toBe(true)
    const global =
      '<workbook><definedNames>' +
      '<definedName name="Global">SUM(DecoTable[Amt])</definedName>' +
      '</definedNames></workbook>'
    expect(definedNamesUseToken(global, 'DecoTable[', new Set([1]))).toBe(true)
    expect(definedNamesUseToken(global, 'OtherTable[', new Set())).toBe(false)
    // table names are case-insensitive in Excel
    const cased =
      '<workbook><definedNames>' +
      '<definedName name="Global">SUM(decotable[Amt])</definedName>' +
      '</definedNames></workbook>'
    expect(definedNamesUseToken(cased, 'DecoTable[', new Set())).toBe(true)
  })

  it('pivotCacheReadsFromSheet matches the worksheetSource sheet case-insensitively', () => {
    const cache =
      '<pivotCacheDefinition><cacheSource type="worksheet">' +
      '<worksheetSource ref="A1:B9" sheet="P&amp;L Data"/>' +
      '</cacheSource></pivotCacheDefinition>'
    expect(pivotCacheReadsFromSheet(cache, 'P&L Data')).toBe(true)
    expect(pivotCacheReadsFromSheet(cache, 'p&l data')).toBe(true)
    expect(pivotCacheReadsFromSheet(cache, 'Other')).toBe(false)
    expect(pivotCacheReadsFromSheet('<pivotCacheDefinition/>', 'P&L Data')).toBe(false)
  })

  it('renameSheetInPivotCacheSource rewrites only matching sources and encodes the new name', () => {
    const cache =
      '<pivotCacheDefinition><cacheSource type="worksheet">' +
      '<worksheetSource ref="A1:B9" sheet="Data"/>' +
      '</cacheSource></pivotCacheDefinition>'
    expect(renameSheetInPivotCacheSource(cache, 'Data', 'P&L Data')).toContain(
      'sheet="P&amp;L Data"',
    )
    expect(renameSheetInPivotCacheSource(cache, 'Other', 'X')).toBe(cache)
  })

  it('cascade-removes the drawings, charts, images, comments, VML, and tables the sheet owns', async () => {
    const mutation = await applyCellEditsToXlsx(await buildSatelliteSheetFixture(), [], [], [], {
      renames: [],
      additions: [],
      removals: ['Deco'],
      order: ['Keep'],
    })
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    expect([...mutation.removedEntries].sort()).toEqual([
      'xl/charts/_rels/chart1.xml.rels',
      'xl/charts/chart1.xml',
      'xl/charts/colors1.xml',
      'xl/comments1.xml',
      'xl/drawings/_rels/drawing1.xml.rels',
      'xl/drawings/drawing1.xml',
      'xl/drawings/vmlDrawing1.vml',
      'xl/media/image1.png',
      'xl/tables/table1.xml',
      'xl/worksheets/_rels/sheet2.xml.rels',
      'xl/worksheets/sheet2.xml',
    ])
    const contentTypes = await entryText(mutation.buffer, '[Content_Types].xml')
    for (const gone of [
      'sheet2.xml',
      'drawing1.xml',
      'chart1.xml',
      'colors1.xml',
      'comments1.xml',
      'table1.xml',
    ]) {
      expect(contentTypes).not.toContain(gone)
    }
    const workbook = await entryText(mutation.buffer, 'xl/workbook.xml')
    expect(workbook).not.toContain('name="Deco"')
  })

  it('keeps an image a surviving sheet still references', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildSatelliteSheetFixture({ sharedImage: true }),
      [],
      [],
      [],
      { renames: [], additions: [], removals: ['Deco'], order: ['Keep'] },
    )
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    expect(mutation.removedEntries).not.toContain('xl/media/image1.png')
    expect(mutation.removedEntries).toContain('xl/drawings/drawing1.xml')
    const zip = await JSZip.loadAsync(mutation.buffer)
    expect(zip.file('xl/media/image1.png')).not.toBeNull()
    expect(zip.file('xl/drawings/drawing2.xml')).not.toBeNull()
  })

  it('renames another sheet in the same save as the cascade removal', async () => {
    // The rename pass rewrites chart references; charts dying with the
    // removed sheet are already gone from the package and must be skipped.
    const mutation = await applyCellEditsToXlsx(await buildSatelliteSheetFixture(), [], [], [], {
      renames: [{ sheetName: 'Keep', newName: 'Kept' }],
      additions: [],
      removals: ['Deco'],
      order: ['Kept'],
    })
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const workbook = await entryText(mutation.buffer, 'xl/workbook.xml')
    expect(workbook).toContain('name="Kept"')
    expect(workbook).not.toContain('name="Deco"')
    expect(mutation.removedEntries).toContain('xl/charts/chart1.xml')
  })

  it('ignores a structured reference in a defined name scoped to the removed sheet', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildSatelliteSheetFixture({ scopedTableName: true }),
      [],
      [],
      [],
      { renames: [], additions: [], removals: ['Deco'], order: ['Keep'] },
    )
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    // The scoped name dies with the sheet.
    const workbook = await entryText(mutation.buffer, 'xl/workbook.xml')
    expect(workbook).not.toContain('DecoRef')
  })

  it('fails closed on pivot tables and on surviving structured references into the table', async () => {
    await expect(
      applyCellEditsToXlsx(await buildSatelliteSheetFixture({ pivot: true }), [], [], [], {
        renames: [],
        additions: [],
        removals: ['Deco'],
        order: ['Keep'],
      }),
    ).rejects.toThrow(/pivot table/)
    await expect(
      applyCellEditsToXlsx(await buildSatelliteSheetFixture({ tableRef: true }), [], [], [], {
        renames: [],
        additions: [],
        removals: ['Deco'],
        order: ['Keep'],
      }),
    ).rejects.toThrow(/table "DecoTable"/)
  })

  it('fails closed on a differently-cased structured reference into the table', async () => {
    await expect(
      applyCellEditsToXlsx(await buildSatelliteSheetFixture({ tableRefLower: true }), [], [], [], {
        renames: [],
        additions: [],
        removals: ['Deco'],
        order: ['Keep'],
      }),
    ).rejects.toThrow(/table "DecoTable"/)
  })

  it('fails closed when a pivot cache reads its source rows from the removed sheet', async () => {
    await expect(
      applyCellEditsToXlsx(
        await buildSatelliteSheetFixture({ pivotSourceCache: true }),
        [],
        [],
        [],
        { renames: [], additions: [], removals: ['Deco'], order: ['Keep'] },
      ),
    ).rejects.toThrow(/source data/)
  })

  it('rewrites a pivot cache worksheetSource when its source sheet is renamed', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildSatelliteSheetFixture({ pivotSourceCache: true }),
      [],
      [],
      [],
      {
        renames: [{ sheetName: 'Deco', newName: 'Deco 2' }],
        additions: [],
        removals: [],
        order: ['Keep', 'Deco 2'],
      },
    )
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const cache = await entryText(mutation.buffer, 'xl/pivotCache/pivotCacheDefinition1.xml')
    expect(cache).toContain('sheet="Deco 2"')
    expect(cache).not.toContain('sheet="Deco"')
  })
})

describe('sheet duplicate save', () => {
  it('seeds the new part from the source and applies edits on top', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildSheetsFixture(),
      [{ sheetName: 'Data Copy', row: 0, column: 1, writeValue: true, cell: { value: 'copied' } }],
      [],
      [],
      {
        renames: [],
        additions: [{ name: 'Data Copy', sourceSheetName: 'Data' }],
        removals: [],
        order: ['Data', 'Data Copy', 'Other', 'My Sheet'],
      },
    )
    expect(mutation.addedEntries).toEqual(['xl/worksheets/sheet4.xml'])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const workbook = await entryText(mutation.buffer, 'xl/workbook.xml')
    expect(workbook).toContain('<sheet name="Data Copy" sheetId="4" r:id="rId6"/>')
    const copy = await entryText(mutation.buffer, 'xl/worksheets/sheet4.xml')
    // Source content survives — including its formula — plus the new edit.
    expect(copy).toContain("<f>'My Sheet'!A1*2</f>")
    expect(copy).toContain(
      '<c r="B1" t="inlineStr"><is><t xml:space="preserve">copied</t></is></c>',
    )
    // The source part itself is byte-untouched.
    expect(mutation.touchedEntries).not.toContain('xl/worksheets/sheet1.xml')
    const source = await entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(source).not.toContain('copied')
    expect(mutation.removedEntries).toContain('xl/calcChain.xml')
  })

  it('clones hyperlink relationships and drops printer settings', async () => {
    const zip = await JSZip.loadAsync(await buildSheetsFixture())
    const sheet1 = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    zip.file(
      'xl/worksheets/sheet1.xml',
      sheet1.replace(
        '</worksheet>',
        '<pageSetup orientation="landscape" r:id="rId2"/></worksheet>',
      ),
    )
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/>' +
        '</Relationships>',
    )
    zip.file('xl/printerSettings/printerSettings1.bin', Buffer.from([1, 2, 3]))
    const source = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    const mutation = await applyCellEditsToXlsx(source, [], [], [], {
      renames: [],
      additions: [{ name: 'Data Copy', sourceSheetName: 'Data' }],
      removals: [],
      order: ['Data', 'Data Copy', 'Other', 'My Sheet'],
    })
    expect(mutation.addedEntries).toEqual([
      'xl/worksheets/_rels/sheet4.xml.rels',
      'xl/worksheets/sheet4.xml',
    ])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const clonedRels = await entryText(mutation.buffer, 'xl/worksheets/_rels/sheet4.xml.rels')
    expect(clonedRels).toContain('Target="https://example.com"')
    expect(clonedRels).not.toContain('printerSettings')
    const copy = await entryText(mutation.buffer, 'xl/worksheets/sheet4.xml')
    expect(copy).toContain('<pageSetup orientation="landscape"/>')
    expect(copy).not.toContain('r:id="rId2"')
  })

  it('fails closed on drawings, tables, and sheet-scoped defined names', async () => {
    const zip = await JSZip.loadAsync(await buildSheetsFixture())
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>' +
        '</Relationships>',
    )
    const withDrawing = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    await expect(
      applyCellEditsToXlsx(withDrawing, [], [], [], {
        renames: [],
        additions: [{ name: 'Data Copy', sourceSheetName: 'Data' }],
        removals: [],
        order: ['Data', 'Data Copy', 'Other', 'My Sheet'],
      }),
    ).rejects.toThrow('not supported yet')

    // "Other" carries a sheet-scoped Print_Area (localSheetId=1).
    await expect(
      applyCellEditsToXlsx(await buildSheetsFixture(), [], [], [], {
        renames: [],
        additions: [{ name: 'Other Copy', sourceSheetName: 'Other' }],
        removals: [],
        order: ['Data', 'Other', 'Other Copy', 'My Sheet'],
      }),
    ).rejects.toThrow('sheet-scoped defined names')
  })

  it('clone helpers sanitize part-scoped state', () => {
    expect(
      sanitizeClonedWorksheetXml(
        '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>',
      ),
    ).toBe('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    expect(
      stripPageSetupRelIds('<pageSetup paperSize="9" r:id="rId3" orientation="portrait"/>'),
    ).toBe('<pageSetup paperSize="9" orientation="portrait"/>')
    const hyperlinkOnly =
      '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>'
    expect(prepareClonedSheetRels(hyperlinkOnly, 'Data')).toEqual({
      relsXml: hyperlinkOnly,
      droppedPrinterSettings: false,
    })
    const printerOnly =
      '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/></Relationships>'
    expect(prepareClonedSheetRels(printerOnly, 'Data')).toEqual({
      relsXml: null,
      droppedPrinterSettings: true,
    })
    const commentRels =
      '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>'
    expect(() => prepareClonedSheetRels(commentRels, 'Data')).toThrow(SheetEditError)
  })
})

describe('combined sheet operations', () => {
  it('applies add + rename + remove with a reordered tab list in one save', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildSheetsFixture(),
      [{ sheetName: 'Summary', row: 0, column: 1, writeValue: true, cell: { value: 42 } }],
      [],
      [],
      {
        renames: [{ sheetName: 'Data', newName: 'Numbers' }],
        additions: [{ name: 'Summary' }],
        removals: ['Other'],
        order: ['Summary', 'Numbers', 'My Sheet'],
      },
    )
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const workbook = await entryText(mutation.buffer, 'xl/workbook.xml')
    const sheetsBlock = /<sheets>([\s\S]*?)<\/sheets>/.exec(workbook)?.[1] ?? ''
    expect(sheetsBlock).toBe(
      '<sheet name="Summary" sheetId="4" r:id="rId6"/>' +
        '<sheet name="Numbers" sheetId="1" r:id="rId1"/>' +
        '<sheet name="My Sheet" sheetId="3" r:id="rId3"/>',
    )
    const chart = await entryText(mutation.buffer, 'xl/charts/chart1.xml')
    expect(chart).toContain('<c:f>Numbers!$A$1:$A$2</c:f>')
    const summary = await entryText(mutation.buffer, 'xl/worksheets/sheet4.xml')
    expect(summary).toContain('<c r="B1"><v>42</v></c>')
  })

  it('rejects an order that does not match the final sheet set', async () => {
    await expect(
      applyCellEditsToXlsx(await buildSheetsFixture(), [], [], [], {
        renames: [],
        additions: [],
        removals: ['Other'],
        order: ['Data', 'Other', 'My Sheet'],
      }),
    ).rejects.toThrow('sheet order')
  })
})
