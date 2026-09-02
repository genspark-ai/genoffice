/**
 * OpenXML SDK-family producers prefix the spreadsheetml namespace
 * (<x:workbook><x:sheets><x:sheet .../>) and use non-numeric relationship
 * ids (R19a46abc...). The regex-based save pipeline could not find any
 * <sheet> element in such a workbook, so every save failed after edits were
 * already applied on screen — the file silently never updated.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx } from '../src/gateway/xlsx-gateway'
import {
  ensureRelationshipNamespace,
  normalizeOoxmlPartPrefix,
} from '../src/gateway/xlsx-namespace'
import { maxRelationshipId, parseSheetElements } from '../src/gateway/xlsx-sheets'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const prefixedWorkbook =
  '<?xml version="1.0" encoding="utf-8"?>' +
  `<x:workbook xmlns:x="${MAIN_NS}">` +
  '<x:sheets>' +
  `<x:sheet name="Sheet1" sheetId="1" r:id="R19a46abc15b54fc2" xmlns:r="${REL_NS}" />` +
  `<x:sheet name="R&#233;sum&#233;" sheetId="2" r:id="Re8b07deca8dc4649" xmlns:r="${REL_NS}" />` +
  '</x:sheets></x:workbook>'

const prefixedWorksheet = (marker: string): string =>
  '<?xml version="1.0" encoding="utf-8"?>' +
  `<x:worksheet xmlns:x="${MAIN_NS}"><x:sheetData>` +
  `<x:row r="1"><x:c r="A1" t="str"><x:v>${marker}</x:v></x:c></x:row>` +
  '</x:sheetData></x:worksheet>'

const workbookRels =
  '﻿<?xml version="1.0" encoding="utf-8"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  `<Relationship Type="${REL_NS}/worksheet" Target="/xl/worksheets/sheet1.xml" Id="R19a46abc15b54fc2" />` +
  `<Relationship Type="${REL_NS}/worksheet" Target="/xl/worksheets/sheet2.xml" Id="Re8b07deca8dc4649" />` +
  '</Relationships>'

async function buildPrefixedFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '﻿<?xml version="1.0" encoding="utf-8"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />' +
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '﻿<?xml version="1.0" encoding="utf-8"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Type="${REL_NS}/officeDocument" Target="/xl/workbook.xml" Id="Rf0e3ab4bc7f54e0e" />` +
      '</Relationships>',
  )
  zip.file('xl/workbook.xml', prefixedWorkbook)
  zip.file('xl/_rels/workbook.xml.rels', workbookRels)
  zip.file('xl/worksheets/sheet1.xml', prefixedWorksheet('one'))
  zip.file('xl/worksheets/sheet2.xml', prefixedWorksheet('two'))
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

describe('normalizeOoxmlPartPrefix', () => {
  it('rebinds a prefixed workbook part to the default namespace', () => {
    const normalized = normalizeOoxmlPartPrefix(prefixedWorkbook)
    expect(normalized).toContain(`<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">`)
    expect(normalized).toContain('</workbook>')
    expect(normalized).not.toContain('<x:')
    expect(normalized).toContain('r:id="R19a46abc15b54fc2"')

    const sheets = parseSheetElements(normalized)
    expect(sheets.map((sheet) => sheet.name)).toEqual(['Sheet1', 'Résumé'])
    expect(sheets.map((sheet) => sheet.relationshipId)).toEqual([
      'R19a46abc15b54fc2',
      'Re8b07deca8dc4649',
    ])
  })

  it('returns unprefixed parts unchanged', () => {
    const xml = `<?xml version="1.0"?><workbook xmlns="${MAIN_NS}"><sheets/></workbook>`
    expect(normalizeOoxmlPartPrefix(xml)).toBe(xml)
  })

  it('leaves parts of other namespaces alone', () => {
    const chart =
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart/></c:chartSpace>'
    expect(normalizeOoxmlPartPrefix(chart)).toBe(chart)
  })

  it('bails out when a default namespace is already bound', () => {
    const xml = `<x:workbook xmlns:x="${MAIN_NS}" xmlns="urn:other">` + '<x:sheets/></x:workbook>'
    expect(normalizeOoxmlPartPrefix(xml)).toBe(xml)
  })

  it('bails out when the prefix is rebound deeper in the tree', () => {
    const xml =
      `<x:workbook xmlns:x="${MAIN_NS}"><x:sheets>` +
      '<x:ext xmlns:x="urn:other"/></x:sheets></x:workbook>'
    expect(normalizeOoxmlPartPrefix(xml)).toBe(xml)
  })

  it('keeps the prefix binding when attributes still use it', () => {
    const xml = `<x:worksheet xmlns:x="${MAIN_NS}">` + '<x:oleObject x:custom="1"/></x:worksheet>'
    const normalized = normalizeOoxmlPartPrefix(xml)
    expect(normalized).toContain(`xmlns="${MAIN_NS}"`)
    expect(normalized).toContain(`xmlns:x="${MAIN_NS}"`)
    expect(normalized).toContain('<oleObject x:custom="1"/>')
  })
})

describe('non-numeric relationship ids', () => {
  it('maxRelationshipId ignores them so new ids start fresh without colliding', () => {
    expect(maxRelationshipId(workbookRels)).toBe(0)
    expect(workbookRels).not.toContain('Id="rId1"')
  })
})

describe('saving a workbook with prefixed parts', () => {
  it('writes a cell edit through the prefixed workbook and worksheet', async () => {
    const mutation = await applyCellEditsToXlsx(await buildPrefixedFixture(), [
      { sheetName: 'Sheet1', row: 0, column: 1, writeValue: true, cell: { value: 42 } },
    ])
    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet = (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''
    expect(sheet).toContain(`<worksheet xmlns="${MAIN_NS}"`)
    expect(sheet).toContain('<v>one</v>')
    expect(/<c r="B1"[^>]*><v>42<\/v><\/c>/.test(sheet)).toBe(true)
    const workbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
    expect(workbook).toContain('fullCalcOnLoad="1"')
    // Untouched siblings keep their original bytes.
    const sibling = (await zip.file('xl/worksheets/sheet2.xml')?.async('text')) ?? ''
    expect(sibling).toBe(prefixedWorksheet('two'))
  })

  it('adds a sheet with a fresh rId and a root-level r binding', async () => {
    const mutation = await applyCellEditsToXlsx(await buildPrefixedFixture(), [], [], [], {
      renames: [],
      additions: [{ name: 'Added' }],
      removals: [],
      order: ['Sheet1', 'Résumé', 'Added'],
    })
    const zip = await JSZip.loadAsync(mutation.buffer)
    const workbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
    expect(workbook).toContain(`<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">`)
    expect(workbook).toContain('<sheet name="Added" sheetId="3" r:id="rId1"/>')
    const rels = (await zip.file('xl/_rels/workbook.xml.rels')?.async('text')) ?? ''
    expect(rels.match(/Id="rId1"/g)).toHaveLength(1)
    expect(rels).toContain('Target="worksheets/sheet3.xml"')
    expect(await zip.file('xl/worksheets/sheet3.xml')?.async('text')).toContain('<sheetData/>')
  })
})

/// prod_038 shape: default main namespace on the root, but xmlns:r declared
/// on each <sheet> element — a new <sheet r:id=.../> appended without its
/// own binding made the saved workbook.xml ill-formed.
const perElementRWorkbook =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<workbook xmlns="${MAIN_NS}"><sheets>` +
  `<sheet xmlns:r="${REL_NS}" name="Data" sheetId="1" state="visible" r:id="rId1"/>` +
  '</sheets></workbook>'

async function buildPerElementRFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>` +
      '</Relationships>',
  )
  zip.file('xl/workbook.xml', perElementRWorkbook)
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>` +
      '</Relationships>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<worksheet xmlns="${MAIN_NS}"><sheetData/></worksheet>`,
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

describe('ensureRelationshipNamespace', () => {
  it('binds r on a root that lacks it and leaves bound roots alone', () => {
    const bound = ensureRelationshipNamespace(`<workbook xmlns="${MAIN_NS}"><sheets/></workbook>`)
    expect(bound).toContain(`<workbook xmlns:r="${REL_NS}" xmlns="${MAIN_NS}">`)
    const already = `<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheets/></workbook>`
    expect(ensureRelationshipNamespace(already)).toBe(already)
  })

  it('reads the root tag quote-aware', () => {
    const xml = `<workbook xmlns="${MAIN_NS}" title="a > b"><sheets/></workbook>`
    expect(ensureRelationshipNamespace(xml)).toContain(`title="a > b"><sheets/>`)
    expect(ensureRelationshipNamespace(xml)).toContain('xmlns:r=')
  })
})

describe('adding a sheet when r is bound per element only', () => {
  it('hoists a root-level r binding so the new element stays well-formed', async () => {
    const mutation = await applyCellEditsToXlsx(await buildPerElementRFixture(), [], [], [], {
      renames: [],
      additions: [{ name: 'MergeOut' }],
      removals: [],
      order: ['Data', 'MergeOut'],
    })
    const zip = await JSZip.loadAsync(mutation.buffer)
    const savedWorkbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
    expect(savedWorkbook).toContain(`<workbook xmlns:r="${REL_NS}" xmlns="${MAIN_NS}">`)
    expect(savedWorkbook).toContain('<sheet name="MergeOut" sheetId="2" r:id="rId2"/>')
  })
})
