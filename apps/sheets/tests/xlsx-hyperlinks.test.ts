import { describe, expect, it } from 'vitest'

import { applyHyperlinkEdits, ensureRelationshipNamespace } from '../src/gateway/xlsx-hyperlinks'

const WORKSHEET =
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
  '<pageMargins left="0.7"/></worksheet>'

const RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://old.example" TargetMode="External"/>' +
  '</Relationships>'

describe('applyHyperlinkEdits', () => {
  it('adds an external link, creating the hyperlinks section and the rels file', () => {
    const patch = applyHyperlinkEdits(WORKSHEET, null, [
      { row: 0, column: 0, target: 'https://example.com/a&b' },
    ])
    expect(patch.worksheetXml).toContain(
      '<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks><pageMargins',
    )
    expect(patch.relsChanged).toBe(true)
    expect(patch.relsXml).toContain('Target="https://example.com/a&amp;b" TargetMode="External"')
  })

  it('writes an internal anchor as a location attribute with no rel', () => {
    const patch = applyHyperlinkEdits(WORKSHEET, null, [
      { row: 1, column: 1, target: "#'My Sheet'!B2" },
    ])
    expect(patch.worksheetXml).toContain(`<hyperlink ref="B2" location="'My Sheet'!B2"/>`)
    expect(patch.relsChanged).toBe(false)
    expect(patch.relsXml).toBeNull()
  })

  it('replaces an existing link and drops its now-unused rel', () => {
    const withLink = WORKSHEET.replace(
      '<pageMargins',
      '<hyperlinks><hyperlink ref="A1" r:id="rId3"/></hyperlinks><pageMargins',
    )
    const patch = applyHyperlinkEdits(withLink, RELS, [
      { row: 0, column: 0, target: 'https://new.example' },
    ])
    // The old rel is dropped first, freeing its id space — rId1 is reused.
    expect(patch.worksheetXml).toContain('<hyperlink ref="A1" r:id="rId1"/>')
    expect(patch.relsXml).not.toContain('https://old.example')
    expect(patch.relsXml).toContain('Id="rId1" Type=')
    expect(patch.relsXml).toContain('Target="https://new.example"')
  })

  it('removes a link, its rel, and an emptied hyperlinks section', () => {
    const withLink = WORKSHEET.replace(
      '<pageMargins',
      '<hyperlinks><hyperlink ref="A1" r:id="rId3"/></hyperlinks><pageMargins',
    )
    const patch = applyHyperlinkEdits(withLink, RELS, [{ row: 0, column: 0, target: null }])
    expect(patch.worksheetXml).not.toContain('<hyperlinks')
    expect(patch.relsXml).not.toContain('rId3')
    expect(patch.relsChanged).toBe(true)
  })

  it('keeps a rel still referenced by another hyperlink', () => {
    const withLinks = WORKSHEET.replace(
      '<pageMargins',
      '<hyperlinks><hyperlink ref="A1" r:id="rId3"/><hyperlink ref="B1" r:id="rId3"/></hyperlinks><pageMargins',
    )
    const patch = applyHyperlinkEdits(withLinks, RELS, [{ row: 0, column: 0, target: null }])
    expect(patch.worksheetXml).toContain('<hyperlink ref="B1" r:id="rId3"/>')
    expect(patch.relsXml).toContain('rId3')
  })

  it('appends before </worksheet> when no anchor element exists', () => {
    const bare = '<worksheet><sheetData/></worksheet>'
    const patch = applyHyperlinkEdits(bare, null, [{ row: 0, column: 0, target: '#Data!A1' }])
    expect(patch.worksheetXml).toBe(
      '<worksheet><sheetData/><hyperlinks><hyperlink ref="A1" location="Data!A1"/></hyperlinks></worksheet>',
    )
  })
})

describe('ensureRelationshipNamespace', () => {
  it('injects xmlns:r when the root lacks it and leaves it alone otherwise', () => {
    expect(ensureRelationshipNamespace('<worksheet xmlns="x"><sheetData/></worksheet>')).toContain(
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    )
    expect(ensureRelationshipNamespace(WORKSHEET)).toBe(WORKSHEET)
  })
})
