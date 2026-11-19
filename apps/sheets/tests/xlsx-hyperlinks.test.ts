import { describe, expect, it } from 'vitest'

import {
  applyHyperlinkEdits,
  ensureRelationshipNamespace,
} from '@genoffice/xlsx-gateway/gateway/xlsx-hyperlinks'

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

  it('allocates a fresh id when an existing numeric id exceeds the safe integer range', () => {
    const rels = RELS.replace('rId3', 'rId9007199254740992')
    const patch = applyHyperlinkEdits(WORKSHEET, rels, [
      { row: 0, column: 0, target: 'https://new.example' },
    ])
    expect(patch.worksheetXml).toContain('r:id="rId1"')
    expect(patch.relsXml).toContain('Id="rId1"')
    expect(patch.relsXml).toContain('Id="rId9007199254740992"')
  })

  it('expands a self-closing relationships root before adding a link', () => {
    const patch = applyHyperlinkEdits(
      WORKSHEET,
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
      [{ row: 0, column: 0, target: 'https://example.com' }],
    )
    expect(patch.relsXml).toContain('<Relationship Id="rId1"')
    expect(patch.relsXml).toContain('</Relationships>')
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

  it('reclaims and allocates ids in a rels part that uses single quotes', () => {
    // rId1 is taken by a drawing and rId3 by the link being replaced; both are
    // spelled with single quotes, which the id scan used to miss entirely.
    const singleQuoted = RELS.replace(
      '<Relationship Id="rId3"',
      "<Relationship Id='rId1' Type='drawing' Target='../drawings/drawing1.xml'/>" +
        '<Relationship Id="rId3"',
    ).replace(/="(rId\d+|https:[^"]*)"/g, "='$1'")
    const withLink = WORKSHEET.replace(
      '<pageMargins',
      '<hyperlinks><hyperlink ref="A1" r:id="rId3"/></hyperlinks><pageMargins',
    )
    const patch = applyHyperlinkEdits(withLink, singleQuoted, [
      { row: 0, column: 0, target: 'https://new.example' },
    ])

    // The stale rel is dropped and the new one takes the next free id — never
    // rId1 again, which would duplicate the drawing's relationship.
    expect(patch.relsXml).not.toContain('https://old.example')
    expect(patch.worksheetXml).toContain('r:id="rId2"')
    const ids = [...(patch.relsXml ?? '').matchAll(/\bId=["'](rId\d+)["']/g)].map(
      (match) => match[1],
    )
    expect(ids).toEqual(['rId1', 'rId2'])
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
