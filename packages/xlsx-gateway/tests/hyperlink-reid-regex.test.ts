import { describe, expect, it } from 'vitest'
import { applyHyperlinkEdits } from '../src/gateway/xlsx-hyperlinks'

const RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>' +
  '<Relationship Id=".*" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>' +
  '</Relationships>'

function sheetWith(reference: string): string {
  return (
    '<worksheet><hyperlinks><hyperlink ref="A1" r:id="' +
    reference +
    '"/></hyperlinks><sheetData/></worksheet>'
  )
}

describe('xlsx-hyperlink relationship reclaim', () => {
  it('does not delete unrelated relationships when the reclaimed r:id holds regex metacharacters', () => {
    const patch = applyHyperlinkEdits(sheetWith('.*'), RELS, [
      { row: 0, column: 0, target: 'https://example.com/next' },
    ])

    expect(patch.relsChanged).toBe(true)
    // The drawing and table relationships must survive the reclaim.
    expect(patch.relsXml).toContain('relationships/drawing')
    expect(patch.relsXml).toContain('relationships/table')
    // The reclaimed hyperlink relationship is gone.
    expect(patch.relsXml).not.toContain('Id=".*"')
    // Exactly one hyperlink relationship remains: the newly allocated one.
    expect(patch.relsXml!.match(/relationships\/hyperlink/g) ?? []).toHaveLength(1)
  })

  it('still reclaims a well-formed r:id', () => {
    const rels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>' +
      '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://old.example" TargetMode="External"/>' +
      '</Relationships>'
    const patch = applyHyperlinkEdits(sheetWith('rId7'), rels, [
      { row: 0, column: 0, target: 'https://new.example' },
    ])
    expect(patch.relsChanged).toBe(true)
    expect(patch.relsXml).toContain('Id="rId1"')
    expect(patch.relsXml).not.toContain('https://old.example')
    expect(patch.relsXml).toContain('https://new.example')
  })
})
