import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { docxToText } from '../src/docx'
it('reads the styles and notes named by document relationships', async () => {
  const z = new JSZip()
  const w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  z.file(
    'word/document.xml',
    `<w:document ${w}><w:body><w:p><w:pPr><w:pStyle w:val="Custom"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p></w:body></w:document>`,
  )
  z.file(
    'word/_rels/document.xml.rels',
    '<Relationships><Relationship Id="s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="metadata/styles%20part.xml"/><Relationship Id="f" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="../notes/foot.xml"/><Relationship Id="e" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="/notes/end.xml"/></Relationships>',
  )
  z.file(
    'word/metadata/styles part.xml',
    `<w:styles ${w}><w:style w:type="paragraph" w:styleId="Custom"><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style></w:styles>`,
  )
  z.file(
    'notes/foot.xml',
    `<w:footnotes ${w}><w:footnote w:id="1"><w:p><w:r><w:t>Foot note</w:t></w:r></w:p></w:footnote></w:footnotes>`,
  )
  z.file(
    'notes/end.xml',
    `<w:endnotes ${w}><w:endnote w:id="1"><w:p><w:r><w:t>End note</w:t></w:r></w:p></w:endnote></w:endnotes>`,
  )
  expect(await docxToText(await z.generateAsync({ type: 'uint8array' }))).toBe(
    '## Heading\n## Footnotes\nFoot note\n## Endnotes\nEnd note',
  )
})

it.each(['TargetMode="External" Target="notes.xml"', 'Target=""'])(
  'keeps conventional notes when the relationship cannot name a local part: %s',
  async (target) => {
    const zip = new JSZip()
    const w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    zip.file(
      'word/document.xml',
      `<w:document ${w}><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>`,
    )
    zip.file(
      'word/_rels/document.xml.rels',
      `<Relationships><Relationship Id="n" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" ${target}/></Relationships>`,
    )
    zip.file(
      'word/footnotes.xml',
      `<w:footnotes ${w}><w:footnote w:id="1"><w:p><w:r><w:t>Local note</w:t></w:r></w:p></w:footnote></w:footnotes>`,
    )
    zip.file(
      'word/notes.xml',
      `<w:footnotes ${w}><w:footnote w:id="1"><w:p><w:r><w:t>Unrelated part</w:t></w:r></w:p></w:footnote></w:footnotes>`,
    )
    expect(await docxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
      'Body\n## Footnotes\nLocal note',
    )
  },
)
