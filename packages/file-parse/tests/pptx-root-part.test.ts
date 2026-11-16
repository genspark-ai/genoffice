import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { pptxToText } from '../src/pptx'
it('uses the package relationship to locate a presentation outside ppt', async () => {
  const zip = new JSZip()
  zip.file(
    '_rels/.rels',
    '<Relationships><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/deck.xml"/></Relationships>',
  )
  zip.file(
    'custom/deck.xml',
    '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="s"/></p:sldIdLst></p:presentation>',
  )
  zip.file(
    'custom/_rels/deck.xml.rels',
    '<Relationships><Relationship Id="s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../pages/first.xml"/></Relationships>',
  )
  zip.file(
    'pages/first.xml',
    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Relocated slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
  )
  expect(await pptxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
    '## Slide 1\nRelocated slide',
  )
})
