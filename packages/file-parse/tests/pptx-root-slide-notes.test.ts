import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { pptxToText } from '../src/pptx'
it.each(['slide.xml', 'pages/slide.xml'])('reads notes for a slide at %s', async (path) => {
  const zip = new JSZip()
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
  zip.file(
    'ppt/presentation.xml',
    '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="' +
      rel.slice(0, -1) +
      '"><p:sldIdLst><p:sldId id="256" r:id="s"/></p:sldIdLst></p:presentation>',
  )
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    '<Relationships><Relationship Id="s" Type="' +
      rel +
      'slide" Target="/' +
      path +
      '"/></Relationships>',
  )
  zip.file(
    path,
    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Body</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
  )
  zip.file(
    path.replace(/([^/]+)$/, '_rels/$1.rels'),
    '<Relationships><Relationship Id="n" Type="' +
      rel +
      'notesSlide" Target="/notes.xml"/></Relationships>',
  )
  zip.file(
    'notes.xml',
    '<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Speaker notes</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>',
  )
  expect(await pptxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
    '## Slide 1\nBody\n### Notes\nSpeaker notes',
  )
})
