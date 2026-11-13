import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { pptxToText } from '../src/pptx'
it.each([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
])('extracts notes through %s relationships', async (ns) => {
  const z = new JSZip()
  z.file(
    'ppt/slides/slide1.xml',
    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
  )
  z.file(
    'ppt/slides/_rels/slide1.xml.rels',
    `<Relationships><Relationship Id="n" Type="${ns}/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>`,
  )
  z.file(
    'ppt/notesSlides/notesSlide1.xml',
    '<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Speaker note</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>',
  )
  expect(await pptxToText(await z.generateAsync({ type: 'uint8array' }))).toContain(
    '### Notes\nSpeaker note',
  )
})
