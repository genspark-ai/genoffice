import { describe, expect, it } from 'vitest'

import { acceptRevisionsInPartXml, acceptRevisionsInSideParts } from '../src/revision-parts'
import { buildDocx } from './helpers/build-docx'

const AT = 'w:author="X" w:date="2026-01-01T00:00:00Z"'
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`

describe('accepting tracked changes in a side part', () => {
  it('keeps inserted text, drops deleted text, and counts both', () => {
    const xml = `<w:p>${run('kept ')}<w:ins w:id="1" ${AT}>${run('added ')}</w:ins><w:del w:id="2" ${AT}><w:r><w:delText>gone</w:delText></w:r></w:del>${run('tail')}</w:p>`
    const result = acceptRevisionsInPartXml(xml)
    expect(result.xml).toBe(`<w:p>${run('kept ')}${run('added ')}${run('tail')}</w:p>`)
    expect(result.accepted).toBe(2)
    expect(result.unhandledParaMarks).toBe(0)
  })

  it('resolves a move the way Word does: destination stays, source goes, markers vanish', () => {
    const xml =
      `<w:p>${run('P1 ')}<w:moveFromRangeStart w:id="1" ${AT} w:name="m"/>` +
      `<w:moveFrom w:id="2" ${AT}>${run('MOVED ')}</w:moveFrom><w:moveFromRangeEnd w:id="1"/>${run('beta')}</w:p>` +
      `<w:p>${run('P2 ')}<w:moveToRangeStart w:id="3" ${AT} w:name="m"/>` +
      `<w:moveTo w:id="4" ${AT}>${run('MOVED ')}</w:moveTo><w:moveToRangeEnd w:id="3"/>${run('delta')}</w:p>`
    const result = acceptRevisionsInPartXml(xml)
    expect(result.xml).toBe(
      `<w:p>${run('P1 ')}${run('beta')}</w:p><w:p>${run('P2 ')}${run('MOVED ')}${run('delta')}</w:p>`,
    )
    expect(result.accepted).toBe(2)
  })

  it('drops formatting-change records so the current format stands', () => {
    const xml = `<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="1" ${AT}><w:pPr/></w:pPrChange></w:pPr><w:r><w:rPr><w:b/><w:rPrChange w:id="2" ${AT}><w:rPr/></w:rPrChange></w:rPr><w:t>bold now</w:t></w:r></w:p>`
    const result = acceptRevisionsInPartXml(xml)
    expect(result.xml).toBe(
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>bold now</w:t></w:r></w:p>',
    )
    expect(result.accepted).toBe(2)
  })

  it('accepts an inserted paragraph mark but reports a deleted one it cannot merge', () => {
    const xml =
      `<w:p><w:pPr><w:rPr><w:ins w:id="1" ${AT}/></w:rPr></w:pPr>${run('new mark')}</w:p>` +
      `<w:p><w:pPr><w:rPr><w:del w:id="2" ${AT}/></w:rPr></w:pPr>${run('merges in Word')}</w:p>` +
      `<w:p>${run('next')}</w:p>`
    const result = acceptRevisionsInPartXml(xml)
    // the inserted mark is permanent now; the deleted one is untouched and counted
    expect(result.xml).toContain(`<w:p><w:pPr><w:rPr></w:rPr></w:pPr>${run('new mark')}</w:p>`)
    expect(result.xml).toContain(`<w:del w:id="2" ${AT}/>`)
    expect(result.accepted).toBe(1)
    expect(result.unhandledParaMarks).toBe(1)
  })
})

describe('accept-all across the parts the editor never opens', () => {
  const HEADER_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:p>${run('Header ')}<w:ins w:id="1" ${AT}>${run('added')}</w:ins></w:p></w:hdr>`
  const FOOTER_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:p>${run('Footer ')}<w:del w:id="2" ${AT}><w:r><w:delText>old</w:delText></w:r></w:del></w:p></w:ftr>`
  const CLEAN_FOOTER_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:p>${run('plain')}</w:p></w:ftr>`

  it('rewrites only the parts that carried a revision', async () => {
    const hf = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'
    const bytes = await buildDocx({
      bodyXml: `<w:p>${run('body')}</w:p>`,
      extraParts: [
        { path: 'word/header1.xml', xml: HEADER_XML, contentType: `${hf}header+xml` },
        { path: 'word/footer1.xml', xml: FOOTER_XML, contentType: `${hf}footer+xml` },
        { path: 'word/footer2.xml', xml: CLEAN_FOOTER_XML, contentType: `${hf}footer+xml` },
      ],
    })
    const result = await acceptRevisionsInSideParts(bytes)
    expect(Object.keys(result.parts).sort()).toEqual(['word/footer1.xml', 'word/header1.xml'])
    expect(result.parts['word/header1.xml']).toContain(`${run('Header ')}${run('added')}`)
    expect(result.parts['word/header1.xml']).not.toContain('<w:ins')
    expect(result.parts['word/footer1.xml']).not.toContain('<w:del')
    expect(result.accepted).toBe(2)
    expect(result.unhandledParaMarks).toBe(0)
  })

  it('is a no-op on a document whose side parts are clean', async () => {
    const bytes = await buildDocx({ bodyXml: `<w:p>${run('body')}</w:p>` })
    const result = await acceptRevisionsInSideParts(bytes)
    expect(result.parts).toEqual({})
    expect(result.accepted).toBe(0)
  })
})
