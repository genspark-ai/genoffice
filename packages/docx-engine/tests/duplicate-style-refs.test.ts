/**
 * Repeated w:pStyle / w:rStyle children of one property set: Word applies the
 * last one. Some Markdown converters write `Heading1` + `MdHeading1`; only the
 * second carries the bold / size / spacing the headings are laid out with.
 */
import { describe, expect, it } from 'vitest'
import { dedupeStyleRefs, hasDuplicateStyleRefs } from '../src/ooxml-normalize'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const extraStyles =
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:sz w:val="32"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="MdHeading1"><w:name w:val="MdHeading1"/><w:pPr><w:spacing w:before="480"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="Red"><w:name w:val="Red"/><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="Blue"><w:name w:val="Blue"/><w:rPr><w:color w:val="0000FF"/></w:rPr></w:style>'

describe('duplicate style references', () => {
  it('collapses consecutive pStyle / rStyle siblings to the last one', () => {
    const xml =
      '<w:pPr><w:pStyle w:val="A"/> <w:pStyle w:val="B"/><w:pStyle w:val="C"/><w:rPr><w:rStyle w:val="X"/><w:rStyle w:val="Y"/></w:rPr></w:pPr>'
    expect(hasDuplicateStyleRefs(xml)).toBe(true)
    expect(dedupeStyleRefs(xml)).toBe(
      '<w:pPr><w:pStyle w:val="C"/><w:rPr><w:rStyle w:val="Y"/></w:rPr></w:pPr>',
    )
  })

  it('leaves single references and non-adjacent ones alone', () => {
    const xml = '<w:pPr><w:pStyle w:val="A"/><w:rPr><w:rStyle w:val="X"/></w:rPr></w:pPr>'
    expect(hasDuplicateStyleRefs(xml)).toBe(false)
    expect(dedupeStyleRefs(xml)).toBe(xml)
    const tracked =
      '<w:pPr><w:pStyle w:val="A"/><w:pPrChange w:id="1"><w:pPr><w:pStyle w:val="B"/></w:pPr></w:pPrChange></w:pPr>'
    expect(hasDuplicateStyleRefs(tracked)).toBe(false)
  })

  it('the last pStyle / rStyle wins through parseDocx', async () => {
    const body =
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:pStyle w:val="MdHeading1"/></w:pPr>' +
      '<w:r><w:rPr><w:rStyle w:val="Red"/><w:rStyle w:val="Blue"/></w:rPr><w:t>Title</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: body, extraStylesXml: extraStyles }))
    const block = doc.blocks[0]!
    expect(block.styleId).toBe('MdHeading1')
    expect(block.format?.effective?.spacing.before).toBe(480)
    const run = block.runs![0]!
    expect(run.styleId).toBe('Blue')
    expect(run.effective).toMatchObject({ bold: true, sizeHalfPoints: 36, color: '0000FF' })
    expect(doc.internal.documentXml).not.toContain('w:val="Heading1"')
  })
})
