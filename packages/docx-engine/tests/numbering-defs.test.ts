import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const MULTILEVEL_NUMBERING =
  XML_DECL +
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:abstractNum w:abstractNumId="0">' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>' +
  '<w:lvl w:ilvl="2"><w:start w:val="3"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%3)"/></w:lvl>' +
  '</w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '<w:num w:numId="2"><w:abstractNumId w:val="0"/>' +
  '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>' +
  '<w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="5"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%2:"/></w:lvl></w:lvlOverride>' +
  '</w:num>' +
  '</w:numbering>'

describe('numbering definitions (word/numbering.xml)', () => {
  it('parses per-level numFmt / lvlText / start and lvlOverrides', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      numberingXml: MULTILEVEL_NUMBERING,
    })
    const doc = await parseDocx(bytes)

    const num1 = doc.numbering.get('1')!
    expect(num1.abstractNumId).toBe('0')
    expect(num1.levels[0]).toEqual({ numFmt: 'decimal', lvlText: '%1.', start: 1 })
    expect(num1.levels[1]).toEqual({ numFmt: 'decimal', lvlText: '%1.%2', start: 1 })
    expect(num1.levels[2]).toEqual({ numFmt: 'lowerLetter', lvlText: '%3)', start: 3 })
    expect(num1.startOverrides).toEqual({})

    const num2 = doc.numbering.get('2')!
    expect(num2.startOverrides).toEqual({ 0: 1 })
    // A full w:lvl inside lvlOverride replaces that level's definition; other levels
    // follow the abstractNum
    expect(num2.levels[1]).toEqual({ numFmt: 'upperRoman', lvlText: '%2:', start: 5 })
    expect(num2.levels[0]).toEqual({ numFmt: 'decimal', lvlText: '%1.', start: 1 })
  })

  it('keeps the bullet/ordered classification for list blocks', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>a</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>b</w:t></w:r></w:p>',
      withNumbering: true,
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].list?.kind).toBe('bullet')
    expect(doc.blocks[1].list?.kind).toBe('ordered')
    expect(doc.numbering.get('2')?.levels[0]?.lvlText).toBe('%1.')
  })
})

describe('mixed multilevel list kind', () => {
  const MIXED_NUMBERING =
    XML_DECL +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="9">' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#61623;"/></w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="7"><w:abstractNumId w:val="9"/></w:num>' +
    '</w:numbering>'

  it('classifies each level by its own numFmt, not level 0', async () => {
    const li = (ilvl: number, text: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="7"/></w:numPr></w:pPr>` +
      `<w:r><w:t>${text}</w:t></w:r></w:p>`
    const doc = await parseDocx(
      await buildDocx({ bodyXml: li(0, 'top') + li(1, 'sub'), numberingXml: MIXED_NUMBERING }),
    )
    expect(doc.blocks[0].list).toMatchObject({ kind: 'ordered', ilvl: 0 })
    expect(doc.blocks[1].list).toMatchObject({ kind: 'bullet', ilvl: 1 })
  })

  it('falls back to the level-0 classification for undefined levels', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:numPr><w:ilvl w:val="4"/><w:numId w:val="7"/></w:numPr></w:pPr>' +
          '<w:r><w:t>deep</w:t></w:r></w:p>',
        numberingXml: MIXED_NUMBERING,
      }),
    )
    expect(doc.blocks[0].list).toMatchObject({ kind: 'ordered', ilvl: 4 })
  })
})
