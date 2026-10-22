import { describe, expect, it } from 'vitest'

import { parseDocx } from '../src/index'
import { onOffTagIn } from '../src/parse-xml-text'
import { parseNotesXml } from '../src/notes'
import { standaloneNoteParagraphReader } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

function footnotesXml(rPr: string): string {
  return (
    XML_DECL +
    '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r>' +
    `<w:r>${rPr}<w:t>note body</w:t></w:r>` +
    '</w:p></w:footnote></w:footnotes>'
  )
}

function bodyRun(rPr: string): { bold?: boolean; text?: string } | undefined {
  const notes = parseNotesXml(footnotesXml(rPr), 'footnote', standaloneNoteParagraphReader())
  return notes
    .flatMap((note) => note.richParas ?? [])
    .flat()
    .find((r) => r.text === 'note body')
}

const tocEntry = (rPr: string): string =>
  '<w:p><w:pPr><w:pStyle w:val="TOC2"/><w:tabs><w:tab w:val="right" w:pos="8786"/></w:tabs>' +
  `<w:rPr><w:sz w:val="24"/></w:rPr></w:pPr><w:hyperlink w:anchor="_Toc1">` +
  `<w:r>${rPr}<w:t>Annexe 1-1 : Classification</w:t></w:r>` +
  '<w:r><w:tab/></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>6</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '</w:hyperlink></w:p>'

const tocBold = async (rPr: string): Promise<boolean | undefined> => {
  const doc = await parseDocx(await buildDocx({ bodyXml: tocEntry(rPr) }))
  return doc.blocks[0]?.fieldDisplay?.bold
}

describe('onOffTagIn', () => {
  it('reads the toggle off the start tag, so both spellings count', () => {
    expect(onOffTagIn('<w:rPr><w:b/></w:rPr>', 'w:b')).toBe(true)
    expect(onOffTagIn('<w:rPr><w:b></w:b></w:rPr>', 'w:b')).toBe(true)
    expect(onOffTagIn('<w:rPr><w:b w:val="true"/></w:rPr>', 'w:b')).toBe(true)
    expect(onOffTagIn('<w:rPr><w:b w:val="true"></w:b></w:rPr>', 'w:b')).toBe(true)
  })

  it('honours an explicit off value in either quote style', () => {
    // ST_OnOff is 0/1/true/false/on/off, case-insensitive, and XML allows
    // either quote character. The old per-site regexes only recognised a
    // double-quoted 0/false/none/off, so these read as SET.
    for (const val of ['0', 'false', 'none', 'off', 'OFF', 'False']) {
      expect(onOffTagIn(`<w:rPr><w:b w:val="${val}"/></w:rPr>`, 'w:b'), val).toBe(false)
      expect(onOffTagIn(`<w:rPr><w:b w:val="${val}"></w:b></w:rPr>`, 'w:b'), val).toBe(false)
    }
    expect(onOffTagIn("<w:rPr><w:b w:val='0'/></w:rPr>", 'w:b')).toBe(false)
    expect(onOffTagIn("<w:rPr><w:b w:val='off'></w:b></w:rPr>", 'w:b')).toBe(false)
  })

  it('does not match a longer element that starts with the same prefix', () => {
    // w:b must not read w:bCs, or a complex-script flag would turn bold on
    expect(onOffTagIn('<w:rPr><w:bCs/></w:rPr>', 'w:b')).toBeUndefined()
    expect(onOffTagIn('<w:rPr><w:bCs w:val="false"/></w:rPr>', 'w:b')).toBeUndefined()
    expect(onOffTagIn('<w:rPr><w:smallCaps/></w:rPr>', 'w:b')).toBeUndefined()
  })

  it('reports an absent element as undefined, not as false', () => {
    expect(onOffTagIn('<w:rPr></w:rPr>', 'w:b')).toBeUndefined()
    expect(onOffTagIn('', 'w:b')).toBeUndefined()
  })

  it('stays case-insensitive, as the regexes it replaces were', () => {
    // the old per-site regexes carried the i flag, so an upper-case spelling
    // that is not valid OOXML but appears in the wild still read as set
    expect(onOffTagIn('<w:rPr><W:B/></w:rPr>', 'w:b')).toBe(true)
    expect(onOffTagIn('<w:rPr><W:B></W:B></w:rPr>', 'w:b')).toBe(true)
    expect(onOffTagIn('<w:rPr><w:b W:VAL="0"/></w:rPr>', 'w:b')).toBe(false)
  })
})

describe('footnote run flags read both spellings', () => {
  it('keeps bold on a run whose w:b is written as an element pair', () => {
    // the paired form of the same CT_OnOff element used to read as "not set",
    // so a bold footnote run silently lost its weight
    expect(bodyRun('<w:rPr><w:b></w:b></w:rPr>')?.bold).toBe(true)
    expect(bodyRun('<w:rPr><w:b w:val="true"></w:b></w:rPr>')?.bold).toBe(true)
  })

  it('still reads the self-closing form, and still honours an explicit off', () => {
    expect(bodyRun('<w:rPr><w:b/></w:rPr>')?.bold).toBe(true)
    expect(bodyRun('<w:rPr><w:b w:val="0"/></w:rPr>')?.bold).toBeUndefined()
    expect(bodyRun('<w:rPr><w:b w:val="off"></w:b></w:rPr>')?.bold).toBeUndefined()
  })
})

describe('TOC entry weight reads both spellings', () => {
  it('keeps a TOC entry bold when the leading run writes w:b as a pair', async () => {
    expect(await tocBold('<w:rPr><w:b></w:b></w:rPr>')).toBe(true)
  })

  it('still reads the self-closing form, and still honours an explicit off', async () => {
    expect(await tocBold('<w:rPr><w:b/></w:rPr>')).toBe(true)
    expect(await tocBold('<w:rPr><w:b w:val="0"/></w:rPr>')).toBeUndefined()
    // the old lookahead only understood a double-quoted value, so this read bold
    expect(await tocBold("<w:rPr><w:b w:val='0'/></w:rPr>")).toBeUndefined()
  })
})
