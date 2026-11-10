import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const STYLES_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

/** Normal with hanging line-end punctuation switched off (Word: Asian typography, allow overflow) */
const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${STYLES_NS}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>' +
  '<w:pPr><w:overflowPunct w:val="0"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="a"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Loose"><w:name w:val="Loose"/></w:style>' +
  '</w:styles>'

async function parseFirst(pPrXml: string, stylesXml?: string) {
  const bytes = await buildDocx({
    bodyXml: `<w:p>${pPrXml ? `<w:pPr>${pPrXml}</w:pPr>` : ''}<w:r><w:t>相違ありません。</w:t></w:r></w:p>`,
    stylesXml,
  })
  return (await parseDocx(bytes)).blocks[0]
}

describe('w:overflowPunct parsing', () => {
  it('absent means Word default (punctuation may hang): no field', async () => {
    expect((await parseFirst('')).format?.overflowPunct).toBeUndefined()
  })

  it('w:val="0" parses to false, bare element to true', async () => {
    expect((await parseFirst('<w:overflowPunct w:val="0"/>')).format?.overflowPunct).toBe(false)
    expect((await parseFirst('<w:overflowPunct/>')).format?.overflowPunct).toBe(true)
  })

  it('reaches paragraphs through the default style and basedOn, not unrelated styles', async () => {
    expect((await parseFirst('', STYLES)).format?.overflowPunct).toBe(false)
    expect((await parseFirst('<w:pStyle w:val="Body"/>', STYLES)).format?.overflowPunct).toBe(false)
    expect(
      (await parseFirst('<w:pStyle w:val="Loose"/>', STYLES)).format?.overflowPunct,
    ).toBeUndefined()
  })
})
