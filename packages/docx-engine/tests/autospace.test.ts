import { describe, expect, it } from 'vitest'
import { mergePPrFormat, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const OFF = '<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>'

const TIGHT_STYLES =
  '<w:style w:type="paragraph" w:styleId="Tight"><w:name w:val="Tight"/>' +
  `<w:basedOn w:val="Normal"/><w:pPr>${OFF}</w:pPr></w:style>` +
  '<w:style w:type="paragraph" w:styleId="TightChild"><w:name w:val="Tight Child"/>' +
  '<w:basedOn w:val="Tight"/></w:style>'

async function parseFirst(pPrXml: string, extraStylesXml?: string) {
  const bytes = await buildDocx({
    bodyXml: `<w:p>${pPrXml ? `<w:pPr>${pPrXml}</w:pPr>` : ''}<w:r><w:t>x</w:t></w:r></w:p>`,
    extraStylesXml,
  })
  return (await parseDocx(bytes)).blocks[0]
}

describe('w:autoSpaceDE/DN parsing', () => {
  it('absent means Word default on: no autoSpace field', async () => {
    const block = await parseFirst('<w:jc w:val="center"/>')
    expect(block.format?.autoSpace).toBeUndefined()
  })

  it('both explicitly off parse to autoSpace false', async () => {
    const block = await parseFirst(OFF)
    expect(block.format?.autoSpace).toBe(false)
  })

  it('only DE off keeps the digit gap, only DN off keeps the letter gap', async () => {
    expect((await parseFirst('<w:autoSpaceDE w:val="0"/>')).format?.autoSpace).toBe('digits')
    expect((await parseFirst('<w:autoSpaceDN w:val="0"/>')).format?.autoSpace).toBe('latin')
  })

  it('one switch off on the paragraph merges with the other off in the style', async () => {
    const styles =
      '<w:style w:type="paragraph" w:styleId="NoDigits"><w:name w:val="No Digits"/>' +
      '<w:basedOn w:val="Normal"/><w:pPr><w:autoSpaceDN w:val="0"/></w:pPr></w:style>'
    const block = await parseFirst('<w:pStyle w:val="NoDigits"/><w:autoSpaceDE w:val="0"/>', styles)
    expect(block.format?.autoSpace).toBe(false)
    expect((await parseFirst('<w:pStyle w:val="NoDigits"/>', styles)).format?.autoSpace).toBe(
      'latin',
    )
  })

  it('unchanged paragraphs keep the raw autoSpace bytes on merge', async () => {
    const block = await parseFirst(OFF)
    expect(mergePPrFormat(block.rawPPr!, block.format)).toBe(`<w:pPr>${OFF}</w:pPr>`)
  })
})

describe('style-chain inheritance', () => {
  it('a style turning autoSpace off reaches its paragraphs', async () => {
    const block = await parseFirst('<w:pStyle w:val="Tight"/>', TIGHT_STYLES)
    expect(block.format?.autoSpace).toBe(false)
  })

  it('inherits through the basedOn chain', async () => {
    const block = await parseFirst('<w:pStyle w:val="TightChild"/>', TIGHT_STYLES)
    expect(block.format?.autoSpace).toBe(false)
  })

  it('explicit on in the pPr overrides a style-level off', async () => {
    const block = await parseFirst(
      '<w:pStyle w:val="Tight"/><w:autoSpaceDE/><w:autoSpaceDN/>',
      TIGHT_STYLES,
    )
    expect(block.format?.autoSpace).toBe(true)
  })
})
