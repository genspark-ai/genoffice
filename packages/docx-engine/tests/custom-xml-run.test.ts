import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type Block } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`
const custom = (inner: string) =>
  '<w:customXml w:uri="urn:example" w:element="party">' +
  '<w:customXmlPr><w:attr w:name="source" w:val="metadata-only"/></w:customXmlPr>' +
  inner +
  '</w:customXml>'
const paragraph = (inner: string) => `<w:p>${run('before')}${inner}${run('after')}</w:p>`
const textOf = (block: Block) => block.runs?.map((part) => part.text).join('')
const bytesOf = () => buildDocx({ bodyXml: paragraph(custom(run('wrapped'))) })

async function editNeighbor(bytes: Uint8Array) {
  const parsed = await parseDocx(bytes)
  const runs = parsed.blocks[0].runs!.map((part) => ({
    ...part,
    text: part.text.replace(/after$/, 'AFTER'),
  }))
  const saved = await saveDocx(parsed, [{ kind: 'generated', block: { type: 'paragraph', runs } }])
  return { parsed, saved, reopened: await parseDocx(saved) }
}

describe('run-level custom XML visible content', () => {
  it('imports visible runs in their original order', async () => {
    const parsed = await parseDocx(await bytesOf())
    expect(textOf(parsed.blocks[0])).toBe('beforewrappedafter')
    expect(textOf(parsed.blocks[0])).not.toContain('metadata-only')
  })

  it('keeps the wrapped text through a neighboring edit, save and reopen', async () => {
    const { saved, reopened } = await editNeighbor(await bytesOf())
    const zip = await JSZip.loadAsync(saved)
    expect(textOf(reopened.blocks[0])).toBe('beforewrappedAFTER')
    expect(await zip.file('word/document.xml')!.async('string')).toContain('wrapped')
  })

  it('leaves an original paragraph and its custom XML metadata byte-identical', async () => {
    const bytes = await bytesOf()
    const parsed = await parseDocx(bytes)
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: parsed.blocks[0].docxIndex! },
    ])
    const before = await JSZip.loadAsync(bytes)
    const after = await JSZip.loadAsync(saved)
    expect(await after.file('word/document.xml')!.async('uint8array')).toEqual(
      await before.file('word/document.xml')!.async('uint8array'),
    )
  })

  it.each([
    [
      'inside a hyperlink',
      `<w:hyperlink r:id="rIdLink">${custom('<w:r><w:rPr><w:b/></w:rPr><w:t>wrapped</w:t></w:r>')}</w:hyperlink>`,
    ],
    [
      'around a hyperlink',
      custom(
        '<w:hyperlink r:id="rIdLink"><w:r><w:rPr><w:b/></w:rPr><w:t>wrapped</w:t></w:r></w:hyperlink>',
      ),
    ],
  ])('carries formatting and hyperlink context %s', async (_, inner) => {
    const bytes = await buildDocx({
      bodyXml: paragraph(inner),
      extraRels:
        '<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/party" TargetMode="External"/>',
    })
    const { parsed, reopened } = await editNeighbor(bytes)
    for (const doc of [parsed, reopened]) {
      const wrapped = doc.blocks[0].runs!.find((part) => part.text === 'wrapped')
      expect(wrapped).toMatchObject({ bold: true, link: { href: 'https://example.com/party' } })
    }
    expect(textOf(reopened.blocks[0])).toBe('beforewrappedAFTER')
  })

  it('recurses into nested custom XML and smart tags without duplicating text', async () => {
    const inner = custom(
      run('one') +
        '<w:smartTag w:uri="urn:example" w:element="tag">' +
        custom(run('two')) +
        '</w:smartTag>' +
        run('three'),
    )
    const { reopened } = await editNeighbor(await buildDocx({ bodyXml: paragraph(inner) }))
    expect(textOf(reopened.blocks[0])).toBe('beforeonetwothreeAFTER')
  })
})

describe('existing transparent run wrappers', () => {
  it.each([
    ['smartTag', '<w:smartTag w:uri="urn:example" w:element="tag">', '</w:smartTag>'],
    ['sdt', '<w:sdt><w:sdtContent>', '</w:sdtContent></w:sdt>'],
    ['dir', '<w:dir w:val="ltr">', '</w:dir>'],
    ['bdo', '<w:bdo w:val="ltr">', '</w:bdo>'],
  ])('keeps %s content through the same save path', async (_, open, close) => {
    const { parsed, reopened } = await editNeighbor(
      await buildDocx({ bodyXml: paragraph(open + run('wrapped') + close) }),
    )
    expect(textOf(parsed.blocks[0])).toBe('beforewrappedafter')
    expect(textOf(reopened.blocks[0])).toBe('beforewrappedAFTER')
  })
})
