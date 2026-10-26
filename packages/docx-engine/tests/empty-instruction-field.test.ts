import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type GeneratedBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
const sep = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const instr = (text: string) =>
  `<w:r><w:instrText xml:space="preserve"> ${text} </w:instrText></w:r>`
const t = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`

const count = (xml: string, type: string) =>
  (xml.match(new RegExp(`w:fldCharType="${type}"`, 'g')) ?? []).length

const documentXmlOf = async (bytes: Uint8Array) =>
  (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')

/** A citation whose instruction was lost (the pre-fix paste in genoffice#2099):
 *  the paragraph must stay editable instead of becoming a protected block. */
describe('a field with an empty instruction', () => {
  const paragraph =
    '<w:p>' +
    t('As shown ') +
    begin +
    instr('') +
    sep +
    t('(Smith, 2020)') +
    end +
    t('.') +
    '</w:p>'

  it('keeps the paragraph editable with the result as an empty-instruction run', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: paragraph }))
    const block = doc.blocks[0]
    expect(block.type).toBe('paragraph')
    expect(block.runs).toEqual([
      expect.objectContaining({ text: 'As shown ' }),
      expect.objectContaining({ text: '(Smith, 2020)', instrField: '' }),
      expect.objectContaining({ text: '.' }),
    ])
  })

  it('saves back as one balanced field and reopens editable', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: paragraph }))
    const bytes = await saveDocx(doc, [
      { kind: 'generated', block: doc.blocks[0] as GeneratedBlock },
    ])
    const xml = await documentXmlOf(bytes)
    expect([count(xml, 'begin'), count(xml, 'end')]).toEqual([1, 1])
    const reopened = await parseDocx(bytes)
    expect(reopened.blocks[0].type).toBe('paragraph')
    expect(reopened.blocks[0].runs?.map((run) => run.text).join('')).toBe('As shown (Smith, 2020).')
  })

  it('still protects a paragraph whose empty field nests another field', async () => {
    const nested =
      '<w:p>' +
      begin +
      instr('') +
      begin +
      instr('PAGE') +
      sep +
      t('3') +
      end +
      sep +
      t('x') +
      end +
      '</w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: nested }))
    expect(doc.blocks[0].type).toBe('passthrough')
  })
})
