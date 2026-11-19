import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const BIDI_PARA =
  '<w:p>' +
  '<w:r><w:t xml:space="preserve">before </w:t></w:r>' +
  '<w:dir w:val="ltr">' +
  '<w:r><w:rPr><w:rtl/></w:rPr><w:t>embedded</w:t></w:r>' +
  '<w:bdo w:val="rtl"><w:r><w:t xml:space="preserve"> overridden</w:t></w:r></w:bdo>' +
  '<w:r><w:t xml:space="preserve"> tail</w:t></w:r>' +
  '</w:dir>' +
  '<w:r><w:t xml:space="preserve"> after</w:t></w:r>' +
  '</w:p>'

describe('w:dir / w:bdo inline containers', () => {
  it('parses the runs inside bidirectional embedding and override containers', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BIDI_PARA }))
    const text = doc.blocks[0].runs!.map((r) => r.text).join('')
    expect(text).toBe('before embedded overridden tail after')
    expect(doc.blocks[0].runs!.find((r) => r.text === 'embedded')?.rtl).toBe(true)
  })

  it('keeps that text when the paragraph is regenerated on save', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BIDI_PARA }))
    const runs = doc.blocks[0].runs!.map((r) => ({
      ...r,
      text: r.text.replace('before', 'BEFORE'),
    }))
    const zip = await JSZip.loadAsync(
      await saveDocx(doc, [{ kind: 'generated', block: { type: 'paragraph', runs } }]),
    )
    const xml = await zip.file('word/document.xml')!.async('string')
    for (const piece of ['BEFORE', 'embedded', 'overridden', 'tail', 'after']) {
      expect(xml).toContain(piece)
    }
  })
})
