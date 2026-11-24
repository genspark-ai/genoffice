import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, rewriteSeqResultsXml, saveDocx, seqFieldsOfXml } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const field = (instr: string, result: string, attrs = '') =>
  `<w:r><w:fldChar w:fldCharType="begin"${attrs}/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

describe('locked fields (w:fldLock)', () => {
  it('are read from complex and simple fields, and from a cross-reference', async () => {
    const body =
      `<w:p>${field('DATE', '1 May', ' w:fldLock="1"')}${field('FILENAME', 'a.docx')}</w:p>` +
      `<w:p>${field('REF _Ref1 \\h', 'Figure 1', ' w:fldLock="true"')}</w:p>` +
      '<w:p><w:fldSimple w:instr=" TIME " w:fldLock="1"><w:r><w:t>10:00</w:t></w:r></w:fldSimple></w:p>' +
      `<w:p>${field('DATE', '2 May', ' w:fldLock="0"')}</w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    const runs = doc.blocks.flatMap((b) => b.runs ?? []).filter((r) => r.instrField || r.refField)
    expect(runs.map((r) => [r.text, r.fldLock ?? false])).toEqual([
      ['1 May', true],
      ['a.docx', false],
      ['Figure 1', true],
      ['10:00', true],
      ['2 May', false],
    ])
  })

  it('are written back on the field that carries the lock and on no other', async () => {
    const body = `<w:p>${field('DATE', '1 May', ' w:fldLock="1"')}${field('FILENAME', 'a.docx')}</w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    const saved = await saveDocx(doc, [
      {
        kind: 'generated',
        block: { type: 'paragraph', runs: [{ text: 'x' }, ...doc.blocks[0].runs!] },
      },
    ])
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml.match(/w:fldLock/g)).toHaveLength(1)
    expect(xml).toContain(
      '<w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> DATE </w:instrText>',
    )
    const again = await parseDocx(saved)
    const runs = again.blocks[0].runs!.filter((r) => r.instrField)
    expect(runs.map((r) => r.fldLock ?? false)).toEqual([true, false])
  })

  it('keep their number when the captions of a block are renumbered', () => {
    const xml = `<w:p>${field('SEQ Figure', '1', ' w:fldLock="1"')}${field('SEQ Figure', '2')}</w:p>`
    expect(seqFieldsOfXml(xml).map((f) => f.locked)).toEqual([true, false])
    const out = rewriteSeqResultsXml(xml, ['7', '8'])
    expect(seqFieldsOfXml(out).map((f) => f.result)).toEqual(['1', '8'])
  })
})
