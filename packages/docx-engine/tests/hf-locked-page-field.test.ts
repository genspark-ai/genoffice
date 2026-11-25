import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { PAGE_MARK, parseDocx, saveDocx } from '../src'
import { buildDocx } from './helpers/build-docx'

const field = (instr: string, result: string, lock: boolean) =>
  `<w:r><w:fldChar w:fldCharType="begin"${lock ? ' w:fldLock="true"' : ''}/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

const footerOf = async (bytes: Uint8Array) => {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/footer1.xml')!.async('string')
}

describe('a locked page field of a header or footer', () => {
  const build = (lock: boolean) =>
    buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      footerXml: `<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${field('PAGE', '7', lock)}</w:p>`,
    })

  it('is read as a field holding the number it was saved with; an unlocked one stays live', async () => {
    const locked = await parseDocx(await build(true))
    expect(locked.footerParas![0].runs).toMatchObject([
      { text: 'Page ' },
      { text: '7', instrField: 'PAGE', fldLock: true },
    ])
    expect(locked.footerText).toBe('Page 7')
    const live = await parseDocx(await build(false))
    expect(live.footerParas![0].runs.map((r) => r.text).join('')).toBe(`Page ${PAGE_MARK}`)
  })

  it('keeps its lock and its number when the footer is written anew', async () => {
    const parsed = await parseDocx(await build(true))
    const paras = parsed.footerParas!.map((p, i) =>
      i === 0 ? { ...p, runs: [{ text: 'Sheet ' }, ...p.runs.slice(1)] } : p,
    )
    const xml = await footerOf(
      await saveDocx(parsed, [], { footer: { text: '', pageNumber: false, paras } }),
    )
    expect(xml).toContain('Sheet ')
    expect(xml).toMatch(
      /<w:fldChar w:fldCharType="begin" w:fldLock="true"\/><\/w:r><w:r><w:instrText[^>]*>\s*PAGE\s*<\/w:instrText>/,
    )
    expect(xml).toMatch(
      /w:fldCharType="separate"\/><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>7<\/w:t>/,
    )
    const again = await parseDocx(
      await saveDocx(parsed, [], { footer: { text: '', pageNumber: false, paras } }),
    )
    expect(again.footerParas![0].runs.at(-1)).toMatchObject({ text: '7', fldLock: true })
  })
})
