import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { PAGE_MARK, SECTION_PAGES_MARK, TOTAL_PAGES_MARK, parseDocx, saveDocx } from '../src'
import { buildDocx } from './helpers/build-docx'

const resultOf = (xml: string, keyword: string): string | undefined =>
  new RegExp(
    `<w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText[^>]*> ${keyword} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>([^<]*)</w:t>`,
  ).exec(xml)?.[1]

describe('the page fields of a header or footer written anew', () => {
  const footer = {
    text: '',
    pageNumber: false,
    paras: [
      { runs: [{ text: `Page ${PAGE_MARK} of ${TOTAL_PAGES_MARK}, ${SECTION_PAGES_MARK} here` }] },
    ],
  }
  const saved = async (options: Parameters<typeof saveDocx>[2]) => {
    const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>' })
    const zip = await JSZip.loadAsync(await saveDocx(await parseDocx(bytes), [], options))
    const name = Object.keys(zip.files).find((n) => /^word\/footer\d*\.xml$/.test(n))!
    return zip.file(name)!.async('string')
  }

  it('store the page the layout first shows the part on, as the page shows it, unmarked', async () => {
    const xml = await saved({
      footer: { ...footer, fieldResults: { page: 'iv', sectionPages: 6 } },
      pageCount: 12,
    })
    expect(resultOf(xml, 'PAGE')).toBe('iv')
    expect(resultOf(xml, 'NUMPAGES')).toBe('12')
    expect(resultOf(xml, 'SECTIONPAGES')).toBe('6')
    expect(xml).not.toContain('w:dirty')
  })

  it('store 1 where no layout was at hand, unmarked as Word leaves them', async () => {
    const xml = await saved({ footer })
    expect(resultOf(xml, 'PAGE')).toBe('1')
    expect(resultOf(xml, 'SECTIONPAGES')).toBe('1')
    expect(xml).not.toContain('w:dirty')
  })

  it('escape a chapter separator the XML would not take as it is', async () => {
    const xml = await saved({
      footer: { ...footer, fieldResults: { page: 'A&B-3', sectionPages: 1 } },
    })
    expect(resultOf(xml, 'PAGE')).toBe('A&amp;B-3')
  })
})
