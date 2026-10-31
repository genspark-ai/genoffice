import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  PAGE_MARK,
  SECTION_PAGES_MARK,
  TOTAL_PAGES_MARK,
  parseDocx,
  saveDocx,
  type SaveBlock,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

/** SECTIONPAGES in a footer (#2331): a mark the renderer fills with the section's page count */

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const FTR_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'

async function footerDocx(inner: string) {
  return buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    extraRels:
      '<Relationship Id="rId90" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
    extraParts: [
      { path: 'word/footer1.xml', xml: `<w:ftr ${W_NS}>${inner}</w:ftr>`, contentType: FTR_TYPE },
    ],
    sectPrExtra: '<w:footerReference w:type="default" r:id="rId90"/>',
  })
}

/** python-docx style: the whole field (begin, code, separate, result, end) in one w:r */
const oneRunField = (instr: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/>' +
  `<w:instrText xml:space="preserve"> ${instr} </w:instrText>` +
  '<w:fldChar w:fldCharType="separate"/><w:t>1</w:t><w:fldChar w:fldCharType="end"/></w:r>'

/** Word style: one run per field part */
const splitRunField = (instr: string, cached: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${cached}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

const footerString = (doc: Awaited<ReturnType<typeof parseDocx>>) =>
  doc.footerParas!.flatMap((p) => p.runs.map((r) => r.text)).join('')

describe('SECTIONPAGES header/footer field', () => {
  it('a one-run field becomes the section-pages mark, not an empty string (#2331)', async () => {
    const doc = await parseDocx(
      await footerDocx(
        '<w:p><w:r><w:t xml:space="preserve">- </w:t></w:r>' +
          oneRunField('PAGE') +
          '<w:r><w:t xml:space="preserve"> / </w:t></w:r>' +
          oneRunField('SECTIONPAGES') +
          '</w:p>',
      ),
    )
    expect(doc.footerText).toBe(`- ${PAGE_MARK} / ${SECTION_PAGES_MARK}`)
    expect(footerString(doc)).toBe(`- ${PAGE_MARK} / ${SECTION_PAGES_MARK}`)
  })

  it('a split-run field drops its stale cached count', async () => {
    const doc = await parseDocx(
      await footerDocx(`<w:p>${splitRunField('SECTIONPAGES', '1')}</w:p>`),
    )
    expect(footerString(doc)).toBe(SECTION_PAGES_MARK)
  })

  it('a fldSimple SECTIONPAGES and NUMPAGES keep apart', async () => {
    const doc = await parseDocx(
      await footerDocx(
        '<w:p><w:fldSimple w:instr=" SECTIONPAGES "><w:r><w:t>3</w:t></w:r></w:fldSimple>' +
          '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>',
      ),
    )
    expect(footerString(doc)).toBe(`${SECTION_PAGES_MARK}${TOTAL_PAGES_MARK}`)
  })

  it('an edited footer writes SECTIONPAGES back as a field and reparses to the mark', async () => {
    const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>' })
    const parsed = await parseDocx(bytes)
    const saveBlocks: SaveBlock[] = [{ kind: 'original', docxIndex: parsed.blocks[0].docxIndex! }]
    const saved = await saveDocx(parsed, saveBlocks, {
      footer: {
        text: '',
        paras: [
          {
            align: 'center',
            runs: [{ text: `${PAGE_MARK} / ${SECTION_PAGES_MARK} (${TOTAL_PAGES_MARK})` }],
          },
        ],
      },
    })
    const zip = await JSZip.loadAsync(saved)
    const name = Object.keys(zip.files).find((n) => /word\/footer\d+\.xml/.test(n))!
    const xml = await zip.file(name)!.async('string')
    expect(xml).toContain('<w:instrText xml:space="preserve"> SECTIONPAGES </w:instrText>')
    expect(xml).toContain('<w:instrText xml:space="preserve"> NUMPAGES </w:instrText>')
    expect(xml).not.toContain(SECTION_PAGES_MARK)
    expect(footerString(await parseDocx(saved))).toBe(
      `${PAGE_MARK} / ${SECTION_PAGES_MARK} (${TOTAL_PAGES_MARK})`,
    )
  })

  it('the legacy single-line footer writes SECTIONPAGES back too', async () => {
    const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>' })
    const parsed = await parseDocx(bytes)
    const saved = await saveDocx(
      parsed,
      [{ kind: 'original', docxIndex: parsed.blocks[0].docxIndex! }],
      { footer: { text: `${PAGE_MARK}/${SECTION_PAGES_MARK}` } },
    )
    const zip = await JSZip.loadAsync(saved)
    const name = Object.keys(zip.files).find((n) => /word\/footer\d+\.xml/.test(n))!
    const xml = await zip.file(name)!.async('string')
    expect(xml).toMatch(/ PAGE <[\s\S]* SECTIONPAGES </)
    expect(xml).not.toContain(SECTION_PAGES_MARK)
  })
})
