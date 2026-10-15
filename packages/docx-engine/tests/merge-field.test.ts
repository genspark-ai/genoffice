import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * MERGEFIELD: Word resolves the result from a mail-merge data source when the
 * user merges, so an unmerged main document carries no meaningful cached text.
 * GenOffice keeps the field so the template survives a round-trip, and shows the
 * guillemet placeholder Word itself uses when nothing is cached yet.
 */

/** the fldChar complex-field form Word writes; `cached === null` omits the result */
const mergeFieldParagraph = (
  instr: string,
  cached: string | null,
  before = '',
  after = '',
): string =>
  `<w:p>${before}` +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  (cached === null
    ? ''
    : '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      `<w:r><w:t xml:space="preserve">${cached}</w:t></w:r>`) +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  `${after}</w:p>`

const documentXmlOf = async (bytes: Uint8Array): Promise<string> =>
  (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')

describe('MERGEFIELD', () => {
  it('keeps the field in a body paragraph instead of falling back to passthrough', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: mergeFieldParagraph('MERGEFIELD CustomerName \\* MERGEFORMAT', '«CustomerName»'),
      }),
    )
    expect(doc.blocks[0].runs).toEqual([
      { text: '«CustomerName»', instrField: 'MERGEFIELD CustomerName \\* MERGEFORMAT' },
    ])
  })

  it('keeps the enclosing hyperlink on a merge field inside w:hyperlink', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:hyperlink r:id="rId20">' +
          mergeFieldParagraph('MERGEFIELD Email', '«Email»').slice(
            '<w:p>'.length,
            -'</w:p>'.length,
          ) +
          '</w:hyperlink></w:p>',
        extraRels:
          '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:a@b.test" TargetMode="External"/>',
      }),
    )
    expect(doc.blocks[0].runs).toEqual([
      {
        text: '«Email»',
        instrField: 'MERGEFIELD Email',
        link: { href: 'mailto:a@b.test', rId: 'rId20' },
      },
    ])
  })

  it('shows the guillemet placeholder when Word cached no result yet', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: mergeFieldParagraph('MERGEFIELD CustomerName', null) }),
    )
    expect(doc.blocks[0].runs).toEqual([
      { text: '«CustomerName»', instrField: 'MERGEFIELD CustomerName' },
    ])
  })

  it('keeps surrounding prose and multiple fields in the same paragraph', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: mergeFieldParagraph(
          'MERGEFIELD FirstName',
          null,
          '<w:r><w:t xml:space="preserve">Dear </w:t></w:r>',
          '<w:r><w:t xml:space="preserve">, see you on </w:t></w:r>' +
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
            '<w:r><w:instrText xml:space="preserve"> DATE \\@ "yyyy-MM-dd" </w:instrText></w:r>' +
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
            '<w:r><w:t>2026-01-01</w:t></w:r>' +
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
        ),
      }),
    )
    expect(doc.blocks[0].runs).toEqual([
      { text: 'Dear ' },
      { text: '«FirstName»', instrField: 'MERGEFIELD FirstName' },
      { text: ', see you on ' },
      { text: '2026-01-01', instrField: 'DATE \\@ "yyyy-MM-dd"' },
    ])
  })

  it('reads a quoted field name whole', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: mergeFieldParagraph('MERGEFIELD "Postal Address" \\* MERGEFORMAT', null),
      }),
    )
    expect(doc.blocks[0].runs).toEqual([
      { text: '«Postal Address»', instrField: 'MERGEFIELD "Postal Address" \\* MERGEFORMAT' },
    ])
  })

  it('reads a CJK field name', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: mergeFieldParagraph('MERGEFIELD \u59d3\u540d', null) }),
    )
    expect(doc.blocks[0].runs).toEqual([
      { text: '«\u59d3\u540d»', instrField: 'MERGEFIELD \u59d3\u540d' },
    ])
  })

  it('folds the w:fldSimple form too', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:fldSimple w:instr=" MERGEFIELD CustomerName ">' +
          '<w:r><w:t>«CustomerName»</w:t></w:r>' +
          '</w:fldSimple></w:p>',
      }),
    )
    expect(doc.blocks[0].runs).toEqual([
      { text: '«CustomerName»', instrField: 'MERGEFIELD CustomerName' },
    ])
  })

  it('writes the field back so Word still sees a mail-merge field after a round-trip', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: mergeFieldParagraph('MERGEFIELD CustomerName \\* MERGEFORMAT', '«CustomerName»'),
      }),
    )
    const out = await saveDocx(parsed, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [
            { text: 'Dear ' },
            { text: '«CustomerName»', instrField: 'MERGEFIELD CustomerName \\* MERGEFORMAT' },
          ],
        },
      },
    ])
    const xml = await documentXmlOf(out)
    expect(xml).toContain(
      '<w:instrText xml:space="preserve"> MERGEFIELD CustomerName \\* MERGEFORMAT </w:instrText>',
    )
    expect(xml).toContain('«CustomerName»')

    const reparsed = await parseDocx(out)
    expect(reparsed.blocks[0].runs).toEqual([
      { text: 'Dear ' },
      { text: '«CustomerName»', instrField: 'MERGEFIELD CustomerName \\* MERGEFORMAT' },
    ])
  })

  it('keeps a w:dirty MERGEFIELD flagged for Word to refresh on open', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
          '<w:r><w:instrText xml:space="preserve"> MERGEFIELD Amount </w:instrText></w:r>' +
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
          '<w:r><w:t>«Amount»</w:t></w:r>' +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
      }),
    )
    expect(doc.blocks[0].runs?.[0]).toMatchObject({
      text: '«Amount»',
      instrField: 'MERGEFIELD Amount',
      fldDirty: true,
    })
  })
})
