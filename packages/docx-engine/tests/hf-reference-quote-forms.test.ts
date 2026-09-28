import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { PAGE_MARK, parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const BODY = '<w:p><w:r><w:t>body</w:t></w:r></w:p>'
const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const FOOTER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const hdrPart = (text: string) =>
  `${XML_DECL}<w:hdr ${W_NS}><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:hdr>`
const ftrPart = (text: string) =>
  `${XML_DECL}<w:ftr ${W_NS}><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:ftr>`

const REL = (id: string, kind: 'header' | 'footer', target: string) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${target}"/>`

async function zipText(bytes: Uint8Array, path: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(bytes)
  const file = zip.file(path)
  return file ? file.async('string') : null
}

describe('header/footer reference quote styles and element forms', () => {
  it('resolves single-quoted default/first/even references', async () => {
    const bytes = await buildDocx({
      bodyXml: BODY,
      extraRels:
        REL('rId80', 'header', 'header1.xml') +
        REL('rId81', 'header', 'header2.xml') +
        REL('rId82', 'header', 'header3.xml'),
      extraParts: [
        { path: 'word/header1.xml', xml: hdrPart('DEFAULT'), contentType: HEADER_TYPE },
        { path: 'word/header2.xml', xml: hdrPart('FIRST'), contentType: HEADER_TYPE },
        { path: 'word/header3.xml', xml: hdrPart('EVEN'), contentType: HEADER_TYPE },
      ],
      sectPrExtra:
        "<w:headerReference w:type='default' r:id='rId80'/>" +
        "<w:headerReference w:type='first' r:id='rId81'/>" +
        "<w:headerReference w:type='even' r:id='rId82'/>" +
        '<w:titlePg/>',
    })
    const parsed = await parseDocx(bytes)
    expect(parsed.headerText).toBe('DEFAULT')
    expect(parsed.headerFirst?.text).toBe('FIRST')
    expect(parsed.headerEven?.text).toBe('EVEN')
  })

  it('resolves empty-element-pair references (LibreOffice form)', async () => {
    const bytes = await buildDocx({
      bodyXml: BODY,
      extraRels:
        REL('rId83', 'header', 'header1.xml') +
        REL('rId84', 'footer', 'footer1.xml') +
        REL('rId85', 'header', 'header2.xml'),
      extraParts: [
        { path: 'word/header1.xml', xml: hdrPart('PAIR_HEADER'), contentType: HEADER_TYPE },
        { path: 'word/footer1.xml', xml: ftrPart('PAIR_FOOTER'), contentType: FOOTER_TYPE },
        { path: 'word/header2.xml', xml: hdrPart('PAIR_EVEN'), contentType: HEADER_TYPE },
      ],
      sectPrExtra:
        '<w:headerReference w:type="default" r:id="rId83"></w:headerReference>' +
        '<w:footerReference w:type="default" r:id="rId84"></w:footerReference>' +
        "<w:headerReference w:type='even' r:id='rId85'></w:headerReference>",
    })
    const parsed = await parseDocx(bytes)
    expect(parsed.headerText).toBe('PAIR_HEADER')
    expect(parsed.footerText).toBe('PAIR_FOOTER')
    expect(parsed.headerEven?.text).toBe('PAIR_EVEN')
  })

  it('treats a single-quoted w:type="odd" and an untyped pair form as the default part', async () => {
    const bytes = await buildDocx({
      bodyXml: BODY,
      extraRels: REL('rId86', 'header', 'header1.xml') + REL('rId87', 'header', 'header2.xml'),
      extraParts: [
        { path: 'word/header1.xml', xml: hdrPart('ODD_DEFAULT'), contentType: HEADER_TYPE },
        { path: 'word/header2.xml', xml: hdrPart('UNTYPED_DEFAULT'), contentType: HEADER_TYPE },
      ],
      sectPrExtra: "<w:headerReference w:type='odd' r:id='rId86'/>",
    })
    expect((await parseDocx(bytes)).headerText).toBe('ODD_DEFAULT')

    const untyped = await buildDocx({
      bodyXml: BODY,
      extraRels: REL('rId87', 'header', 'header2.xml'),
      extraParts: [
        { path: 'word/header2.xml', xml: hdrPart('UNTYPED_DEFAULT'), contentType: HEADER_TYPE },
      ],
      sectPrExtra: '<w:headerReference r:id="rId87"></w:headerReference>',
    })
    expect((await parseDocx(untyped)).headerText).toBe('UNTYPED_DEFAULT')
  })

  it('keeps editing a single-quoted default header pointed at its own part', async () => {
    const bytes = await buildDocx({
      bodyXml: BODY,
      extraRels: REL('rId88', 'header', 'header1.xml'),
      extraParts: [{ path: 'word/header1.xml', xml: hdrPart('BEFORE'), contentType: HEADER_TYPE }],
      sectPrExtra: "<w:headerReference w:type='default' r:id='rId88'/>",
    })
    const saved = await saveDocx(await parseDocx(bytes), [{ kind: 'original', docxIndex: 0 }], {
      header: { text: 'AFTER' },
    })
    expect(await zipText(saved, 'word/header1.xml')).toContain('AFTER')
    expect((await parseDocx(saved)).headerText).toBe('AFTER')
  })

  it('carries a PAGE field through a single-quoted reference', async () => {
    const pageFieldFooter =
      `${XML_DECL}<w:ftr ${W_NS}><w:p>` +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>7</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '</w:p></w:ftr>'
    const bytes = await buildDocx({
      bodyXml: BODY,
      extraRels: REL('rId89', 'footer', 'footer1.xml'),
      extraParts: [{ path: 'word/footer1.xml', xml: pageFieldFooter, contentType: FOOTER_TYPE }],
      sectPrExtra: "<w:footerReference w:type='default' r:id='rId89'></w:footerReference>",
    })
    const parsed = await parseDocx(bytes)
    expect(parsed.footerText).toBe(PAGE_MARK)
    expect(parsed.footerHasPageNumber).toBe(true)
  })
})
