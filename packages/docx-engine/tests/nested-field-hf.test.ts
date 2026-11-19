import { describe, expect, it } from 'vitest'
import { parseDocx, PAGE_MARK } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * A nested complex field inside a header's DOCPROPERTY: the child PAGE's
 * cached result rides the outer instruction as w:instrText, and the lazy
 * begin-to-end span stops at the inner end. The old classification tested
 * `\bPAGE\b` against the whole concatenated instruction — the child's
 * PAGE leaked through and the outer field was flattened into a page-number
 * marker + a stale cache (genoffice#2458).
 */

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const NESTED_FIELD_HEADER =
  XML_DECL +
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<w:p><w:r>' +
  // outer field begin
  '<w:fldChar w:fldCharType="begin"/>' +
  // outer instruction: DOCPROPERTY "STR{ PAGE }" — the property name carries a nested field
  '<w:instrText xml:space="preserve"> DOCPROPERTY  "STR</w:instrText>' +
  // nested child PAGE field begins inside the outer instruction
  '<w:fldChar w:fldCharType="begin"/>' +
  '<w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText>' +
  '<w:fldChar w:fldCharType="separate"/>' +
  // the child's cached result is stored as w:instrText (!) inside the outer instruction
  '<w:instrText>2</w:instrText>' +
  '<w:fldChar w:fldCharType="end"/>' +
  // outer instruction continues after the nested field closes
  '<w:instrText xml:space="preserve">"  \\* MERGEFORMAT </w:instrText>' +
  '<w:fldChar w:fldCharType="separate"/>' +
  // outer cached result: the real display text
  '<w:r><w:rPr><w:b/></w:rPr><w:t>Strana druhá</w:t></w:r>' +
  '<w:fldChar w:fldCharType="end"/>' +
  '</w:r></w:p></w:hdr>'

const CUSTOM_XML =
  XML_DECL +
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-document-properties">' +
  '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="STR1"><vt:lpwstr>Strana první</vt:lpwstr></property>' +
  '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="STR2"><vt:lpwstr>Strana druhá</vt:lpwstr></property>' +
  '</Properties>'

async function buildNestedFieldDocx(): Promise<Uint8Array> {
  return buildDocx({
    bodyXml: '<w:p><w:r><w:t>Body text.</w:t></w:r></w:p>',
    sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
    extraRels:
      '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
    extraParts: [
      {
        path: 'word/header1.xml',
        xml: NESTED_FIELD_HEADER,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
      },
      {
        path: 'docProps/custom.xml',
        xml: CUSTOM_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.custom-properties+xml',
      },
    ],
  })
}

/** The same DOCPROPERTY field, but the child's cached result is a display run
 *  (w:t) instead of field code (w:instrText). */
const NESTED_FIELD_HEADER_WT_CACHE =
  XML_DECL +
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<w:p><w:r>' +
  '<w:fldChar w:fldCharType="begin"/>' +
  '<w:instrText xml:space="preserve"> DOCPROPERTY  "STR</w:instrText>' +
  '<w:fldChar w:fldCharType="begin"/>' +
  '<w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText>' +
  '<w:fldChar w:fldCharType="separate"/>' +
  '<w:t>2</w:t>' +
  '<w:fldChar w:fldCharType="end"/>' +
  '<w:instrText xml:space="preserve">"  \\* MERGEFORMAT </w:instrText>' +
  '<w:fldChar w:fldCharType="separate"/>' +
  '<w:r><w:rPr><w:b/></w:rPr><w:t>Strana druhá</w:t></w:r>' +
  '<w:fldChar w:fldCharType="end"/>' +
  '</w:r></w:p></w:hdr>'

/** { = { PAGE } - 1 } — a formula field whose instruction holds the child. */
const FORMULA_FIELD_HEADER =
  XML_DECL +
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:p><w:r>' +
  '<w:fldChar w:fldCharType="begin"/>' +
  '<w:instrText xml:space="preserve"> = </w:instrText>' +
  '<w:fldChar w:fldCharType="begin"/>' +
  '<w:instrText xml:space="preserve"> PAGE </w:instrText>' +
  '<w:fldChar w:fldCharType="separate"/>' +
  '<w:t>2</w:t>' +
  '<w:fldChar w:fldCharType="end"/>' +
  '<w:instrText xml:space="preserve"> - 1 </w:instrText>' +
  '<w:fldChar w:fldCharType="separate"/>' +
  '<w:t>1</w:t>' +
  '<w:fldChar w:fldCharType="end"/>' +
  '</w:r></w:p></w:hdr>'

/** { IF … } whose cached result holds a nested PAGE field. */
const RESULT_NESTED_FIELD_HEADER =
  XML_DECL +
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:p><w:r>' +
  '<w:fldChar w:fldCharType="begin"/>' +
  '<w:instrText xml:space="preserve"> IF </w:instrText>' +
  '<w:fldChar w:fldCharType="separate"/>' +
  '<w:fldChar w:fldCharType="begin"/>' +
  '<w:instrText xml:space="preserve"> PAGE </w:instrText>' +
  '<w:fldChar w:fldCharType="separate"/>' +
  '<w:t>7</w:t>' +
  '<w:fldChar w:fldCharType="end"/>' +
  '<w:fldChar w:fldCharType="end"/>' +
  '</w:r></w:p></w:hdr>'

async function headerDocx(headerXml: string): Promise<Uint8Array> {
  return buildDocx({
    bodyXml: '<w:p><w:r><w:t>Body text.</w:t></w:r></w:p>',
    sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
    extraRels:
      '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
    extraParts: [
      {
        path: 'word/header1.xml',
        xml: headerXml,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
      },
    ],
  })
}

describe('nested complex field in header (genoffice#2458)', () => {
  it('the outer DOCPROPERTY cache survives without a PAGE marker prefix', async () => {
    const doc = await parseDocx(await buildNestedFieldDocx())
    // the old code produced: PAGE_MARK + "Strana druhá" (the child's PAGE
    // leaked into the parent's classification)
    const text = doc.headerText ?? ''
    expect(text).not.toContain('\uE001') // no PAGE_MARK
    expect(text).toContain('Strana druhá') // the outer cache survives
    expect(doc.headerHasPageNumber).toBe(false) // the child PAGE doesn't classify the parent
  })

  it('a real PAGE field still classifies as PAGE', async () => {
    // the header has a standalone PAGE field — the classification must not regress
    const simplePageHeader =
      XML_DECL +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:p><w:r>' +
      '<w:fldChar w:fldCharType="begin"/>' +
      '<w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText>' +
      '<w:fldChar w:fldCharType="separate"/>' +
      '<w:r><w:t>1</w:t></w:r>' +
      '<w:fldChar w:fldCharType="end"/>' +
      '</w:r></w:p></w:hdr>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>Body.</w:t></w:r></w:p>',
        sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
        extraRels:
          '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
        extraParts: [
          {
            path: 'word/header1.xml',
            xml: simplePageHeader,
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
          },
        ],
      }),
    )
    expect(doc.headerHasPageNumber).toBe(true)
    expect(doc.headerText).toContain(PAGE_MARK)
  })

  it('the same DOCPROPERTY with the child cache in a display run keeps the outer cache alone', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>Body text.</w:t></w:r></w:p>',
        sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
        extraRels:
          '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
        extraParts: [
          {
            path: 'word/header1.xml',
            xml: NESTED_FIELD_HEADER_WT_CACHE,
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
          },
          {
            path: 'docProps/custom.xml',
            xml: CUSTOM_XML,
            contentType: 'application/vnd.openxmlformats-officedocument.custom-properties+xml',
          },
        ],
      }),
    )
    const text = doc.headerText ?? ''
    // the child's stale cache ("2") must not ride along with the outer cache
    expect(text).not.toContain('2')
    expect(text).toContain('Strana druhá')
    expect(doc.headerHasPageNumber).toBe(false)
  })

  it('a formula field keeps its own result, not the child page number', async () => {
    const doc = await parseDocx(await headerDocx(FORMULA_FIELD_HEADER))
    const text = doc.headerText ?? ''
    // "{ = { PAGE } - 1 }": the child's PAGE is field code, so only the
    // formula's own cached result displays — not the child's stale "2"
    expect(text).toBe('1')
    expect(text).not.toContain('2')
    expect(doc.headerHasPageNumber).toBe(false)
  })

  it('a nested PAGE in the cached result still becomes a page marker', async () => {
    const doc = await parseDocx(await headerDocx(RESULT_NESTED_FIELD_HEADER))
    const text = doc.headerText ?? ''
    expect(text).toContain(PAGE_MARK)
    expect(text).not.toContain('7')
    expect(doc.headerHasPageNumber).toBe(true)
  })

  it('a parent field with no instruction of its own is not a PAGE field', async () => {
    // the parent's own instruction is a space; the PAGE belongs to the child.
    // The parent's instruction is its own — a child's instruction must not
    // make the parent a page-number field.
    const doc = await parseDocx(
      await headerDocx(
        XML_DECL +
          '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          '<w:p><w:r>' +
          '<w:fldChar w:fldCharType="begin"/>' +
          '<w:instrText xml:space="preserve"> </w:instrText>' +
          '<w:fldChar w:fldCharType="begin"/>' +
          '<w:instrText xml:space="preserve"> PAGE </w:instrText>' +
          '<w:fldChar w:fldCharType="separate"/>' +
          '<w:t>2</w:t>' +
          '<w:fldChar w:fldCharType="end"/>' +
          '<w:fldChar w:fldCharType="separate"/>' +
          '<w:t>cached</w:t>' +
          '<w:fldChar w:fldCharType="end"/>' +
          '</w:r></w:p></w:hdr>',
      ),
    )
    const text = doc.headerText ?? ''
    expect(text).toBe('cached')
    expect(text).not.toContain(PAGE_MARK)
  })
})
