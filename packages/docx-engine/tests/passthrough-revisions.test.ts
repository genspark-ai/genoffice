import { describe, expect, it } from 'vitest'
import { applyTrackedChangesXml, parseDocx, readTrackedChangeRecords } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const AT = 'w:author="A" w:date="2026-01-01T00:00:00Z"'

// issue #2808: deleting a hyperlink with Track Changes on wraps the field's runs
// in w:del (fldChar + delInstrText + delText) — that paragraph is a passthrough
// block, and its revisions must still be listable and resolvable
const DELETED_FIELD_P =
  '<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>' +
  `<w:del w:id="1" ${AT}>` +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:delInstrText xml:space="preserve"> HYPERLINK &quot;https://example.com&quot; </w:delInstrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:delText>link text</w:delText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '</w:del>' +
  `<w:ins w:id="90" ${AT}><w:r><w:t> other</w:t></w:r></w:ins>` +
  '</w:p>'

// an *inserted* field parses editable (no passthrough) — the reporter confirms
// that path already works; the summary must not appear there
const INSERTED_FIELD_P =
  '<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>' +
  `<w:ins w:id="5" ${AT}>` +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> HYPERLINK &quot;https://example.com&quot; </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>link text</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '</w:ins></w:p>'

describe('passthrough block revision records', () => {
  it('the deleted-field paragraph surfaces its records as a block summary', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: DELETED_FIELD_P }))
    expect(doc.blocks[0].type).toBe('passthrough')
    expect(doc.blocks[0].revisionRecords).toEqual([
      { kind: 'del', author: 'A', date: '2026-01-01T00:00:00Z', id: '1' },
      { kind: 'ins', author: 'A', date: '2026-01-01T00:00:00Z', id: '90' },
    ])
  })

  it('an inserted field stays editable and carries no summary', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: INSERTED_FIELD_P }))
    expect(doc.blocks[0].type).not.toBe('passthrough')
    expect(doc.blocks[0].revisionRecords).toBeUndefined()
  })

  it('readTrackedChangeRecords maps kinds and skips revision-free fragments', () => {
    expect(
      readTrackedChangeRecords(
        `<w:p><w:pPr><w:rPr><w:del w:id="7" ${AT}/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
      ),
    ).toEqual([{ kind: 'del', author: 'A', date: '2026-01-01T00:00:00Z', id: '7' }])
    expect(readTrackedChangeRecords(`<w:p><w:r><w:t>clean</w:t></w:r></w:p>`)).toEqual([])
  })

  it('the XML-level accept removes the deleted field with its result runs', async () => {
    const { xml } = applyTrackedChangesXml(DELETED_FIELD_P, 'accept')
    expect(xml).not.toContain('<w:del ')
    expect(xml).not.toContain('delInstrText')
    expect(xml).not.toContain('link text')
    expect(xml).not.toContain('fldChar')
    expect(xml).toContain(' other')
  })

  it('the XML-level reject restores the deleted text as live runs', async () => {
    const { xml } = applyTrackedChangesXml(DELETED_FIELD_P, 'reject')
    expect(xml).not.toContain('<w:del ')
    expect(xml).not.toContain('delInstrText')
    expect(xml).toContain('link text')
    expect(xml).not.toContain('<w:ins ')
    expect(xml).not.toContain(' other')
  })
})
