import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

describe('Zotero document properties', () => {
  it('creates Word-compatible chunked preferences and reads them back', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Text</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const data = JSON.stringify({
      style: 'http://www.zotero.org/styles/apa',
      value: 'x'.repeat(600),
    })
    const saved = await saveDocx(
      parsed,
      parsed.blocks
        .filter((block) => !block.hidden)
        .map((block) => ({ kind: 'original' as const, docxIndex: block.docxIndex! })),
      { zoteroDocumentData: data },
    )

    const reparsed = await parseDocx(saved)
    expect(reparsed.zoteroDocumentData).toBe(data)

    const zip = await JSZip.loadAsync(saved)
    const custom = await zip.file('docProps/custom.xml')!.async('string')
    expect(custom).toContain('name="ZOTERO_PREF_1"')
    expect(custom).toContain('name="ZOTERO_PREF_3"')
    expect(await zip.file('_rels/.rels')!.async('string')).toContain(
      'relationships/custom-properties',
    )
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
      'officedocument.custom-properties+xml',
    )
  })

  it('replaces only Zotero properties and preserves unrelated custom properties', async () => {
    const customXml =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ' +
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      '<property fmtid="x" pid="2" name="Owner"><vt:lpwstr>A&amp;B</vt:lpwstr></property>' +
      '<property fmtid="x" pid="3" name="ZOTERO_PREF_1"><vt:lpwstr>old</vt:lpwstr></property>' +
      '</Properties>'
    const source = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>Text</w:t></w:r></w:p>',
      extraParts: [
        {
          path: 'docProps/custom.xml',
          xml: customXml,
          contentType: 'application/vnd.openxmlformats-officedocument.custom-properties+xml',
        },
      ],
    })
    const parsed = await parseDocx(source)
    expect(parsed.zoteroDocumentData).toBe('old')
    const saved = await saveDocx(
      parsed,
      parsed.blocks
        .filter((block) => !block.hidden)
        .map((block) => ({ kind: 'original' as const, docxIndex: block.docxIndex! })),
      { zoteroDocumentData: 'new' },
    )
    const custom = await (await JSZip.loadAsync(saved)).file('docProps/custom.xml')!.async('string')
    expect(custom).toContain('name="Owner"')
    expect(custom).toContain('A&amp;B')
    expect(custom).toContain('name="ZOTERO_PREF_1"')
    expect(custom).toContain('>new<')
    expect(custom).not.toContain('>old<')
  })
})
