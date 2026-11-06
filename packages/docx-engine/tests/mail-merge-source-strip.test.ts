import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { parseDocx, saveDocx, stripMailMergeSource } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * A merged document carries the merged values; the data source it came from
 * stays behind. Word keeps that source in settings.xml (w:mailMerge with
 * w:dataSource / w:query / ODSO's connection string) and in the external
 * relationship settings.xml.rels names. `exportAsMerged` drops both, and the
 * template's own save keeps them. genoffice#2097.
 */

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const BODY = '<w:p><w:r><w:t>Dear Ada</w:t></w:r></w:p>'
const ORIGINAL = [{ kind: 'original' as const, docxIndex: 0 }]

const SOURCE_PATH = 'C:\\Users\\me\\Mail Merge\\recipients.xlsx'

const SETTINGS_XML =
  `${XML_DECL}<w:settings ${W_NS} ${R_NS}>` +
  '<w:mailMerge>' +
  '<w:mainDocumentType w:val="formLetters"/>' +
  '<w:dataSource r:id="rId10"/>' +
  '<w:query w:val="SELECT * FROM `Sheet1$`"/>' +
  '<w:odso>' +
  `<w:udl w:val="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=${SOURCE_PATH}"/>` +
  '<w:table w:val="Sheet1$"/>' +
  '<w:srcRel r:id="rId10"/>' +
  '</w:odso>' +
  '</w:mailMerge>' +
  '</w:settings>'

const SETTINGS_RELS_XML =
  `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  '<Relationship Id="rId10" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource" ' +
  'Target="file:///C:/Users/me/Mail%20Merge/recipients.xlsx" TargetMode="External"/>' +
  '</Relationships>'

async function templateWithDataSource(): Promise<Uint8Array> {
  const base = await buildDocx({
    bodyXml: BODY,
    extraParts: [{ path: 'word/settings.xml', xml: SETTINGS_XML, contentType: SETTINGS_TYPE }],
  })
  const zip = await JSZip.loadAsync(base)
  zip.file('word/_rels/settings.xml.rels', SETTINGS_RELS_XML)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

async function partOf(bytes: Uint8Array, path: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(bytes)
  const file = zip.file(path)
  return file ? file.async('string') : null
}

describe('exportAsMerged', () => {
  it('drops the mail merge block and the data-source relationship from the merge product', async () => {
    const parsed = await parseDocx(await templateWithDataSource())
    const merged = await saveDocx(parsed, ORIGINAL, { exportAsMerged: true })

    const settings = await partOf(merged, 'word/settings.xml')
    expect(settings).not.toContain('mailMerge')
    expect(settings).not.toContain('recipients.xlsx')
    expect(settings).not.toContain('Data Source')
    const rels = await partOf(merged, 'word/_rels/settings.xml.rels')
    expect(rels).not.toContain('recipients.xlsx')
    expect(rels).not.toContain('rId10')
  })

  it('leaves the merged body untouched', async () => {
    const parsed = await parseDocx(await templateWithDataSource())
    const merged = await saveDocx(parsed, ORIGINAL, { exportAsMerged: true })
    expect(await partOf(merged, 'word/document.xml')).toContain('Dear Ada')
  })

  it('keeps the block when the template itself is saved', async () => {
    const parsed = await parseDocx(await templateWithDataSource())
    const saved = await saveDocx(parsed, ORIGINAL)
    const settings = await partOf(saved, 'word/settings.xml')
    expect(settings).toContain('mailMerge')
    expect(settings).toContain('recipients.xlsx')
    const rels = await partOf(saved, 'word/_rels/settings.xml.rels')
    expect(rels).toContain('recipients.xlsx')
  })

  it('returns bytes untouched when there is no data source', async () => {
    const plain = await buildDocx({ bodyXml: BODY })
    expect(await stripMailMergeSource(plain)).toEqual(plain)
  })
})
