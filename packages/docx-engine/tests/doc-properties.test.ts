import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const CORE_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
  'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
  'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
  '<dc:title>测试文档</dc:title><dc:subject>案号 2026-114</dc:subject><dc:creator>作者甲</dc:creator>' +
  '<cp:keywords>保密</cp:keywords><cp:lastModifiedBy>作者甲</cp:lastModifiedBy><cp:revision>3</cp:revision>' +
  '<dcterms:created xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:created>' +
  '<dcterms:modified xsi:type="dcterms:W3CDTF">2020-01-02T00:00:00Z</dcterms:modified>' +
  '</cp:coreProperties>'

const APP_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
  'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
  '<Template>Normal.dotm</Template><Company>某某律师事务所</Company><Application>Microsoft Office Word</Application>' +
  '<AppVersion>16.0000</AppVersion></Properties>'

const CUSTOM_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ' +
  'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
  '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="MatterId">' +
  '<vt:lpwstr>M-2026-114</vt:lpwstr></property>' +
  '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Reviewer">' +
  '<vt:lpwstr>乙律师</vt:lpwstr></property>' +
  '</Properties>'

const BODY = '<w:p><w:r><w:t>正文</w:t></w:r></w:p>'

const docxWithProps = () =>
  buildDocx({
    bodyXml: BODY,
    extraParts: [
      {
        path: 'docProps/core.xml',
        xml: CORE_XML,
        contentType: 'application/vnd.openxmlformats-package.core-properties+xml',
      },
      {
        path: 'docProps/app.xml',
        xml: APP_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
      },
      {
        path: 'docProps/custom.xml',
        xml: CUSTOM_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.custom-properties+xml',
      },
    ],
  })

const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
  doc.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))

const partOf = async (bytes: Uint8Array, path: string): Promise<string> => {
  const file = (await JSZip.loadAsync(bytes)).file(path)
  return file ? file.async('string') : ''
}

describe('document properties', () => {
  it('reads the core, app and custom properties of a package', async () => {
    const doc = await parseDocx(await docxWithProps())
    expect(doc.properties.core).toEqual({
      title: '测试文档',
      subject: '案号 2026-114',
      creator: '作者甲',
      keywords: '保密',
      lastModifiedBy: '作者甲',
      revision: '3',
      created: '2020-01-01T00:00:00Z',
      modified: '2020-01-02T00:00:00Z',
    })
    expect(doc.properties.app).toEqual({
      template: 'Normal.dotm',
      company: '某某律师事务所',
      application: 'Microsoft Office Word',
      appVersion: '16.0000',
    })
    expect(doc.properties.custom).toEqual({ MatterId: 'M-2026-114', Reviewer: '乙律师' })
  })

  it('a document without docProps has no properties but still opens', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    expect(doc.properties).toEqual({ core: {}, app: {}, custom: {} })
  })

  it('sets a property and leaves every other one byte-identical', async () => {
    const doc = await parseDocx(await docxWithProps())
    const out = await saveDocx(doc, originalOrder(doc), {
      savedAt: '2026-07-28T08:30:00Z',
      properties: {
        core: { title: 'Motion to Dismiss' },
        app: { company: 'Barnes & Cole LLP' },
        custom: { MatterId: 'M-2026-999' },
      },
    })
    const core = await partOf(out, 'docProps/core.xml')
    expect(core).toContain('<dc:title>Motion to Dismiss</dc:title>')
    expect(core).toContain('<dc:subject>案号 2026-114</dc:subject>')
    expect(core).toContain('<cp:keywords>保密</cp:keywords>')
    expect(core).toContain(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:created>',
    )
    const app = await partOf(out, 'docProps/app.xml')
    expect(app).toContain('<Company>Barnes &amp; Cole LLP</Company>')
    expect(app).toContain('<Template>Normal.dotm</Template>')
    expect(app).toContain('<Application>Microsoft Office Word</Application>')
    const custom = await partOf(out, 'docProps/custom.xml')
    expect(custom).toContain('<vt:lpwstr>M-2026-999</vt:lpwstr>')
    expect(custom).toContain('<vt:lpwstr>乙律师</vt:lpwstr>')
    expect(custom).toContain('name="MatterId"')
  })

  it('removes the property whose value is null', async () => {
    const doc = await parseDocx(await docxWithProps())
    const out = await saveDocx(doc, originalOrder(doc), {
      properties: {
        core: { creator: null, lastModifiedBy: null },
        app: { company: null },
        custom: { Reviewer: null },
      },
    })
    const core = await partOf(out, 'docProps/core.xml')
    expect(core).not.toContain('dc:creator')
    expect(core).not.toContain('lastModifiedBy')
    expect(core).toContain('<dc:title>测试文档</dc:title>')
    expect(core).toContain('<cp:keywords>保密</cp:keywords>')
    const app = await partOf(out, 'docProps/app.xml')
    expect(app).not.toContain('Company')
    expect(app).toContain('<Template>Normal.dotm</Template>')
    const custom = await partOf(out, 'docProps/custom.xml')
    expect(custom).not.toContain('Reviewer')
    expect(custom).toContain('M-2026-114')
  })

  it('clears every property of the three parts', async () => {
    const doc = await parseDocx(await docxWithProps())
    const out = await saveDocx(doc, originalOrder(doc), { properties: { clear: 'all' } })
    const core = await partOf(out, 'docProps/core.xml')
    expect(core).not.toContain('dc:title')
    expect(core).not.toContain('dc:creator')
    expect(core).not.toContain('cp:keywords')
    expect(core).toContain('<cp:coreProperties')
    const app = await partOf(out, 'docProps/app.xml')
    expect(app).not.toContain('Company')
    expect(app).not.toContain('Template')
    expect(app).toContain('<Properties')
    const custom = await partOf(out, 'docProps/custom.xml')
    expect(custom).not.toContain('MatterId')
    expect(custom).not.toContain('Reviewer')
    expect(custom).toContain('<Properties')
  })

  it('clears every property and still sets the values given in the same op', async () => {
    const doc = await parseDocx(await docxWithProps())
    const out = await saveDocx(doc, originalOrder(doc), {
      properties: { clear: 'all', core: { title: 'Motion to Dismiss' } },
    })
    const core = await partOf(out, 'docProps/core.xml')
    expect(core).not.toContain('dc:creator')
    expect(core).not.toContain('cp:keywords')
    expect(core).toContain('<dc:title>Motion to Dismiss</dc:title>')
  })

  it('a properties-only save writes the file instead of returning the original bytes', async () => {
    const bytes = await docxWithProps()
    const doc = await parseDocx(bytes)
    const out = await saveDocx(doc, originalOrder(doc), {
      properties: { core: { title: 'Motion to Dismiss' } },
    })
    expect(out).not.toEqual(bytes)
    expect(await partOf(out, 'docProps/core.xml')).toContain('Motion to Dismiss')
  })

  it('creates the docProps part of a document that has none', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const out = await saveDocx(doc, originalOrder(doc), {
      properties: { core: { title: '新建文档' } },
    })
    const zip = await JSZip.loadAsync(out)
    const core = await zip.file('docProps/core.xml')!.async('string')
    expect(core).toContain('<dc:title>新建文档</dc:title>')
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
      'PartName="/docProps/core.xml"',
    )
    expect(await zip.file('_rels/.rels')!.async('string')).toContain('Target="docProps/core.xml"')
  })

  it('a document without docProps still saves', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const out = await saveDocx(doc, originalOrder(doc), { watermark: '草稿' })
    expect((await JSZip.loadAsync(out)).file('docProps/core.xml')).toBeNull()
  })
})
