import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const CORE_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
  'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
  'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
  '<dc:title>测试文档</dc:title><dc:creator>作者甲</dc:creator>' +
  '<cp:lastModifiedBy>作者甲</cp:lastModifiedBy><cp:revision>3</cp:revision>' +
  '<dcterms:created xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:created>' +
  '<dcterms:modified xsi:type="dcterms:W3CDTF">2020-01-02T00:00:00Z</dcterms:modified>' +
  '</cp:coreProperties>'

const BODY = '<w:p><w:r><w:t>正文</w:t></w:r></w:p>'

const docxWithCore = () =>
  buildDocx({
    bodyXml: BODY,
    extraParts: [
      {
        path: 'docProps/core.xml',
        xml: CORE_XML,
        contentType: 'application/vnd.openxmlformats-package.core-properties+xml',
      },
    ],
  })

const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
  doc.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))

describe('docProps/core.xml on save', () => {
  it('a real save updates dcterms:modified and bumps cp:revision, keeping other fields', async () => {
    const doc = await parseDocx(await docxWithCore())
    const out = await saveDocx(doc, originalOrder(doc), {
      watermark: '草稿',
      savedAt: '2026-07-28T08:30:00Z',
    })
    const core = await (await JSZip.loadAsync(out)).file('docProps/core.xml')!.async('string')
    expect(core).toContain(
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-28T08:30:00Z</dcterms:modified>',
    )
    expect(core).toContain('<cp:revision>4</cp:revision>')
    expect(core).toContain('<dc:title>测试文档</dc:title>')
    expect(core).toContain('<dcterms:created xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:created>')
  })

  it('an unedited save still returns the original bytes untouched', async () => {
    const bytes = await docxWithCore()
    const doc = await parseDocx(bytes)
    const out = await saveDocx(doc, originalOrder(doc), { savedAt: '2026-07-28T08:30:00Z' })
    expect(out).toEqual(bytes)
  })

  it('a document without docProps saves fine (nothing injected)', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const out = await saveDocx(doc, originalOrder(doc), { watermark: '草稿' })
    expect((await JSZip.loadAsync(out)).file('docProps/core.xml')).toBeNull()
  })
})
