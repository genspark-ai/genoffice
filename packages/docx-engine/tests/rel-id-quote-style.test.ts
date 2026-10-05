import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { findChartWorkbookPath, parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'

const P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const HEADER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const HEADER_PART_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  '<w:p><w:r><w:t>Original header</w:t></w:r></w:p></w:hdr>'

/**
 * Every relationship id in a .rels part, read the way a consumer would:
 * either quote style, any spacing around `=`. Used to assert the invariant
 * (ids are unique) independently of how the engine happens to spell them.
 */
function relIds(relsXml: string): string[] {
  return [...relsXml.matchAll(/\bId\s*=\s*(["'])([^"']+)\1/g)].map((m) => m[2])
}

async function visibleBlocks(bytes: Uint8Array): Promise<{
  doc: Awaited<ReturnType<typeof parseDocx>>
  blocks: SaveBlock[]
}> {
  const doc = await parseDocx(bytes)
  const blocks: SaveBlock[] = doc.blocks
    .filter((b) => !b.hidden)
    .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
  return { doc, blocks }
}

describe('relationship id allocation across quote styles', () => {
  it('allocates above a single-quoted existing id instead of duplicating it', async () => {
    const bytes = await buildDocx({
      bodyXml: P('hello'),
      extraRels: `<Relationship Id='rId2' Type="${REL}/hyperlink" Target="https://example.com/" TargetMode="External"/>`,
    })
    const { doc, blocks } = await visibleBlocks(bytes)
    const saved = await saveDocx(doc, [
      ...blocks,
      {
        kind: 'image',
        image: { base64: TINY_PNG_BASE64, mime: 'image/png', widthPx: 100, heightPx: 60 },
      },
    ])
    const zip = await JSZip.loadAsync(saved)
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    const ids = relIds(rels)
    expect(new Set(ids).size).toBe(ids.length)
    expect(ids).toContain('rId3')
  })

  it('allocates above an existing id spelled with spaces around the equals sign', async () => {
    const bytes = await buildDocx({
      bodyXml: P('hello'),
      extraRels: `<Relationship Id = "rId2" Type="${REL}/hyperlink" Target="https://example.com/" TargetMode="External"/>`,
    })
    const { doc, blocks } = await visibleBlocks(bytes)
    const saved = await saveDocx(doc, [
      ...blocks,
      {
        kind: 'image',
        image: { base64: TINY_PNG_BASE64, mime: 'image/png', widthPx: 100, heightPx: 60 },
      },
    ])
    const zip = await JSZip.loadAsync(saved)
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    const ids = relIds(rels)
    expect(new Set(ids).size).toBe(ids.length)
    expect(ids).toContain('rId3')
  })

  it('rewrites the existing header part when its relationship id is single-quoted', async () => {
    const bytes = await buildDocx({
      bodyXml: P('hello'),
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId2"/>',
      extraRels: `<Relationship Id='rId2' Type="${REL}/header" Target="header1.xml"/>`,
      extraParts: [
        { path: 'word/header1.xml', xml: HEADER_PART_XML, contentType: HEADER_CONTENT_TYPE },
      ],
    })
    const { doc, blocks } = await visibleBlocks(bytes)
    const saved = await saveDocx(doc, blocks, { header: { text: 'Updated header' } })
    const zip = await JSZip.loadAsync(saved)
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    const ids = relIds(rels)
    expect(new Set(ids).size).toBe(ids.length)
    // the referenced part is patched in place, so no second header part appears
    expect(Object.keys(zip.files).filter((n) => /^word\/header\d+\.xml$/.test(n))).toEqual([
      'word/header1.xml',
    ])
    const headerXml = await zip.file('word/header1.xml')!.async('string')
    expect(headerXml).toContain('Updated header')
    const documentXml = await zip.file('word/document.xml')!.async('string')
    expect(documentXml.match(/<w:headerReference/g) ?? []).toHaveLength(1)
  })
})

describe('findChartWorkbookPath across quote styles and attribute order', () => {
  it('reads a single-quoted target written before its Type attribute', async () => {
    const zip = await JSZip.loadAsync(await buildDocx({ bodyXml: '<w:p/>' }))
    zip.file(
      'word/charts/_rels/chart1.xml.rels',
      '<Relationships>' +
        `<Relationship Target='/word/embeddings/workbook%201.xlsx' Id='rId1' Type='${REL}/package'/>` +
        '</Relationships>',
    )
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    expect(await findChartWorkbookPath(bytes, 'word/charts/chart1.xml')).toBe(
      'word/embeddings/workbook 1.xlsx',
    )
  })
})
