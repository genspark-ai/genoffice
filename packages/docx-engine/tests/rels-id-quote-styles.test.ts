import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { findChartWorkbookPath, parseDocx, saveDocx } from '../src/index'
import { TINY_PNG_BASE64, buildDocx } from './helpers/build-docx'

/**
 * A .rels part is XML, so nothing obliges a writer to spell its attributes
 * with double quotes: plenty of producers (python-docx hand edits, lxml, some
 * Java libraries) emit Id='rId7' or Id = "rId7". Every reader that allocates a
 * relationship id has to accept every spelling, or it hands out an id that is
 * already in use and the second relationship silently wins.
 */

const BODY = '<w:p><w:r><w:t>body</w:t></w:r></w:p>'
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const OLD_TARGET = 'https://old.example.com/'
const NEW_TARGET = 'https://new.example.com/'
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'

const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>) =>
  doc.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))

/** Re-spell every relationship attribute with single quotes. */
const singleQuoted = (xml: string): string =>
  xml.replace(
    /\b(Id|Type|Target|TargetMode)="([^"]*)"/g,
    (_m, name: string, value: string) => `${name}='${value}'`,
  )

/** Re-spell every Id with whitespace padding around the '='. */
const paddedIds = (xml: string): string =>
  xml.replace(/\bId="([^"]*)"/g, (_m, value: string) => `Id = "${value}"`)

async function mapPart(
  bytes: Uint8Array,
  path: string,
  fn: (xml: string) => string,
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  const file = zip.file(path)
  if (!file) throw new Error(`missing part ${path}`)
  zip.file(path, fn(await file.async('string')))
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

const putPart = async (bytes: Uint8Array, path: string, xml: string): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(bytes)
  zip.file(path, xml)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

async function zipText(bytes: Uint8Array, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  const file = zip.file(path)
  if (!file) throw new Error(`missing part ${path}`)
  return file.async('string')
}

/** Every relationship id in a rels part, in whichever quote style it is spelled. */
const relIds = (relsXml: string): string[] =>
  [...relsXml.matchAll(/\bId\s*=\s*(["'])(.*?)\1/g)].map((m) => m[2])

const relTag = (relsXml: string, id: string): string | undefined =>
  (relsXml.match(/<Relationship\b[^>]*>/g) ?? []).find((tag) =>
    new RegExp(`\\bId\\s*=\\s*(["'])${id}\\1`).test(tag),
  )

const relTarget = (relsXml: string, id: string): string | undefined =>
  /\bTarget\s*=\s*(["'])(.*?)\1/.exec(relTag(relsXml, id) ?? '')?.[2]

/** ids must be unique: a duplicate makes the second relationship unreachable */
const expectUniqueIds = (relsXml: string): void => {
  const ids = relIds(relsXml)
  expect(ids.length).toBeGreaterThan(0)
  expect(new Set(ids).size).toBe(ids.length)
}

const imageRels = (relsXml: string): string[] =>
  (relsXml.match(/<Relationship\b[^>]*>/g) ?? []).filter((t) => t.includes('/image'))

/**
 * Keep every original block and append one generated paragraph carrying a new
 * link, so the document still references the relationship it started with.
 */
const saveWithLink = (doc: Awaited<ReturnType<typeof parseDocx>>) =>
  saveDocx(doc, [
    ...originalOrder(doc),
    {
      kind: 'generated',
      block: { type: 'paragraph', runs: [{ text: 'link', link: { href: NEW_TARGET } }] },
    },
  ])

/** a document whose only hyperlink already points somewhere, on rId20 */
const linkDoc = async () =>
  mapPart(
    await buildDocx({
      bodyXml: `<w:p><w:hyperlink r:id="rId20"><w:r><w:t>old</w:t></w:r></w:hyperlink></w:p>`,
      extraRels:
        `<Relationship Id="rId20" Type="${REL_BASE}/hyperlink" ` +
        `Target="${OLD_TARGET}" TargetMode="External"/>`,
    }),
    'word/_rels/document.xml.rels',
    singleQuoted,
  )

describe('relationship ids in either quote style', () => {
  it('does not reissue an id the document already spells with single quotes', async () => {
    const saved = await saveWithLink(await parseDocx(await linkDoc()))
    const relsXml = await zipText(saved, 'word/_rels/document.xml.rels')
    expectUniqueIds(relsXml)
  })

  it('points the new hyperlink at its own target, not the older relationship', async () => {
    const saved = await saveWithLink(await parseDocx(await linkDoc()))
    const relsXml = await zipText(saved, 'word/_rels/document.xml.rels')
    const docXml = await zipText(saved, 'word/document.xml')
    const newRId = [...docXml.matchAll(/<w:hyperlink r:id="([^"]+)"/g)].map((m) => m[1]).at(-1)!
    expect(newRId).toBeTruthy()
    expect(relTarget(relsXml, newRId)).toBe(NEW_TARGET)
    // the pre-existing link is untouched
    expect(relTarget(relsXml, 'rId20')).toBe(OLD_TARGET)
  })

  it('does not reissue an id spelled with whitespace around the =', async () => {
    const bytes = await mapPart(
      await buildDocx({ bodyXml: BODY }),
      'word/_rels/document.xml.rels',
      paddedIds,
    )
    const saved = await saveWithLink(await parseDocx(bytes))
    expectUniqueIds(await zipText(saved, 'word/_rels/document.xml.rels'))
  })

  it('reclaims and reissues a picture-watermark id in a single-quoted header rels', async () => {
    const image = { base64: TINY_PNG_BASE64, mime: 'image/png' as const, widthPx: 96, heightPx: 48 }
    const source = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const first = await saveDocx(source, originalOrder(source), { watermark: { image } })
    // both halves of the pair single-quoted: the header's own r:id reference and
    // the relationship it points at
    const requoted = await mapPart(
      await mapPart(first, 'word/_rels/header1.xml.rels', singleQuoted),
      'word/header1.xml',
      (xml) => xml.replace(/\br:id="([^"]*)"/g, (_m, v: string) => `r:id='${v}'`),
    )
    const doc = await parseDocx(requoted)
    const saved = await saveDocx(doc, originalOrder(doc), { watermark: { image } })
    const relsXml = await zipText(saved, 'word/_rels/header1.xml.rels')
    expectUniqueIds(relsXml)
    // the superseded relationship is reclaimed, not left behind alongside the
    // new one - two image relationships in one part means one of them is dead
    expect(imageRels(relsXml)).toHaveLength(1)
    const headerXml = await zipText(saved, 'word/header1.xml')
    const wmRId = /\br:id=(["'])(.*?)\1/.exec(headerXml)![2]
    expect(wmRId).toBeTruthy()
    expect(relTarget(relsXml, wmRId)).toBe('media/aidocs2.png')
  })

  it('keeps a single-quoted header relationship addressable when editing its text', async () => {
    const bytes = await buildDocx({
      bodyXml: BODY,
      extraRels: `<Relationship Id='rId88' Type="${REL_BASE}/header" Target="header1.xml"/>`,
      extraParts: [
        {
          path: 'word/header1.xml',
          xml:
            `${XML_DECL}<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
            '<w:p><w:r><w:t>BEFORE</w:t></w:r></w:p></w:hdr>',
          contentType: HEADER_TYPE,
        },
      ],
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId88"/>',
    })
    const saved = await saveDocx(await parseDocx(bytes), [{ kind: 'original', docxIndex: 0 }], {
      header: { text: 'AFTER' },
    })
    expect(await zipText(saved, 'word/header1.xml')).toContain('AFTER')
    expectUniqueIds(await zipText(saved, 'word/_rels/document.xml.rels'))
    const zip = await JSZip.loadAsync(saved)
    expect(Object.keys(zip.files).filter((n) => /^word\/header\d+\.xml$/.test(n))).toEqual([
      'word/header1.xml',
    ])
    expect((await zipText(saved, 'word/document.xml')).match(/<w:headerReference/g)).toHaveLength(1)
  })

  it('allocates an image id above an existing single-quoted one', async () => {
    const bytes = await buildDocx({
      bodyXml: BODY,
      extraRels: `<Relationship Id='rId2' Type="${REL_BASE}/hyperlink" Target="${OLD_TARGET}" TargetMode="External"/>`,
    })
    const doc = await parseDocx(bytes)
    const saved = await saveDocx(doc, [
      ...originalOrder(doc),
      {
        kind: 'image',
        image: { base64: TINY_PNG_BASE64, mime: 'image/png', widthPx: 100, heightPx: 60 },
      },
    ])
    const relsXml = await zipText(saved, 'word/_rels/document.xml.rels')
    expectUniqueIds(relsXml)
    expect(relIds(relsXml)).toContain('rId3')
  })

  it('allocates a numbering picture-bullet id above an existing single-quoted one', async () => {
    const bytes = await putPart(
      await buildDocx({ bodyXml: BODY, withNumbering: true }),
      'word/_rels/numbering.xml.rels',
      `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id='rId1' Type="${REL_BASE}/image" Target="media/image1.png"/></Relationships>`,
    )
    const saved = await saveDocx(await parseDocx(bytes), [{ kind: 'original', docxIndex: 0 }], {
      numbering: { picBullets: [{ id: 0, base64: TINY_PNG_BASE64, mime: 'image/png' }] },
    })
    const relsXml = await zipText(saved, 'word/_rels/numbering.xml.rels')
    expectUniqueIds(relsXml)
    expect(relTarget(relsXml, 'rId1')).toBe('media/image1.png')
  })

  it('finds a chart workbook relationship however its attributes are spelled', async () => {
    const bytes = await putPart(
      await buildDocx({ bodyXml: BODY }),
      'word/charts/_rels/chart1.xml.rels',
      `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id='rId1' Type='${REL_BASE}/package' Target='embeddings/workbook1.xlsx'/>` +
        '</Relationships>',
    )
    expect(await findChartWorkbookPath(bytes, 'word/charts/chart1.xml')).toBe(
      'word/charts/embeddings/workbook1.xlsx',
    )
  })

  it('finds a single-quoted absolute chart workbook target written before its Type', async () => {
    const bytes = await putPart(
      await buildDocx({ bodyXml: BODY }),
      'word/charts/_rels/chart1.xml.rels',
      `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Target='/word/embeddings/workbook%201.xlsx' Id='rId1' Type='${REL_BASE}/package'/>` +
        '</Relationships>',
    )
    expect(await findChartWorkbookPath(bytes, 'word/charts/chart1.xml')).toBe(
      'word/embeddings/workbook 1.xlsx',
    )
  })

  it('finds a chart workbook relationship whose Target precedes its Type', async () => {
    const bytes = await putPart(
      await buildDocx({ bodyXml: BODY }),
      'word/charts/_rels/chart1.xml.rels',
      `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Target="embeddings/workbook1.xlsx" Type="${REL_BASE}/package"/>` +
        '</Relationships>',
    )
    expect(await findChartWorkbookPath(bytes, 'word/charts/chart1.xml')).toBe(
      'word/charts/embeddings/workbook1.xlsx',
    )
  })
})
