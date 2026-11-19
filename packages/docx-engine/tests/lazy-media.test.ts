import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { crc32, inflateRawSync } from 'node:zlib'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import {
  LAZY_MEDIA_PLACEHOLDER_BYTES,
  isLazyMediaPart,
  lazyMediaHashOf,
  lazyMediaPlaceholder,
  lazyMediaUrl,
  parseLazyMediaUrl,
} from '../src/lazy-media'
import {
  bufferSource,
  lazyMediaHashesIn,
  materializeDocx,
  openZipFile,
  readZipEntries,
  slimDocx,
  writeZip,
  type ZipFile,
} from '../src/zip-splice'
import { IMAGE_PARAGRAPH_XML, TINY_PNG_BASE64, buildDocx } from './helpers/build-docx'

const HASH = 'a'.repeat(64)
const PNG = Buffer.from(TINY_PNG_BASE64, 'base64')

async function tempZipFile(bytes: Uint8Array): Promise<{ file: ZipFile; path: string }> {
  const dir = await mkdtemp(join(tmpdir(), 'lazy-media-'))
  const path = join(dir, 'doc.docx')
  await writeFile(path, bytes)
  return { file: await openZipFile(path), path }
}

/** one STORE entry written with a data descriptor (bit 3), as streaming writers do */
function descriptorZip(name: string, data: Buffer): Buffer {
  const nameBytes = Buffer.from(name)
  const local = Buffer.alloc(30)
  local.writeUInt32LE(0x04034b50, 0)
  local.writeUInt16LE(20, 4)
  local.writeUInt16LE(0x8, 6)
  local.writeUInt16LE(0, 8)
  local.writeUInt16LE(nameBytes.length, 26)
  const descriptor = Buffer.alloc(16)
  descriptor.writeUInt32LE(0x08074b50, 0)
  descriptor.writeUInt32LE(crc32(data), 4)
  descriptor.writeUInt32LE(data.length, 8)
  descriptor.writeUInt32LE(data.length, 12)
  const central = Buffer.alloc(46)
  central.writeUInt32LE(0x02014b50, 0)
  central.writeUInt16LE(20, 4)
  central.writeUInt16LE(20, 6)
  central.writeUInt16LE(0x8, 8)
  central.writeUInt32LE(crc32(data), 16)
  central.writeUInt32LE(data.length, 20)
  central.writeUInt32LE(data.length, 24)
  central.writeUInt16LE(nameBytes.length, 28)
  central.writeUInt32LE(0, 42)
  const cdOffset = 30 + nameBytes.length + data.length + 16
  const eocd = Buffer.alloc(22)
  eocd.writeUInt32LE(0x06054b50, 0)
  eocd.writeUInt16LE(1, 8)
  eocd.writeUInt16LE(1, 10)
  eocd.writeUInt32LE(46 + nameBytes.length, 12)
  eocd.writeUInt32LE(cdOffset, 16)
  return Buffer.concat([local, nameBytes, data, descriptor, central, nameBytes, eocd])
}

describe('lazy media placeholders', () => {
  it('round-trips the hash and rejects other content', () => {
    const ph = lazyMediaPlaceholder(HASH)
    expect(ph.length).toBe(LAZY_MEDIA_PLACEHOLDER_BYTES)
    expect(lazyMediaHashOf(ph)).toBe(HASH)
    expect(lazyMediaHashOf(new Uint8Array(LAZY_MEDIA_PLACEHOLDER_BYTES))).toBeNull()
    expect(lazyMediaHashOf(PNG)).toBeNull()
    expect(() => lazyMediaPlaceholder('nope')).toThrow()
  })

  it('only browser-decodable word/media pictures are lazy', () => {
    expect(isLazyMediaPart('word/media/image1.png')).toBe(true)
    expect(isLazyMediaPart('word/media/Image 1.JPG')).toBe(false)
    expect(isLazyMediaPart('word/media/image1.emf')).toBe(false)
    expect(isLazyMediaPart('word/embeddings/oleObject1.bin')).toBe(false)
  })

  it('builds and parses URLs', () => {
    const url = lazyMediaUrl(HASH, 'word/media/image1.png')
    expect(url).toBe(`genoffice-docx-media://${HASH}/word/media/image1.png`)
    expect(parseLazyMediaUrl(url)).toEqual({ hash: HASH, partPath: 'word/media/image1.png' })
    expect(parseLazyMediaUrl('data:image/png;base64,AAAA')).toBeNull()
  })

  it('rejects bad hashes and traversal/overlong part paths', () => {
    expect(() => lazyMediaUrl('not-a-hash', 'word/media/image1.png')).toThrow(/sha256/)
    expect(parseLazyMediaUrl(`genoffice-docx-media://${HASH}/word/media/../evil.png`)).toBeNull()
    expect(
      parseLazyMediaUrl(`genoffice-docx-media://${HASH}/word/media/${'a'.repeat(600)}.png`),
    ).toBeNull()
  })
})

describe('zip splice', () => {
  it('reads entries and rewrites an archive byte-for-byte per part', async () => {
    const zip = new JSZip()
    zip.file('word/document.xml', '<w:document/>', { compression: 'DEFLATE' })
    zip.file('word/media/image1.png', PNG, { compression: 'STORE' })
    const bytes = Buffer.from(await zip.generateAsync({ type: 'uint8array' }))
    const src = bufferSource(bytes)
    const entries = await readZipEntries(src)
    expect(
      entries
        .map((e) => e.name)
        .filter((n) => !n.endsWith('/'))
        .sort(),
    ).toEqual(['word/document.xml', 'word/media/image1.png'])
    const out = writeZip(
      await Promise.all(
        entries.map(async (e) => ({ meta: e, data: await src.read(e.dataOffset, e.csize) })),
      ),
    )
    const reread = await JSZip.loadAsync(out)
    expect(await reread.file('word/document.xml')!.async('string')).toBe('<w:document/>')
    expect(Buffer.from(await reread.file('word/media/image1.png')!.async('uint8array'))).toEqual(
      PNG,
    )
  })

  it('folds data descriptors into the headers', async () => {
    const bytes = descriptorZip('word/media/image1.png', PNG)
    const src = bufferSource(bytes)
    const [entry] = await readZipEntries(src)
    expect(entry.csize).toBe(PNG.length)
    const out = writeZip([{ meta: entry, data: await src.read(entry.dataOffset, entry.csize) }])
    const [rewritten] = await readZipEntries(bufferSource(out))
    expect(rewritten.flags & 0x8).toBe(0)
    expect(
      Buffer.from(
        await (await JSZip.loadAsync(out)).file('word/media/image1.png')!.async('uint8array'),
      ),
    ).toEqual(PNG)
  })

  it('slims a document, serves parts from the file and materializes the save', async () => {
    const big = Buffer.alloc(4096, 7)
    const original = await buildDocx({
      bodyXml: IMAGE_PARAGRAPH_XML,
      withImage: true,
      binaryParts: [
        {
          path: 'word/media/image2.png',
          base64: big.toString('base64'),
          extension: 'png',
          contentType: 'image/png',
        },
        {
          path: 'word/media/image3.emf',
          base64: big.toString('base64'),
          extension: 'emf',
          contentType: 'image/x-emf',
        },
      ],
    })
    const { file, path } = await tempZipFile(original)
    try {
      expect(await slimDocx(file, HASH, 1 << 20)).toBeNull()
      const slim = await slimDocx(file, HASH, 1024)
      expect(slim).not.toBeNull()
      expect(slim!.lazyParts.sort()).toEqual(['word/media/image1.png', 'word/media/image2.png'])

      const slimZip = await JSZip.loadAsync(slim!.bytes)
      expect(
        lazyMediaHashOf(await slimZip.file('word/media/image1.png')!.async('uint8array')),
      ).toBe(HASH)
      expect((await slimZip.file('word/media/image3.emf')!.async('uint8array')).length).toBe(
        big.length,
      )
      expect(await lazyMediaHashesIn(slim!.bytes)).toEqual(new Set([HASH]))

      const part = file.entries.get('word/media/image2.png')!
      const served = await file.read(part.dataOffset, part.csize)
      expect(part.method === 8 ? inflateRawSync(served) : served).toEqual(big)

      const parsed = await parseDocx(slim!.bytes)
      const image = parsed.blocks.find((b) => b.type === 'image')!
      expect(image.imageDataUrl).toBe(lazyMediaUrl(HASH, 'word/media/image1.png'))
      expect(parsed.extras.lazyMediaHashes).toEqual([HASH])

      const originals = parsed.blocks.flatMap((b) =>
        b.docxIndex == null ? [] : [{ kind: 'original' as const, docxIndex: b.docxIndex }],
      )
      const sourceFor = async (hash: string) => (hash === HASH ? await openZipFile(path) : null)
      // untouched document: the save hands back the slim bytes themselves
      expect(
        await materializeDocx(Buffer.from(await saveDocx(parsed, originals, {})), sourceFor),
      ).not.toBe(slim!.bytes)
      const saved = Buffer.from(
        await saveDocx(
          parsed,
          [...originals, { kind: 'xml', xml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }],
          {},
        ),
      )
      const full = await materializeDocx(saved, sourceFor)
      const fullZip = await JSZip.loadAsync(full)
      expect(Buffer.from(await fullZip.file('word/media/image1.png')!.async('uint8array'))).toEqual(
        PNG,
      )
      expect(Buffer.from(await fullZip.file('word/media/image2.png')!.async('uint8array'))).toEqual(
        big,
      )
      expect(await fullZip.file('word/document.xml')!.async('string')).toContain('<w:t>x</w:t>')

      const untouched = Buffer.from(original)
      expect(await materializeDocx(untouched, sourceFor)).toBe(untouched)
      await expect(materializeDocx(saved, async () => null)).rejects.toThrow(/source unavailable/)
    } finally {
      await file.close()
    }
  })

  it('reuses a lazy part for an image saved by part reference', async () => {
    const original = await buildDocx({ bodyXml: IMAGE_PARAGRAPH_XML, withImage: true })
    const parsed = await parseDocx(original)
    const saved = await saveDocx(
      parsed,
      [
        {
          kind: 'image',
          image: {
            base64: '',
            mime: 'image/png',
            sourcePart: 'word/media/image1.png',
            widthPx: 100,
            heightPx: 50,
          },
        },
      ],
      {},
    )
    const zip = await JSZip.loadAsync(saved)
    expect(
      Object.keys(zip.files).filter((n) => n.startsWith('word/media/') && !n.endsWith('/')),
    ).toEqual(['word/media/image1.png'])
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels.match(/Target="media\/image1\.png"/g)?.length).toBeGreaterThanOrEqual(1)
    const doc = await zip.file('word/document.xml')!.async('string')
    expect(doc).toMatch(/r:embed="rId\d+"/)
  })

  it('rejects zip64 and encrypted archives', async () => {
    const zip = new JSZip()
    zip.file('a.txt', 'a')
    const bytes = Buffer.from(await zip.generateAsync({ type: 'uint8array' }))
    const encrypted = Buffer.from(bytes)
    const cdOffset = encrypted.readUInt32LE(encrypted.length - 22 + 16)
    encrypted.writeUInt16LE(encrypted.readUInt16LE(cdOffset + 8) | 0x1, cdOffset + 8)
    await expect(readZipEntries(bufferSource(encrypted))).rejects.toThrow(/encrypted/)
    const zip64 = Buffer.from(bytes)
    zip64.writeUInt16LE(0xffff, zip64.length - 22 + 10)
    await expect(readZipEntries(bufferSource(zip64))).rejects.toThrow(/zip64/)
  })
})
