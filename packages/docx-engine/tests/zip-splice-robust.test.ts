import { crc32 } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { lazyMediaPlaceholder } from '../src/lazy-media'
import {
  bufferSource,
  materializeDocx,
  readZipEntries,
  writeZip,
  type ZipFile,
} from '../src/zip-splice'

function storeEntry(name: string, data: Buffer) {
  return {
    meta: {
      name,
      nameBytes: Buffer.from(name),
      flags: 0,
      method: 0,
      time: 0,
      date: 0,
      crc: crc32(data),
      csize: data.length,
      usize: data.length,
      verMade: 20,
      verNeed: 20,
      intAttr: 0,
      extAttr: 0,
    },
    data,
  }
}

function eocdPos(bytes: Buffer): number {
  return bytes.length - 22
}

async function trackableZip(bytes: Buffer, onClose: () => void): Promise<ZipFile> {
  const src = bufferSource(bytes)
  const entries = new Map()
  for (const e of await readZipEntries(src)) entries.set(e.name, e)
  return {
    size: bytes.length,
    read: (offset, length) => src.read(offset, length),
    entries,
    close: async () => {
      onClose()
    },
  }
}

describe('zip splice robustness', () => {
  it('rejects a truncated archive with a descriptive error', async () => {
    const valid = writeZip([
      storeEntry('word/document.xml', Buffer.from('<w:document/>')),
      storeEntry('word/media/image1.png', Buffer.from('original-image-bytes')),
    ])
    const truncated = valid.subarray(0, valid.length - 10)
    let error: unknown = null
    try {
      await readZipEntries(bufferSource(Buffer.from(truncated)))
    } catch (err) {
      error = err
    }
    expect(error).toBeInstanceOf(Error)
    expect(error).not.toBeInstanceOf(RangeError)
    expect((error as Error).message).toMatch(/zip: /)

    const patched = Buffer.from(valid)
    const pos = eocdPos(patched)
    const cdSize = patched.readUInt32LE(pos + 12)
    patched.writeUInt32LE(cdSize + 50, pos + 12)
    let claimed: unknown = null
    try {
      await readZipEntries(bufferSource(patched))
    } catch (err) {
      claimed = err
    }
    expect(claimed).toBeInstanceOf(Error)
    expect(claimed).not.toBeInstanceOf(RangeError)
    expect((claimed as Error).message).toMatch(/corrupt central directory/)
  })

  it('rejects a central directory outside the file bounds', async () => {
    const valid = writeZip([storeEntry('a.txt', Buffer.from('a'))])
    const outsideOffset = Buffer.from(valid)
    const pos = eocdPos(outsideOffset)
    outsideOffset.writeUInt32LE(outsideOffset.length + 100, pos + 16)
    await expect(readZipEntries(bufferSource(outsideOffset))).rejects.toThrow(
      /corrupt central directory/,
    )

    const outsideSize = Buffer.from(valid)
    const pos2 = eocdPos(outsideSize)
    outsideSize.writeUInt32LE(outsideSize.length + 100, pos2 + 12)
    await expect(readZipEntries(bufferSource(outsideSize))).rejects.toThrow(
      /corrupt central directory/,
    )
  })

  it('rejects entries whose declared data lies outside the archive', async () => {
    const valid = writeZip([storeEntry('word/document.xml', Buffer.from('<w:document/>'))])
    const CENTRAL_SIG = Buffer.from([0x50, 0x4b, 0x01, 0x02])

    // readers allocate from the declared csize before touching the file; an
    // >int32 length additionally aborts Node outright inside fs.read
    const oversized = Buffer.from(valid)
    const c1 = oversized.indexOf(CENTRAL_SIG)
    oversized.writeUInt32LE(0xffffff00, c1 + 20)
    await expect(readZipEntries(bufferSource(oversized))).rejects.toThrow(
      /outside the .*-byte archive/,
    )

    // so must a csize that merely overshoots the end of the file
    const overhang = Buffer.from(valid)
    const c2 = overhang.indexOf(CENTRAL_SIG)
    overhang.writeUInt32LE(overhang.length, c2 + 20)
    await expect(readZipEntries(bufferSource(overhang))).rejects.toThrow(
      /outside the .*-byte archive/,
    )

    // the unmodified archive still parses
    expect((await readZipEntries(bufferSource(Buffer.from(valid)))).length).toBe(1)
  })

  it('closes materialize sources on success and lets the caller use the result', async () => {
    const hash = 'c'.repeat(64)
    const placeholder = Buffer.from(lazyMediaPlaceholder(hash))
    const original = Buffer.from('original-image-bytes')
    const docXml = Buffer.from('<w:document/>')
    const slim = writeZip([
      storeEntry('word/document.xml', docXml),
      {
        meta: {
          name: 'word/media/image1.png',
          nameBytes: Buffer.from('word/media/image1.png'),
          flags: 0,
          method: 0,
          time: 0,
          date: 0,
          crc: crc32(placeholder),
          csize: placeholder.length,
          usize: placeholder.length,
          verMade: 20,
          verNeed: 20,
          intAttr: 0,
          extAttr: 0,
        },
        data: placeholder,
      },
    ])
    const sourceBytes = writeZip([storeEntry('word/media/image1.png', original)])

    let closes = 0
    let calls = 0
    const sourceFor = async (wanted: string) => {
      expect(wanted).toBe(hash)
      calls += 1
      return trackableZip(sourceBytes, () => {
        closes += 1
      })
    }

    const result = await materializeDocx(Buffer.from(slim), sourceFor)
    expect(calls).toBe(1)
    expect(closes).toBe(1)

    const src = bufferSource(result)
    const entries = await readZipEntries(src)
    const restored = entries.find((e) => e.name === 'word/media/image1.png')
    expect(restored).toBeDefined()
    const data = await src.read(restored!.dataOffset, restored!.csize)
    expect(Buffer.from(data)).toEqual(original)
  })

  it('caches one source per hash and closes error paths', async () => {
    const hash = 'c'.repeat(64)
    const placeholder = Buffer.from(lazyMediaPlaceholder(hash))
    const original = Buffer.from('original-image-bytes')
    const placeholderMeta = (name: string) => ({
      name,
      nameBytes: Buffer.from(name),
      flags: 0,
      method: 0,
      time: 0,
      date: 0,
      crc: crc32(placeholder),
      csize: placeholder.length,
      usize: placeholder.length,
      verMade: 20,
      verNeed: 20,
      intAttr: 0,
      extAttr: 0,
    })
    const slim = writeZip([
      { meta: placeholderMeta('word/media/image1.png'), data: placeholder },
      { meta: placeholderMeta('word/media/image2.png'), data: placeholder },
    ])
    const sourceBytes = writeZip([
      storeEntry('word/media/image1.png', original),
      storeEntry('word/media/image2.png', original),
    ])

    let closes = 0
    let calls = 0
    const sourceFor = async () => {
      calls += 1
      return trackableZip(sourceBytes, () => {
        closes += 1
      })
    }
    const result = await materializeDocx(Buffer.from(slim), sourceFor)
    expect(calls).toBe(1)
    expect(closes).toBe(1)
    expect((await readZipEntries(bufferSource(result))).length).toBe(2)

    const otherHash = 'd'.repeat(64)
    const otherPlaceholder = Buffer.from(lazyMediaPlaceholder(otherHash))
    const mixed = writeZip([
      { meta: placeholderMeta('word/media/image1.png'), data: placeholder },
      {
        meta: {
          ...placeholderMeta('word/media/image2.png'),
          crc: crc32(otherPlaceholder),
        },
        data: otherPlaceholder,
      },
    ])
    let errorCloses = 0
    const failingSourceFor = async (wanted: string) => {
      if (wanted !== hash) return null
      return trackableZip(sourceBytes, () => {
        errorCloses += 1
      })
    }
    await expect(materializeDocx(Buffer.from(mixed), failingSourceFor)).rejects.toThrow(
      /source unavailable/,
    )
    expect(errorCloses).toBe(1)
  })
})
