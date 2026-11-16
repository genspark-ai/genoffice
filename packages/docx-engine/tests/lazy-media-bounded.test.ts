import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { DEFAULT_ZIP_LIMITS } from '@genoffice/zip-gate'
import { MAX_LAZY_SERVED_MEDIA_BYTES, readMediaPartBounded } from '../src/lazy-media'
import { bufferSource, readZipEntries, type ZipEntry } from '../src/zip-splice'

const KB = 1024

const MEDIA = 'word/media/image1.png'

/** A docx archive holding one media part (deflated by default), as bytes. */
async function archiveWithMedia(content: Buffer, opts?: { store?: boolean }): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('word/document.xml', '<w:p/>')
  zip.file(MEDIA, content, opts?.store ? { compression: 'STORE' } : undefined)
  return (await zip.generateAsync({ type: 'nodebuffer' })) as Buffer
}

/** Overwrite one header word of the media part's central-directory record. */
function patchMediaHeader(bytes: Buffer, field: 'csize' | 'usize', value: number): Buffer {
  const out = Buffer.from(bytes)
  const centralName = out.lastIndexOf(Buffer.from(MEDIA))
  out.writeUInt32LE(value, centralName - 46 + (field === 'csize' ? 20 : 24))
  return out
}

async function mediaEntryOf(
  bytes: Buffer,
): Promise<{ entry: ZipEntry; read: (offset: number, length: number) => Promise<Buffer> }> {
  const src = bufferSource(bytes)
  const entries = await readZipEntries(src)
  const entry = entries.find((e) => e.name === MEDIA)
  if (!entry) throw new Error('media part missing from archive')
  return { entry, read: (offset, length) => src.read(offset, length) }
}

describe('readMediaPartBounded', () => {
  it('serves an honest deflated part', async () => {
    const content = Buffer.from([0x89, 0x50, 0x4e, 0x47, ...Buffer.alloc(16 * KB)])
    const bytes = await archiveWithMedia(content)
    const { entry, read } = await mediaEntryOf(bytes)
    const out = await readMediaPartBounded(entry, read, 64 * KB)
    expect(out).not.toBeNull()
    expect(Buffer.from(out!).equals(content)).toBe(true)
  })

  it('serves an honest stored part as-is', async () => {
    const content = Buffer.alloc(4 * KB, 0xab)
    const bytes = await archiveWithMedia(content, { store: true })
    const { entry, read } = await mediaEntryOf(bytes)
    expect(entry.method).toBe(0)
    const out = await readMediaPartBounded(entry, read, 64 * KB)
    expect(Buffer.from(out!).equals(content)).toBe(true)
  })

  it('catches the serving-side bomb at the inflate: header within the cap, payload past it', async () => {
    // The review's measured attack, shrunk to a test budget: 1 MB of real
    // payload declaring 60 KB. The lazy path never gate-checks declarations
    // in bulk, so the inflate itself has to be the guard.
    const bytes = patchMediaHeader(
      await archiveWithMedia(Buffer.alloc(1024 * KB)),
      'usize',
      60 * KB,
    )
    const { entry, read } = await mediaEntryOf(bytes)
    expect(entry.usize).toBe(60 * KB)
    const out = await readMediaPartBounded(entry, read, 64 * KB)
    expect(out).toBeNull()
  })

  it('returns null without reading when a header exceeds the cap', async () => {
    let readCalls = 0
    const spy = async (offset: number, length: number): Promise<Buffer> => {
      readCalls += 1
      return Buffer.alloc(length)
    }
    const base = { dataOffset: 0, method: 8 }
    await expect(
      readMediaPartBounded({ ...base, csize: 128 * KB + 1, usize: 10 }, spy, 128 * KB),
    ).resolves.toBeNull()
    await expect(
      readMediaPartBounded({ ...base, csize: 10, usize: 128 * KB + 1 }, spy, 128 * KB),
    ).resolves.toBeNull()
    expect(readCalls).toBe(0)
  })

  it('holds on a hostile archive whose declared usize passes but payload lies', async () => {
    // End-to-end through the real reader: patch usize upward past the cap on
    // a real archive (usize is not bounds-checked by readZipEntries) and the
    // header check must refuse before any read happens.
    const bytes = patchMediaHeader(
      await archiveWithMedia(Buffer.alloc(2 * KB)),
      'usize',
      128 * KB + 1,
    )
    const { entry, read } = await mediaEntryOf(bytes)
    let readCalls = 0
    const spy = async (offset: number, length: number): Promise<Buffer> => {
      readCalls += 1
      return read(offset, length)
    }
    const out = await readMediaPartBounded(entry, spy, 128 * KB)
    expect(out).toBeNull()
    expect(readCalls).toBe(0)
  })

  it('returns null for an unsupported compression method', async () => {
    const bytes = await archiveWithMedia(Buffer.alloc(2 * KB))
    const { entry, read } = await mediaEntryOf(bytes)
    const out = await readMediaPartBounded({ ...entry, method: 12 }, read, 64 * KB)
    expect(out).toBeNull()
  })

  it('keeps the serving cap aligned with the bulk gate per-part cap', () => {
    expect(MAX_LAZY_SERVED_MEDIA_BYTES).toBe(DEFAULT_ZIP_LIMITS.maxPartBytes)
  })
})
