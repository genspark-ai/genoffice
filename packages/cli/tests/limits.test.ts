import { readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { readZipDirectory, MAX_COMPRESSION_RATIO } from '../src/limits'
import { run, tempDir } from './helpers'

const REPO = resolve(__dirname, '../../..')
const DOCX = join(REPO, 'apps/docs/tests/pagination-corpus/docx/01-simple-english.docx')

interface Declared {
  name: string
  compressed?: number
  uncompressed?: number
}

/** A zip whose central directory declares whatever sizes the test wants; entries carry no data. */
function rawZip(entries: Declared[], zip64 = false): Buffer {
  const locals: Buffer[] = []
  const centrals: Buffer[] = []
  let offset = 0
  for (const e of entries) {
    const name = Buffer.from(e.name, 'utf8')
    const local = Buffer.alloc(30 + name.length)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(name.length, 26)
    name.copy(local, 30)
    const central = Buffer.alloc(46 + name.length)
    central.writeUInt32LE(0x02014b50, 0)
    central.writeUInt32LE(e.compressed ?? 0, 20)
    central.writeUInt32LE(e.uncompressed ?? 0, 24)
    central.writeUInt16LE(name.length, 28)
    central.writeUInt32LE(offset, 42)
    name.copy(central, 46)
    locals.push(local)
    centrals.push(central)
    offset += local.length
  }
  const cd = Buffer.concat(centrals)
  const eocd = Buffer.alloc(22)
  eocd.writeUInt32LE(0x06054b50, 0)
  if (zip64) {
    const rec = Buffer.alloc(56)
    rec.writeUInt32LE(0x06064b50, 0)
    rec.writeBigUInt64LE(44n, 4)
    rec.writeBigUInt64LE(BigInt(entries.length), 24)
    rec.writeBigUInt64LE(BigInt(entries.length), 32)
    rec.writeBigUInt64LE(BigInt(cd.length), 40)
    rec.writeBigUInt64LE(BigInt(offset), 48)
    const loc = Buffer.alloc(20)
    loc.writeUInt32LE(0x07064b50, 0)
    loc.writeBigUInt64LE(BigInt(offset + cd.length), 8)
    loc.writeUInt32LE(1, 16)
    eocd.writeUInt16LE(0xffff, 8)
    eocd.writeUInt16LE(0xffff, 10)
    eocd.writeUInt32LE(0xffffffff, 12)
    eocd.writeUInt32LE(0xffffffff, 16)
    return Buffer.concat([...locals, cd, rec, loc, eocd])
  }
  eocd.writeUInt16LE(entries.length, 8)
  eocd.writeUInt16LE(entries.length, 10)
  eocd.writeUInt32LE(cd.length, 12)
  eocd.writeUInt32LE(offset, 16)
  return Buffer.concat([...locals, cd, eocd])
}

function file(name: string, bytes: Buffer): string {
  const path = join(tempDir(), name)
  writeFileSync(path, bytes)
  return path
}

const MB = 1024 * 1024

describe('package limits', () => {
  it('reads declared sizes from the central directory, zip64 included', () => {
    const entries = [
      { name: 'word/document.xml', compressed: 10, uncompressed: 100 },
      { name: 'word/media/', compressed: 0, uncompressed: 0 },
    ]
    const expected = {
      count: 2,
      entries: [{ name: 'word/document.xml', compressed: 10, uncompressed: 100 }],
    }
    expect(readZipDirectory(file('a.docx', rawZip(entries)))).toEqual(expected)
    expect(readZipDirectory(file('b.docx', rawZip(entries, true)))).toEqual(expected)
    expect(readZipDirectory(file('c.docx', Buffer.from('not a zip')))).toBeNull()
  })

  it('refuses entries that escape the package', async () => {
    const path = file(
      'escape.docx',
      rawZip([{ name: '../evil.xml' }, { name: 'word/document.xml' }]),
    )
    const r = (await run(['info', path, '--json'])).json()
    expect(r).toMatchObject({ status: 'error', code: 2, error: 'resource_limit' })
    expect(r.detail.entry).toBe('../evil.xml')
  })

  it('refuses too many parts without reading the directory', async () => {
    const entries = Array.from({ length: 10001 }, (_, i) => ({ name: `p/${i}.xml` }))
    const many = file('many.xlsx', rawZip(entries))
    expect(readZipDirectory(many, 10000)).toEqual({ count: 10001, entries: [] })
    const r = (await run(['sheet', 'read', many, '--json'])).json()
    expect(r).toMatchObject({ error: 'resource_limit' })
    expect(r.detail).toMatchObject({ entries: 10001, limit: 10000 })

    const hostile = rawZip([{ name: 'a.xml' }], true)
    hostile.writeBigUInt64LE(50_000_000n, hostile.length - 22 - 20 - 56 + 32)
    expect(readZipDirectory(file('hostile.pptx', hostile), 10000)).toEqual({
      count: 50_000_000,
      entries: [],
    })
  })

  it('ignores a zip64 locator whose record is missing and keeps checking', async () => {
    const bomb = rawZip([{ name: 'word/document.xml', compressed: 1024, uncompressed: 200 * MB }])
    const fakeLocator = Buffer.alloc(20)
    fakeLocator.writeUInt32LE(0x07064b50, 0)
    fakeLocator.writeBigUInt64LE(BigInt(bomb.length + 4096), 8)
    const crafted = Buffer.concat([
      bomb.subarray(0, bomb.length - 22),
      fakeLocator,
      bomb.subarray(-22),
    ])
    const r = (await run(['docs', 'read', file('fake64.docx', crafted), '--json'])).json()
    expect(r).toMatchObject({ error: 'resource_limit' })
  })

  it('ignores a zip64 locator when the end record has no overflow sentinel (JSZip rule)', async () => {
    // A valid zip64 record pointing at an innocent one-entry directory sits in front of a
    // plain end record that still points at the bomb: the engine reads the bomb, so must we.
    const bomb = rawZip([{ name: 'word/document.xml', compressed: 1024, uncompressed: 200 * MB }])
    const body = bomb.subarray(0, bomb.length - 22)
    const innocent = rawZip([{ name: 'a.xml' }])
    const innocentCd = innocent.subarray(35, 35 + 51)
    const rec = Buffer.alloc(56)
    rec.writeUInt32LE(0x06064b50, 0)
    rec.writeBigUInt64LE(44n, 4)
    rec.writeBigUInt64LE(1n, 24)
    rec.writeBigUInt64LE(1n, 32)
    rec.writeBigUInt64LE(BigInt(innocentCd.length), 40)
    rec.writeBigUInt64LE(BigInt(body.length), 48)
    const loc = Buffer.alloc(20)
    loc.writeUInt32LE(0x07064b50, 0)
    loc.writeBigUInt64LE(BigInt(body.length + innocentCd.length), 8)
    loc.writeUInt32LE(1, 16)
    const crafted = Buffer.concat([body, innocentCd, rec, loc, bomb.subarray(-22)])
    const path = file('decoy64.docx', crafted)
    expect(readZipDirectory(path)!.entries).toEqual([
      { name: 'word/document.xml', compressed: 1024, uncompressed: 200 * MB },
    ])
    const r = (await run(['docs', 'read', path, '--json'])).json()
    expect(r).toMatchObject({ error: 'resource_limit' })
  })

  it('treats 0xFFFF without a zip64 locator as a plain 16-bit count', () => {
    const plain = rawZip([{ name: 'a.xml' }])
    plain.writeUInt16LE(0xffff, plain.length - 22 + 10)
    expect(readZipDirectory(file('ffff.docx', plain), 10000)).toEqual({
      count: 0xffff,
      entries: [],
    })
    expect(readZipDirectory(file('ffff2.docx', plain))!.entries).toHaveLength(1)
  })

  it('refuses oversized parts and totals', async () => {
    const part = file(
      'big.xlsx',
      rawZip([{ name: 'xl/worksheets/sheet1.xml', compressed: 100 * MB, uncompressed: 300 * MB }]),
    )
    expect((await run(['info', part, '--json'])).json()).toMatchObject({
      error: 'resource_limit',
      detail: { entry: 'xl/worksheets/sheet1.xml', limit: 256 * MB },
    })
    const total = file(
      'total.pptx',
      rawZip(
        Array.from({ length: 4 }, (_, i) => ({
          name: `ppt/media/${i}.bin`,
          compressed: 400 * MB,
          uncompressed: 400 * MB,
        })),
      ),
    )
    expect((await run(['slides', 'read', total, '--json'])).json()).toMatchObject({
      error: 'resource_limit',
      detail: { bytes: 1600 * MB },
    })
  })

  it('refuses decompression bombs by declared ratio', async () => {
    const path = file(
      'bomb.docx',
      rawZip([{ name: 'word/document.xml', compressed: 100 * 1024, uncompressed: 100 * MB }]),
    )
    const r = (await run(['docs', 'read', path, '--json'])).json()
    expect(r).toMatchObject({ error: 'resource_limit', detail: { limit: MAX_COMPRESSION_RATIO } })
    expect(r.suggestion).toContain('bomb')
  })

  it('lets a real document through untouched', async () => {
    const path = file('ok.docx', readFileSync(DOCX))
    const r = (await run(['info', path, '--json'])).json()
    expect(r.status).toBe('ok')
    const { entries } = readZipDirectory(path)!
    expect(entries.some((e) => e.name === 'word/document.xml')).toBe(true)
  })
})
