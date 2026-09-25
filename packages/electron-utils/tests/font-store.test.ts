import { existsSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { FONT_CATALOG } from '../src/font-catalog'
import {
  downloadFontFamily,
  installLocalFontFiles,
  readStoreFontFace,
  resolveFontCdnBaseUrl,
  storeFontFaces,
} from '../src/font-store'

/** UTF-16BE bytes: the name table's Windows-platform string encoding. */
function utf16be(s: string): Buffer {
  const out = Buffer.alloc(s.length * 2)
  for (let i = 0; i < s.length; i++) out.writeUInt16BE(s.charCodeAt(i), i * 2)
  return out
}

/** Minimal real sfnt with a name table (nameID 1 family + 2 subfamily) and an
 *  OS/2 table carrying usWeightClass, enough for the shared store parser. */
function buildTtf(family: string, subfamily = 'Regular', weight = 400): Buffer {
  const ids = [1, 2]
  const strings = [utf16be(family), utf16be(subfamily)]
  const strBase = 6 + 12 * strings.length
  const nameLen = strBase + strings.reduce((n, s) => n + s.length, 0)
  const name = Buffer.alloc(nameLen)
  name.writeUInt16BE(0, 0)
  name.writeUInt16BE(strings.length, 2)
  name.writeUInt16BE(strBase, 4)
  let strOff = 0
  for (let i = 0; i < strings.length; i++) {
    const r = 6 + 12 * i
    name.writeUInt16BE(3, r) // platform: Windows
    name.writeUInt16BE(1, r + 2) // encoding: UTF-16
    name.writeUInt16BE(0x409, r + 4) // language: en-US
    name.writeUInt16BE(ids[i]!, r + 6)
    name.writeUInt16BE(strings[i]!.length, r + 8)
    name.writeUInt16BE(strOff, r + 10)
    strings[i]!.copy(name, strBase + strOff)
    strOff += strings[i]!.length
  }
  const os2 = Buffer.alloc(8)
  os2.writeUInt16BE(weight, 4) // usWeightClass sits at +4
  const numTables = 2
  const header = Buffer.alloc(12 + 16 * numTables)
  header.writeUInt32BE(0x00010000, 0) // sfnt version
  header.writeUInt16BE(numTables, 4)
  header.write('name', 12, 'ascii')
  header.writeUInt32BE(12 + 16 * numTables, 20) // name table offset: right after the header
  header.writeUInt32BE(nameLen, 24)
  header.write('OS/2', 28, 'ascii')
  header.writeUInt32BE(12 + 16 * numTables + nameLen, 36) // OS/2 offset
  header.writeUInt32BE(os2.length, 40)
  return Buffer.concat([header, name, os2])
}

let storeDir = mkdtempSync(join(tmpdir(), 'font-store-core-'))
afterEach(() => {
  storeDir = mkdtempSync(join(tmpdir(), 'font-store-core-'))
})

describe('font store core', () => {
  it('catalog families ship regular+bold with pinned hashes', () => {
    for (const fam of FONT_CATALOG) {
      const styles = fam.files.map((f) => f.style)
      expect(styles).toContain('regular')
      expect(styles).toContain('bold')
      for (const f of fam.files) {
        expect(f.file).toMatch(/\.ttf$/)
        expect(f.sha256).toMatch(/^[0-9a-f]{64}$/)
        expect(f.bytes).toBeGreaterThan(10_000)
      }
    }
  })

  it('resolves the CDN URL from packaged metadata or env, never from insecure input', () => {
    expect(
      resolveFontCdnBaseUrl({
        isPackaged: false,
        appPath: '/x',
        envUrl: ' https://fonts.example.test/v1/ ',
      }),
    ).toBe('https://fonts.example.test/v1')
    expect(resolveFontCdnBaseUrl({ isPackaged: true, appPath: '/nonexistent', envUrl: '' })).toBeNull()
    expect(
      resolveFontCdnBaseUrl({ isPackaged: false, appPath: '/x', envUrl: 'http://fonts.example.test' }),
    ).toBeNull()
  })

  it('downloads catalog files with checksum verification into the store', async () => {
    const family = FONT_CATALOG[0]!
    // Re-pin the first family's hashes to the fake payloads for the test
    for (const f of family.files) {
      f.sha256 = createHash('sha256').update(Buffer.from(`sfnt-${f.style}`)).digest('hex')
    }
    const fetchImpl = vi.fn(async (input: string) => {
      const file = decodeURIComponent(input.split('/').pop()!)
      const style = family.files.find((f) => f.file === file)!.style
      return new Response(new Uint8Array(Buffer.from(`sfnt-${style}`)), { status: 200 })
    })
    await downloadFontFamily({ baseUrl: 'https://fonts.example.test/v1', storeDir, fetchImpl }, family.family)
    expect(fetchImpl).toHaveBeenCalled()
    for (const f of family.files) expect(existsSync(join(storeDir, f.file))).toBe(true)
  })

  it('rejects unknown families and checksum mismatches', async () => {
    const fetchImpl = vi.fn(async () => new Response(new Uint8Array(Buffer.from('tampered')), { status: 200 }))
    await expect(
      downloadFontFamily({ baseUrl: 'https://fonts.example.test/v1', storeDir, fetchImpl }, 'Meiryo UI'),
    ).rejects.toThrow(/not in catalog/)
    await expect(
      downloadFontFamily(
        { baseUrl: 'https://fonts.example.test/v1', storeDir, fetchImpl },
        FONT_CATALOG[1]!.family,
      ),
    ).rejects.toThrow(/checksum/)
  })

  it('installs user font files renamed to their family and skips junk', () => {
    const srcDir = mkdtempSync(join(tmpdir(), 'font-store-src-'))
    const src = join(srcDir, 'brand_v2_final.ttf')
    writeFileSync(src, buildTtf('Brand Sans'))
    const bold = join(srcDir, 'brand-bold.ttf')
    writeFileSync(bold, buildTtf('Brand Sans', 'Bold', 700))
    const junk = join(srcDir, 'junk.ttf')
    writeFileSync(junk, Buffer.from('MZ not a font'))
    const families = installLocalFontFiles(storeDir, [src, bold, junk])
    expect(families).toEqual(['Brand Sans'])
    const stored = storeFontFaces(storeDir)
    // readdir order is not contractual
    expect(stored.map((f) => `${f.family}/${f.weight}`).sort()).toEqual(['Brand Sans/400', 'Brand Sans/700'])
  })

  it('serves store faces for renderer FontFace registration and rejects traversal', () => {
    writeFileSync(join(storeDir, 'Test Family.ttf'), buildTtf('Test Family'))
    const faces = storeFontFaces(storeDir)
    expect(faces).toHaveLength(1)
    const data = readStoreFontFace(storeDir, faces[0]!.file, faces[0]!.faceOffset)
    expect(data.byteLength).toBeGreaterThan(0)
    expect(() => readStoreFontFace(storeDir, '../escape.ttf', 0)).toThrow(/not a store font/)
    expect(() => readStoreFontFace(storeDir, 'sub/dir.ttf', 0)).toThrow(/not a store font/)
  })
})
