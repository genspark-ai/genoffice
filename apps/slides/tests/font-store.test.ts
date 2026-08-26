import { mkdtempSync, readFileSync, writeFileSync, existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest'
import { createBlankPptx, openPptx } from '@genoffice/pptx-engine'

const storeDir = mkdtempSync(join(tmpdir(), 'font-store-'))
const fontCdnBaseUrl = 'https://fonts.example.test/v1'

vi.mock('electron', () => ({
  app: { getPath: () => storeDir, getAppPath: () => storeDir, isPackaged: false },
  net: { fetch: vi.fn() },
}))

const availability = new Map<string, boolean>()
vi.mock('../src/main/fonts', () => ({
  familyAvailable: (f: string) => availability.get(f) ?? false,
  fontFileFamilies: (p: string) => {
    // The magic gate runs before this; tests hand-label families per path
    return p.includes('brand') ? ['Brand Sans'] : ['Test Family']
  },
  setUserFontDir: vi.fn(),
}))

import { FONT_CATALOG } from '../src/main/font-catalog'
import {
  downloadFontFamily,
  extractFontCdnBaseUrl,
  installLocalFontFiles,
  listFontCatalog,
  missingCatalogFonts,
} from '../src/main/font-store'
import { net } from 'electron'
import { createHash } from 'node:crypto'

beforeEach(() => {
  availability.clear()
  vi.stubEnv('GENOFFICE_FONT_CDN_URL', fontCdnBaseUrl)
  vi.mocked(net.fetch).mockReset()
})
afterEach(() => vi.unstubAllEnvs())

describe('font catalog', () => {
  it('every family ships regular+bold files with pinned hashes but no endpoint', () => {
    expect(FONT_CATALOG.length).toBeGreaterThanOrEqual(15)
    for (const fam of FONT_CATALOG) {
      const styles = fam.files.map((f) => f.style)
      expect(styles).toContain('regular')
      expect(styles).toContain('bold')
      for (const f of fam.files) {
        expect(f.file).toMatch(/\.ttf$/)
        expect(f).not.toHaveProperty('url')
        expect(f.sha256).toMatch(/^[0-9a-f]{64}$/)
        expect(f.bytes).toBeGreaterThan(10_000)
      }
    }
  })

  it('reports installed state from the registry', () => {
    availability.set('Poppins', true)
    const list = listFontCatalog()
    expect(list.find((e) => e.family === 'Poppins')?.installed).toBe(true)
    expect(list.find((e) => e.family === 'Montserrat')?.installed).toBe(false)
  })

  it('hides the downloadable catalog when no CDN URL is configured', () => {
    vi.stubEnv('GENOFFICE_FONT_CDN_URL', '')
    expect(listFontCatalog()).toEqual([])
  })

  it('extracts and validates the packaged CDN URL', () => {
    expect(
      extractFontCdnBaseUrl({
        genofficeFontCdn: { baseUrl: ' https://fonts.example.test/v1/ ' },
      }),
    ).toBe(fontCdnBaseUrl)
    expect(
      extractFontCdnBaseUrl({ genofficeFontCdn: { baseUrl: 'http://fonts.example.test/v1' } }),
    ).toBeNull()
    expect(
      extractFontCdnBaseUrl({
        genofficeFontCdn: { baseUrl: 'https://user@fonts.example.test/v1' },
      }),
    ).toBeNull()
    expect(
      extractFontCdnBaseUrl({
        genofficeFontCdn: { baseUrl: 'https://fonts.example.test/v1?token=secret' },
      }),
    ).toBeNull()
  })
})

describe('downloadFontFamily', () => {
  it('verifies the checksum and writes files into the store', async () => {
    const fam = FONT_CATALOG[0]!
    const payloads = new Map(
      fam.files.map((f) => {
        const buf = Buffer.from(`sfnt-bytes-${f.style}`)
        const url = new URL(encodeURIComponent(f.file), `${fontCdnBaseUrl}/`).toString()
        return [url, buf] as const
      }),
    )
    // Re-pin hashes to the fake payloads for the test
    for (const f of fam.files) {
      const url = new URL(encodeURIComponent(f.file), `${fontCdnBaseUrl}/`).toString()
      f.sha256 = createHash('sha256').update(payloads.get(url)!).digest('hex')
    }
    vi.mocked(net.fetch).mockImplementation(async (url: unknown) => {
      const buf = payloads.get(String(url))!
      return new Response(new Uint8Array(buf), { status: 200 })
    })
    await downloadFontFamily(fam.family)
    for (const f of fam.files) {
      const p = join(storeDir, 'fonts', f.file)
      expect(existsSync(p)).toBe(true)
      expect(readFileSync(p).toString()).toContain('sfnt-bytes')
    }
  })

  it('rejects a checksum mismatch', async () => {
    const fam = FONT_CATALOG[1]!
    vi.mocked(net.fetch).mockResolvedValue(
      new Response(new Uint8Array(Buffer.from('tampered')), { status: 200 }),
    )
    await expect(downloadFontFamily(fam.family)).rejects.toThrow(/checksum/)
  })

  it('rejects unknown families', async () => {
    await expect(downloadFontFamily('Meiryo UI')).rejects.toThrow(/not in catalog/)
  })

  it('rejects downloads when no CDN URL is configured', async () => {
    vi.stubEnv('GENOFFICE_FONT_CDN_URL', '')
    await expect(downloadFontFamily(FONT_CATALOG[0]!.family)).rejects.toThrow(/unavailable/)
    expect(net.fetch).not.toHaveBeenCalled()
  })
})

describe('installLocalFontFiles', () => {
  it('accepts sfnt files, renames to the family, and skips non-fonts', () => {
    const src = join(storeDir, 'brand_v2_final.ttf')
    writeFileSync(src, Buffer.concat([Buffer.from([0, 1, 0, 0]), Buffer.from('x'.repeat(64))]))
    const junk = join(storeDir, 'junk.ttf')
    writeFileSync(junk, Buffer.from('MZ not a font'))
    const families = installLocalFontFiles([src, junk])
    expect(families).toEqual(['Brand Sans'])
    expect(existsSync(join(storeDir, 'fonts', 'Brand Sans.ttf'))).toBe(true)
  })
})

describe('missingCatalogFonts', () => {
  it('reports deck-referenced catalog families that are unavailable, including table cells', async () => {
    const { deck } = await openPptx(await createBlankPptx())
    const slide = deck.slides[0]!
    ;(slide.elements as unknown[]).push(
      {
        type: 'text',
        text: {
          paragraphs: [
            { runs: [{ fontFamily: 'Poppins' }, { fontFamily: 'Arial' }] },
            { runs: [{ fontFamily: 'Montserrat' }] },
          ],
        },
      },
      {
        type: 'table',
        rows: [[{ text: { paragraphs: [{ runs: [{ fontFamily: 'Rubik' }] }] } }]],
      },
    )
    availability.set('Montserrat', true)
    const missing = missingCatalogFonts({ deck } as never)
    expect(missing).toEqual(['Poppins', 'Rubik'])

    vi.stubEnv('GENOFFICE_FONT_CDN_URL', '')
    expect(missingCatalogFonts({ deck } as never)).toEqual([])
  })
})

describe('concurrent downloads', () => {
  it('a second call joins the in-flight download instead of resolving early', async () => {
    const fam = FONT_CATALOG[2]!
    let releaseFetch: (() => void) | null = null
    const gate = new Promise<void>((res) => (releaseFetch = res))
    const payload = Buffer.from('slow-bytes')
    for (const f of fam.files) f.sha256 = createHash('sha256').update(payload).digest('hex')
    vi.mocked(net.fetch).mockImplementation(async () => {
      await gate
      return new Response(new Uint8Array(payload), { status: 200 })
    })
    const first = downloadFontFamily(fam.family)
    const second = downloadFontFamily(fam.family)
    let secondDone = false
    void second.then(() => (secondDone = true))
    await new Promise((r) => setTimeout(r, 20))
    expect(secondDone).toBe(false)
    releaseFetch!()
    await Promise.all([first, second])
    expect(secondDone).toBe(true)
  })
})
