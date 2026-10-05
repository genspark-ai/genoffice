import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { afterAll, afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createBlankPptx, openPptx } from '@genoffice/pptx-engine'

/**
 * The slides side of the font store: the adapter, not the store.
 *
 * The download discipline — checksums, the in-flight join, the local install —
 * belongs to @genoffice/electron-utils and is tested there against a synthetic
 * family. What cannot be covered from there is what only this app knows: where
 * the mirror URL comes from, and which catalog families a deck actually asks
 * for.
 */

const storeDir = mkdtempSync(join(tmpdir(), 'slides-font-store-'))
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

import {
  FONT_CATALOG,
  downloadCatalogEntry,
  downloadFontFamily,
  extractFontCdnBaseUrl,
  installLocalFontFiles,
  listFontCatalog,
  missingCatalogFonts,
} from '../src/main/font-store'
import { net } from 'electron'

/**
 * The catalog as it stood when this file loaded.
 *
 * FONT_CATALOG is a module-level array, so a download test that re-pins a row's
 * sha256 to make its fake bytes verify silently replaces the real hashes for
 * every later test in the run — and the corruption is invisible, because the
 * replacements are themselves valid hex. Snapshotting at import and comparing
 * after the last test is the only check that actually catches it: a guard
 * written as a test *inside* the suite passes as long as no earlier test has
 * corrupted anything yet.
 */
const CATALOG_AT_IMPORT = JSON.stringify(FONT_CATALOG)

/**
 * Add one file's fake bytes to the mirror this test pretends to have, and return
 * the hash the catalog would pin. Accumulates, so a family of several cuts is
 * served by one call rather than each call replacing the last.
 */
const payloads = new Map<string, string>()
function serve(entry: { file: string; style: string }): string {
  const body = `sfnt-bytes-${entry.style}`
  payloads.set(entry.file, body)
  vi.mocked(net.fetch).mockImplementation(async (url: unknown) => {
    const wanted = String(url).split('/').pop() ?? ''
    const hit = payloads.get(decodeURIComponent(wanted))
    return hit
      ? new Response(new Uint8Array(Buffer.from(hit)), { status: 200 })
      : new Response(new Uint8Array(0), { status: 404 })
  })
  return createHash('sha256').update(Buffer.from(body)).digest('hex')
}

beforeEach(() => {
  availability.clear()
  payloads.clear()
  vi.stubEnv('GENOFFICE_FONT_CDN_URL', fontCdnBaseUrl)
  vi.mocked(net.fetch).mockReset()
})
afterEach(() => {
  vi.unstubAllEnvs()
  rmSync(join(storeDir, 'fonts'), { recursive: true, force: true })
})
afterAll(() => {
  expect(
    JSON.stringify(FONT_CATALOG),
    'a test re-pinned a live catalog entry; drive a synthetic family via downloadCatalogEntry instead',
  ).toBe(CATALOG_AT_IMPORT)
  rmSync(storeDir, { recursive: true, force: true })
})

describe('the catalog rows this app hands its ribbon', () => {
  it('every family ships regular+bold files with pinned hashes but no endpoint', () => {
    expect(FONT_CATALOG.length).toBeGreaterThanOrEqual(15)
    for (const fam of FONT_CATALOG) {
      const styles = fam.files.map((f) => f.style)
      expect(styles).toContain('regular')
      expect(styles).toContain('bold')
      for (const f of fam.files) {
        expect(f.file).toMatch(/\.ttf$/)
        // no per-file endpoint: the mirror is one base URL, so a catalog that
        // grew per-file URLs would be a second place to get the base wrong
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

  it('prices every row, so the ribbon can say what a pick would fetch', () => {
    for (const entry of listFontCatalog()) expect(entry.bytes).toBeGreaterThan(10_000)
  })

  it('hides the downloadable catalog when no CDN URL is configured', () => {
    vi.stubEnv('GENOFFICE_FONT_CDN_URL', '')
    expect(listFontCatalog()).toEqual([])
  })

  it('offers only families whose files are published', () => {
    // The CJK serif families land before the CDN carries their files: until the
    // mirror publishes them, a picker row would offer a download that 404s.
    const unpublished = FONT_CATALOG.filter((f) => f.published === false)
    expect(unpublished.length).toBeGreaterThan(0)
    const offered = new Set(listFontCatalog().map((e) => e.family))
    for (const fam of unpublished) expect(offered.has(fam.family)).toBe(false)
    // …and the published rows are all still offered
    expect(offered.has('Poppins')).toBe(true)
  })

  it('refuses to download a family that is not published', async () => {
    const family = FONT_CATALOG.find((f) => f.published === false)!.family
    await expect(downloadFontFamily(family)).rejects.toThrow(/not in catalog/)
    expect(net.fetch).not.toHaveBeenCalled()
  })
})

describe('extractFontCdnBaseUrl', () => {
  it('reads and normalises the packaged mirror URL', () => {
    expect(
      extractFontCdnBaseUrl({
        genofficeFontCdn: { baseUrl: ' https://fonts.example.test/v1/ ' },
      }),
    ).toBe(fontCdnBaseUrl)
  })

  it.each([
    ['plain http', 'http://fonts.example.test/v1'],
    ['credentials in the URL', 'https://user@fonts.example.test/v1'],
    [
      'a query string, which is where a signed-URL token would leak',
      'https://fonts.example.test/v1?token=secret',
    ],
    ['no font section at all', { name: 'slides' }],
    ['nothing', null],
  ])('refuses %s', (_label, pkg) => {
    expect(extractFontCdnBaseUrl(pkg)).toBeNull()
  })
})

describe('downloadFontFamily, through this app env', () => {
  it('verifies the checksum and writes files into the store', async () => {
    const entry = {
      family: 'Test Download Family',
      script: 'latin' as const,
      license: 'OFL-1.1' as const,
      files: (['regular', 'bold'] as const).map((style) => ({
        style,
        file: `TestDownload-${style}.ttf`,
        sha256: '',
        bytes: 0,
      })),
    }
    for (const file of entry.files) file.sha256 = serve(file)
    await downloadCatalogEntry(entry)
    for (const file of entry.files) {
      const path = join(storeDir, 'fonts', file.file)
      expect(existsSync(path), file.file).toBe(true)
      expect(readFileSync(path).toString()).toContain('sfnt-bytes')
    }
  })

  it('rejects a checksum mismatch from the mirror', async () => {
    const family = FONT_CATALOG[1]!
    vi.mocked(net.fetch).mockResolvedValue(
      new Response(new Uint8Array(Buffer.from('tampered')), { status: 200 }),
    )
    await expect(downloadFontFamily(family.family)).rejects.toThrow(/checksum/)
  })

  it('rejects unknown families', async () => {
    await expect(downloadFontFamily('Meiryo UI')).rejects.toThrow(/not in catalog/)
  })

  it('rejects downloads when no CDN URL is configured', async () => {
    vi.stubEnv('GENOFFICE_FONT_CDN_URL', '')
    await expect(downloadFontFamily(FONT_CATALOG[0]!.family)).rejects.toThrow(/unavailable/)
    expect(net.fetch).not.toHaveBeenCalled()
  })

  it('a second call joins the in-flight download instead of resolving early', async () => {
    // driven through a family of our own: re-pinning a real row's hashes to
    // make fake bytes verify would edit the catalog for every later test
    const entry = {
      family: 'Test Join Family',
      script: 'latin' as const,
      license: 'OFL-1.1' as const,
      files: [{ style: 'regular' as const, file: 'TestJoin-regular.ttf', sha256: '', bytes: 0 }],
    }
    const payload = Buffer.from('slow-bytes')
    entry.files[0]!.sha256 = createHash('sha256').update(payload).digest('hex')
    let releaseFetch: (() => void) | null = null
    const gate = new Promise<void>((res) => (releaseFetch = res))
    vi.mocked(net.fetch).mockImplementation(async () => {
      await gate
      return new Response(new Uint8Array(payload), { status: 200 })
    })
    const first = downloadCatalogEntry(entry)
    const second = downloadCatalogEntry(entry)
    let secondDone = false
    void second.then(() => (secondDone = true))
    // with the in-flight map deleted, `second` is a second fetch that also has
    // not resolved — but it would be *counted*, which is what this asserts
    await new Promise((r) => setTimeout(r, 20))
    expect(secondDone).toBe(false)
    expect(vi.mocked(net.fetch).mock.calls.length).toBe(1)
    releaseFetch!()
    await Promise.all([first, second])
    expect(secondDone).toBe(true)
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
    // Arial is not in the catalog, so it is nobody's download to offer
    const missing = missingCatalogFonts({ deck } as never)
    expect(missing).toEqual(['Poppins', 'Rubik'])

    vi.stubEnv('GENOFFICE_FONT_CDN_URL', '')
    expect(missingCatalogFonts({ deck } as never)).toEqual([])
  })
})
