import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { FONT_CATALOG, type CatalogFamily } from '../src/font-catalog'
import {
  downloadCatalogEntry,
  downloadFontFamily,
  familyDownloadBytes,
  installLocalFontFiles,
  listCatalog,
  normalizeCdnBaseUrl,
  type FontStoreEnv,
} from '../src/font-store'

/**
 * The store, driven directly — no electron mock, because the environment
 * arrives as an argument. Downloads run against a synthetic family, never a
 * catalog entry: re-pinning a live row's hash to make fake bytes verify edits
 * generated data for every later test, and a hash-shape check cannot see it.
 */

const CDN = 'https://fonts.example.test/v1'
let dir = ''
let env: FontStoreEnv
let served: Map<string, Uint8Array>

function fakeEnv(overrides: Partial<FontStoreEnv> = {}): FontStoreEnv {
  return {
    dir,
    cdnBaseUrl: CDN,
    fetchBytes: async (url) => {
      const body = served.get(url)
      return body
        ? { ok: true, status: 200, bytes: body }
        : { ok: false, status: 404, bytes: new Uint8Array() }
    },
    isFamilyAvailable: () => false,
    // deliberately unlike the filename, so a store that guessed from the
    // filename instead of asking would be caught
    fontFileFamilies: (path) => (path.includes('brand') ? ['Brand Sans'] : ['Some Other Family']),
    ...overrides,
  }
}

/** A family of our own, with its hashes left for `serve` to pin. */
function syntheticFamily(overrides: Partial<CatalogFamily> = {}): CatalogFamily {
  return {
    family: 'Test Sans',
    script: 'latin',
    license: 'OFL-1.1',
    files: [
      { style: 'regular', file: 'TestSans-regular.ttf', sha256: '', bytes: 0 },
      { style: 'bold', file: 'TestSans-bold.ttf', sha256: '', bytes: 0 },
    ],
    ...overrides,
  }
}

/** Serve fake bytes for every file of a family, pinned to match. */
function serve(family: CatalogFamily): void {
  for (const file of family.files) {
    const bytes = Buffer.from(`sfnt-bytes-${file.style}`)
    served.set(new URL(encodeURIComponent(file.file), `${CDN}/`).toString(), bytes)
    file.sha256 = createHash('sha256').update(bytes).digest('hex')
    file.bytes = bytes.length
  }
}

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'font-store-'))
  served = new Map()
  env = fakeEnv()
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

describe('the catalog', () => {
  it('ships regular+bold files with pinned hashes', () => {
    expect(FONT_CATALOG.length).toBeGreaterThanOrEqual(15)
    for (const family of FONT_CATALOG) {
      const styles = new Set(family.files.map((f) => f.style))
      expect(styles.has('regular'), family.family).toBe(true)
      expect(styles.has('bold'), family.family).toBe(true)
      for (const file of family.files) expect(file.sha256).toMatch(/^[0-9a-f]{64}$/)
    }
  })

  it('names a licence inside the allowlist on every family', () => {
    // The bytes come off the mirror at the reader's request, so the release
    // that ships this catalog is the one that has to account for their terms.
    const allowed = new Set(['OFL-1.1', 'Apache-2.0', 'MIT'])
    for (const family of FONT_CATALOG) {
      expect(allowed.has(family.license), `${family.family}: ${family.license}`).toBe(true)
    }
  })

  it('covers the CJK families, not only the Latin rows', () => {
    const cjk = FONT_CATALOG.filter((f) => f.script !== 'latin')
    expect(cjk.length).toBeGreaterThanOrEqual(9)
  })
})

describe('what listing tells the caller', () => {
  it('reports the download size, so a pick never hides a 28 MiB fetch', () => {
    for (const entry of listCatalog(env)) {
      expect(entry.bytes).toBe(
        familyDownloadBytes(FONT_CATALOG.find((f) => f.family === entry.family)!),
      )
    }
    // the expensive one the maintainer named, whether or not it is published yet
    const serif = FONT_CATALOG.find((f) => f.family === 'Noto Serif SC')!
    expect(familyDownloadBytes(serif)).toBeGreaterThan(20 * 1024 * 1024)
  })

  it('carries the licence through to the row', () => {
    expect(listCatalog(env).every((e) => e.license === 'OFL-1.1')).toBe(true)
  })

  it('reports installed state from the host', () => {
    const rows = listCatalog(fakeEnv({ isFamilyAvailable: (f) => f === 'Roboto' }))
    expect(rows.find((e) => e.family === 'Roboto')?.installed).toBe(true)
    expect(rows.find((e) => e.family === 'Open Sans')?.installed).toBe(false)
  })

  it('offers nothing when the build ships no mirror', () => {
    // An affordance that cannot work is worse than no affordance.
    expect(listCatalog(fakeEnv({ cdnBaseUrl: null }))).toEqual([])
  })

  it('offers only families whose files are published', () => {
    const families = listCatalog(env).map((e) => e.family)
    for (const family of FONT_CATALOG) {
      if (family.published === false) expect(families).not.toContain(family.family)
    }
  })
})

describe('downloadCatalogEntry', () => {
  it('verifies the checksum and writes files into the store', async () => {
    const family = syntheticFamily()
    serve(family)
    await downloadCatalogEntry(env, family)
    for (const file of family.files) {
      const path = join(dir, file.file)
      expect(existsSync(path), file.file).toBe(true)
      expect(readFileSync(path).toString()).toContain('sfnt-bytes')
    }
    // familyDownloaded asks about a *catalog* family, so it has nothing to say
    // about a synthetic one — what matters here is that every cut landed
    expect(family.files.every((file) => existsSync(join(dir, file.file)))).toBe(true)
  })

  it('leaves an already-fetched file alone', async () => {
    const family = syntheticFamily()
    serve(family)
    await downloadCatalogEntry(env, family)
    const refetch = fakeEnv({
      fetchBytes: async () => ({ ok: false, status: 500, bytes: new Uint8Array() }),
    })
    // a second run must not depend on the mirror being reachable at all
    await expect(downloadCatalogEntry(refetch, family)).resolves.toBeUndefined()
  })

  it('rejects a non-2xx response even when the bytes hash correctly', async () => {
    // The checksum cannot catch this one: a captive portal can serve the right
    // bytes under an error status, and installing them would register a family
    // the mirror never offered.
    const family = syntheticFamily()
    serve(family)
    const statusOnly = fakeEnv({
      fetchBytes: async (url) => ({ ok: false, status: 403, bytes: served.get(url)! }),
    })
    await expect(downloadCatalogEntry(statusOnly, family)).rejects.toThrow(/HTTP 403/)
    expect(existsSync(join(dir, family.files[0]!.file))).toBe(false)
  })

  it('rejects a checksum mismatch and writes nothing', async () => {
    const family = syntheticFamily()
    serve(family)
    // the mirror starts serving different bytes than the catalog pins
    for (const file of family.files) {
      served.set(
        new URL(encodeURIComponent(file.file), `${CDN}/`).toString(),
        Buffer.from('tampered'),
      )
    }
    await expect(downloadCatalogEntry(env, family)).rejects.toThrow(/checksum/)
    expect(existsSync(join(dir, family.files[0]!.file))).toBe(false)
  })

  it('rejects downloads when no CDN URL is configured', async () => {
    const fetchBytes = vi.fn()
    await expect(
      downloadCatalogEntry(fakeEnv({ cdnBaseUrl: null, fetchBytes }), syntheticFamily()),
    ).rejects.toThrow(/unavailable/)
    expect(fetchBytes).not.toHaveBeenCalled()
  })

  it('a second caller joins the first rather than fetching again', async () => {
    const family = syntheticFamily()
    serve(family)
    let release: (() => void) | null = null
    const gate = new Promise<void>((resolve) => {
      release = resolve
    })
    let fetches = 0
    const slow = fakeEnv({
      fetchBytes: async (url) => {
        fetches++
        await gate
        const body = served.get(url)!
        return { ok: true, status: 200, bytes: body }
      },
    })
    const first = downloadCatalogEntry(slow, family)
    const second = downloadCatalogEntry(slow, family)
    // the join is the whole claim: with the in-flight map deleted, `second`
    // would be a fresh fetch that has also not resolved yet, and this race
    // would resolve first — so assert on the settled flag, not on timing alone
    let secondSettled = false
    void second.then(() => {
      secondSettled = true
    })
    await Promise.resolve()
    await Promise.resolve()
    expect(secondSettled).toBe(false)
    release!()
    await Promise.all([first, second])
    // one fetch per file, not two
    expect(fetches).toBe(family.files.length)
    for (const file of family.files) expect(existsSync(join(dir, file.file)), file.file).toBe(true)
  })
})

describe('downloadFontFamily', () => {
  it('rejects an unknown family', async () => {
    await expect(downloadFontFamily(env, 'Meiryo UI')).rejects.toThrow(/not in catalog/)
  })

  it('refuses a family whose files are not published yet', async () => {
    const family = FONT_CATALOG.find((f) => f.published === false)
    expect(family).toBeDefined()
    await expect(downloadFontFamily(env, family!.family)).rejects.toThrow(/not in catalog/)
  })

  it('refuses to download without a mirror, before touching the network', async () => {
    const fetchBytes = vi.fn()
    await expect(
      downloadFontFamily(fakeEnv({ cdnBaseUrl: null, fetchBytes }), FONT_CATALOG[0]!.family),
    ).rejects.toThrow(/unavailable/)
    expect(fetchBytes).not.toHaveBeenCalled()
  })
})

describe('the shipped catalog', () => {
  it('is not edited by these tests', async () => {
    // the guard the pinned-hash test above cannot be: a download test that
    // re-pinned a live entry would still leave every hash 64 hex characters
    const before = JSON.stringify(FONT_CATALOG)
    const family = syntheticFamily()
    serve(family)
    await downloadCatalogEntry(env, family)
    await expect(downloadFontFamily(env, 'Meiryo UI')).rejects.toThrow()
    expect(JSON.stringify(FONT_CATALOG)).toBe(before)
  })
})

describe('installLocalFontFiles', () => {
  it('accepts sfnt files, copies them in, and skips non-fonts', () => {
    // Picked files live outside the store, which is the whole point of copying
    // them in — a source already at the destination is skipped by design.
    const picked = mkdtempSync(join(tmpdir(), 'font-pick-'))
    try {
      const good = join(picked, 'brand_v2.ttf')
      writeFileSync(good, Buffer.concat([Buffer.from([0, 1, 0, 0]), Buffer.from('x'.repeat(64))]))
      const junk = join(picked, 'notes.txt')
      writeFileSync(junk, 'not a font at all, definitely not sfnt bytes here')
      const families = installLocalFontFiles(env, [good, junk, join(picked, 'missing.ttf')])
      // the family comes from the file's own name table, not from its filename
      expect(families).toEqual(['Brand Sans'])
      expect(existsSync(join(dir, 'Brand Sans.ttf'))).toBe(true)
      // the text file is not a font, so its family is never reported and no
      // store entry appears under any name derived from it
      expect(families).not.toContain('Some Other Family')
    } finally {
      rmSync(picked, { recursive: true, force: true })
    }
  })

  it('installs nothing for a host that offers no local install', () => {
    const picked = mkdtempSync(join(tmpdir(), 'font-pick-'))
    try {
      const good = join(picked, 'brand_v2.ttf')
      writeFileSync(good, Buffer.concat([Buffer.from([0, 1, 0, 0]), Buffer.from('x'.repeat(64))]))
      // a host with no local-install affordance simply omits the reader
      const { fontFileFamilies: _unused, ...host } = fakeEnv()
      expect(installLocalFontFiles(host, [good])).toEqual([])
    } finally {
      rmSync(picked, { recursive: true, force: true })
    }
  })
})

describe('normalizeCdnBaseUrl', () => {
  it('accepts an https URL and drops a trailing slash', () => {
    expect(normalizeCdnBaseUrl('https://cdn.example.test/fonts/')).toBe(
      'https://cdn.example.test/fonts',
    )
  })

  it('refuses anything that is not a plain https origin', () => {
    for (const bad of [
      'http://cdn.example.test',
      'https://user:pw@cdn.example.test',
      'https://cdn.example.test/fonts?x=1',
      'https://cdn.example.test/fonts#frag',
      'not a url',
      '',
      null,
      42,
    ]) {
      expect(normalizeCdnBaseUrl(bad), String(bad)).toBeNull()
    }
  })
})
