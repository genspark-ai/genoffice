import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  cacheControlFor,
  existingFile,
  isBlockedStudioAsset,
  mimeFor,
  safePathUnder,
} from '../src/shared/static-serve'

describe('static-serve', () => {
  it('maps known extensions', () => {
    expect(mimeFor('/x/app.js')).toBe('text/javascript; charset=utf-8')
    expect(mimeFor('/x/font.woff2')).toBe('font/woff2')
    expect(mimeFor('/x/unknown.bin')).toBe('application/octet-stream')
    expect(cacheControlFor('/x/index.html')).toBe('no-cache')
    expect(cacheControlFor('/x/app.js')).toBe('public, max-age=31536000, immutable')
  })

  it('blocks stock studio PWA files', () => {
    expect(isBlockedStudioAsset('/rhwp/sw.js')).toBe(true)
    expect(isBlockedStudioAsset('/registerSW.js?v=1')).toBe(true)
    expect(isBlockedStudioAsset('/manifest.webmanifest')).toBe(true)
    expect(isBlockedStudioAsset('/assets/app.js')).toBe(false)
  })

  it('resolves files under the root and rejects traversal', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'hwp-static-'))
    await writeFile(join(dir, 'index.html'), 'ok')
    expect(existingFile(safePathUnder(dir, '/'))).toBe(join(dir, 'index.html'))
    expect(safePathUnder(dir, '/../../../etc/passwd')).toBeNull()
    expect(safePathUnder(dir, '/sw.js')).toBeNull()
    expect(existingFile(safePathUnder(dir, '/missing.html'))).toBeNull()
  })
})
