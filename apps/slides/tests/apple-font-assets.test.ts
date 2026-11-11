import { mkdirSync, mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { afterAll, describe, expect, it, vi } from 'vitest'

// Vitest turns ?asset imports into /@fs URLs, while font metrics reads the files from disk.
vi.mock('@genoffice/ui/fonts/Carlito-Regular.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-Regular.ttf'),
}))
vi.mock('@genoffice/ui/fonts/Carlito-Bold.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-Bold.ttf'),
}))
vi.mock('@genoffice/ui/fonts/Carlito-Italic.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-Italic.ttf'),
}))
vi.mock('@genoffice/ui/fonts/Carlito-BoldItalic.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-BoldItalic.ttf'),
}))
// shaped-metrics pulls in the harfbuzz wasm (?asset only resolves in electron builds)
vi.mock('../src/main/shaped-metrics', () => ({
  initShapedMetrics: () => {},
  shapedMeasure: () => null,
  shapedFamily: () => null,
  gtMeasure: () => null,
  complexScriptOf: () => null,
}))

import { appleFontAssetRoots } from '../src/main/fonts'

describe('appleFontAssetRoots (macOS downloadable fonts)', () => {
  const root = mkdtempSync(join(tmpdir(), 'assetsv2-'))
  const pre = join(root, 'PreinstalledAssetsV2/InstallWithOs')
  for (const d of [
    join(root, 'com_apple_MobileAsset_Font6'),
    // macOS 26 leaves the Font7 catalog empty and downloads into Font8: a hardcoded Font7
    // root resolved every downloadable face to its glyph-less stub subset
    join(root, 'com_apple_MobileAsset_Font7'),
    join(root, 'com_apple_MobileAsset_Font8'),
    join(root, 'com_apple_MobileAsset_MacSoftwareUpdate'),
    join(pre, 'com_apple_MobileAsset_Font6'),
    join(pre, 'com_apple_MobileAsset_Font8'),
  ])
    mkdirSync(d, { recursive: true })
  afterAll(() => rmSync(root, { recursive: true, force: true }))

  it('discovers every font catalog, newest version first, downloaded before preinstalled', () => {
    expect(appleFontAssetRoots(root)).toEqual([
      join(root, 'com_apple_MobileAsset_Font8'),
      join(pre, 'com_apple_MobileAsset_Font8'),
      join(root, 'com_apple_MobileAsset_Font7'),
      join(root, 'com_apple_MobileAsset_Font6'),
      join(pre, 'com_apple_MobileAsset_Font6'),
    ])
  })

  it('returns nothing when AssetsV2 is absent (non-macOS)', () => {
    expect(appleFontAssetRoots(join(root, 'missing'))).toEqual([])
  })
})
