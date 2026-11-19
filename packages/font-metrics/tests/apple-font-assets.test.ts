import { mkdirSync, mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, describe, expect, it } from 'vitest'
import { appleFontAssetRoots } from '../src/sfnt'

describe('appleFontAssetRoots', () => {
  const root = mkdtempSync(join(tmpdir(), 'assetsv2-'))
  const pre = join(root, 'PreinstalledAssetsV2/InstallWithOs')
  for (const d of [
    join(root, 'com_apple_MobileAsset_Font6'),
    // macOS 26 leaves the Font7 catalog empty and downloads into Font8
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
