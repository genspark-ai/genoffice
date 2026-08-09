import { existsSync, readdirSync } from 'node:fs'
import { join } from 'node:path'

/** PingFang lives under /System/Library/Fonts on older macOS, AssetsV2 on newer */
export const hasPingFang = (() => {
  if (existsSync('/System/Library/Fonts/PingFang.ttc')) return true
  for (const base of [
    '/System/Library/AssetsV2/com_apple_MobileAsset_Font7',
    '/System/Library/AssetsV2/com_apple_MobileAsset_Font6',
  ]) {
    try {
      if (readdirSync(base).some((d) => existsSync(join(base, d, 'AssetData/PingFang.ttc'))))
        return true
    } catch {
      /* base absent on this platform */
    }
  }
  return false
})()
