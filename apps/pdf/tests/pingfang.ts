import { existsSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { appleFontAssetRoots } from '@genoffice/font-metrics'

/** PingFang lives under /System/Library/Fonts on older macOS, AssetsV2 on newer */
export const hasPingFang = (() => {
  if (existsSync('/System/Library/Fonts/PingFang.ttc')) return true
  for (const base of appleFontAssetRoots()) {
    try {
      if (readdirSync(base).some((d) => existsSync(join(base, d, 'AssetData/PingFang.ttc'))))
        return true
    } catch {
      /* catalog unreadable */
    }
  }
  return false
})()
