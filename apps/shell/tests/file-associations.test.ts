import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { resolve } from 'node:path'
import { runInNewContext } from 'node:vm'
import { describe, expect, it } from 'vitest'
import { LAUNCHABLE_EXTS } from '../src/main/launch-paths'

const require = createRequire(import.meta.url)
// Inspect packaging metadata without downloading Electron or compiling OCR
// helpers (same approach as shell-new.test.ts).
const configModule = { exports: {} as ReturnType<typeof require> }
runInNewContext(readFileSync(resolve(import.meta.dirname, '../electron-builder.cjs'), 'utf8'), {
  module: configModule,
  __dirname: resolve(import.meta.dirname, '..'),
  process: { platform: 'linux', arch: 'x64', env: {} },
  require: (id: string) =>
    id === 'node:fs' ? { ...require(id), existsSync: () => true } : require(id),
})
const config = configModule.exports

describe('packaged file associations', () => {
  // The OS can only offer GenOffice in "Open With" (mac CFBundleDocumentTypes,
  // the Linux desktop MimeType list, the NSIS ProgIds) for types the bundle
  // declares. A routing entry without a matching declaration means a file the
  // app opens fine can never reach it from Finder/Explorer — the .txt/.json
  // gap this closes.
  it('advertises every extension the app launches', () => {
    for (const ext of LAUNCHABLE_EXTS) {
      const association = config.fileAssociations.find(
        (entry: { ext: string }) => entry.ext === ext,
      )
      expect(association, `no file association declared for .${ext}`).toBeTruthy()
    }
  })

  it('gives every association the mimeType the Linux desktop entry needs', () => {
    for (const association of config.fileAssociations) {
      expect(
        association.mimeType,
        `.${association.ext} declares no mimeType — the Linux target drops it`,
      ).toBeTruthy()
    }
  })
})
