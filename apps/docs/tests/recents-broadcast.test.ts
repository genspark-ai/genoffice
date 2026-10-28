import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { LANGS } from '@genoffice/i18n'
import { loadAppStrings } from '../src/renderer/i18n/strings-app'

/**
 * docs-main is the full Electron bootstrap, which no unit test can import, so
 * its recents wiring is asserted against the source (same pattern as
 * open-file.test.ts and docs-settings-menu.test.ts).
 */
const source = readFileSync(join(import.meta.dirname, '../src/main/docs-main.ts'), 'utf8')

describe('recents-changed broadcast', () => {
  const push = source.slice(
    source.indexOf('function pushRecent'),
    source.indexOf('function broadcastRecentsChanged'),
  )
  const broadcast = source.slice(
    source.indexOf('function broadcastRecentsChanged'),
    source.indexOf('export function readRecentFiles'),
  )

  it('fires only when the list head actually changed', () => {
    // autosave re-pushes the same head every 30s; the early return keeps the
    // broadcast (and the Home re-pull it triggers) on real changes
    expect(push).toMatch(/if \(recent\[0\] === filePath\) return/)
    expect(push.match(/broadcastRecentsChanged\(\)/g)).toHaveLength(1)
  })

  it('reaches every webContents on the channel the shell Home listens to', () => {
    expect(broadcast).toContain('webContents.getAllWebContents()')
    // mirrors HOME_CHANNELS.recentsChanged in the shell; the name is pinned on
    // both sides (the shell's file-events test is the other half)
    expect(broadcast).toContain("wc.send('home:recents-changed')")
  })
})

describe('save results carry the absolute path', () => {
  it.each(['docs:save', 'docs:save-as', 'docs:save-new', 'docs:save-to'])(
    '%s names the file it wrote',
    (channel) => {
      const at = source.indexOf(`'${channel}'`)
      expect(at).toBeGreaterThan(-1)
      expect(source.slice(at, at + 4000)).toContain('path:')
    },
  )
})

describe('appSavedWithPath', () => {
  it('exists with the {path} placeholder in every locale shard', async () => {
    for (const lang of LANGS) {
      const dict = await loadAppStrings(lang)
      expect(dict.appSavedWithPath, `${lang} is missing appSavedWithPath`).toMatch(/\{path\}/)
    }
  })

  it('the manual-save toast names the saved path', () => {
    const actions = readFileSync(
      join(import.meta.dirname, '../src/renderer/file-actions.ts'),
      'utf8',
    )
    expect(actions).toContain("t('appSavedWithPath', { path: savedPath })")
  })
})
