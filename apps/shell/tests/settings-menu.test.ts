import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

/**
 * The shell owns four native menu sets (home, pdf, markdown and html tabs).
 * The Settings Center redesign removes the Settings item from every File menu
 * and gives it a top-level menu of its own, inserted right before the Window
 * menu; the macOS app-menu entry stays where the platform convention puts it.
 * The builders run inside Electron startup and cannot be imported here, so the
 * wiring is pinned from the source, the way main-locale-tables.test.ts reads
 * the locale table.
 */
const source = readFileSync(join(__dirname, '../src/main/index.ts'), 'utf8')

describe('the four application-menu sets', () => {
  it('imports settingsMenuTemplate from electron-utils', () => {
    expect(source).toMatch(
      /import \{[^}]*\bsettingsMenuTemplate\b[^}]*\} from '@genoffice\/electron-utils'/,
    )
  })

  it('inserts the top-level settings menu right before the Window menu in every set', () => {
    const pairs = source.match(
      /settingsMenuTemplate\(labels\),\n {4}windowMenuTemplate\(process\.platform/g,
    )
    expect(pairs).toHaveLength(4)
  })

  it('keeps no Settings item inside any File menu', () => {
    // settingsMenuItem survives only as the macOS app-menu extra
    const placements = source.match(/settingsMenuItem\(labels\)/g) ?? []
    const extras = source.match(/extra: \[settingsMenuItem\(labels\)\]/g) ?? []
    expect(placements).toHaveLength(4)
    expect(extras).toHaveLength(4)
  })
})
