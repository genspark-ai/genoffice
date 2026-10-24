import { readFileSync, readdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const read = (rel: string) => readFileSync(join(here, '..', rel), 'utf8')

const ribbonSrc = read('src/renderer/components/Ribbon.tsx')
const sharedSrc = read('src/renderer/components/ribbon-shared.tsx')
const appSrc = read('src/renderer/App.tsx')

/**
 * The ribbon's naming entry, wired end to end.
 *
 * A rendering test would be vacuous here — jsdom applies no stylesheet and the
 * ribbon is a thousand-prop component — so this pins the wiring the way the
 * other slides ribbon test does: the File menu carries a button, the button
 * reaches the naming action, and every language can label it.
 */
describe('the slides ribbon can ask the model for a name', () => {
  it('the File menu carries a button that asks for a name', () => {
    expect(ribbonSrc).toMatch(
      /<button\b[^>]*onClick=\{\(\) => \{[\s\S]*?onNameWithAi\(\)[\s\S]*?\}\}[\s\S]*?>[\s\S]*?\{t\('aiNameFile'\)\}[\s\S]*?<\/button>/,
    )
  })

  it('the ribbon takes the handler as a prop', () => {
    expect(sharedSrc).toMatch(/onNameWithAi: \(\) => void/)
  })

  it('the app wires the button to the naming action', () => {
    expect(appSrc).toMatch(/const nameWithAi = useCallback\(\(\) => fileActions\.nameWithAi\(/)
    expect(appSrc).toMatch(/onNameWithAi=\{\(\) => void nameWithAi\(\)\}/)
  })

  it('every language can label it', () => {
    const dir = join(here, '..', 'src', 'renderer', 'i18n', 'ribbon')
    const files = readdirSync(dir).filter((f) => f.endsWith('.ts'))
    expect(files.length).toBeGreaterThan(0)
    for (const file of files) {
      expect(read(`src/renderer/i18n/ribbon/${file}`), file).toMatch(/aiNameFile:/)
    }
  })
})
