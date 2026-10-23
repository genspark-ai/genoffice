import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

/**
 * `aiOffice` belongs to the shell.
 *
 * The editor apps expose their own bridge for model file naming, and they must
 * not publish it under the shell's name. A second, partial `aiOffice` in an
 * editor window shadows the real one in that window: `window.aiOffice.setTheme`
 * disappears, and anything that identifies the shell window by that global —
 * the e2e suite included — finds the editor instead.
 *
 * These assertions read the preload sources rather than loading them, because
 * the collision lives in the names the preloads register, and a bridge that
 * "works" on its own says nothing about the name it works under.
 */

const APPS = join(fileURLToPath(new URL('.', import.meta.url)), '../../..', 'apps')

interface Preload {
  app: string
  src: string
}

function preloads(): Preload[] {
  return readdirSync(APPS)
    .map((app): Preload | null => {
      const file = join(APPS, app, 'src/preload/index.ts')
      if (!existsSync(file)) return null
      return { app, src: readFileSync(file, 'utf8') }
    })
    .filter((entry): entry is Preload => entry !== null)
}

/** whether this preload registers a global under `name` */
const exposes = (src: string, name: string): boolean =>
  new RegExp(`exposeInMainWorld\\(\\s*'${name}'`).test(src)

describe('the aiOffice namespace', () => {
  it('is exposed by the shell and by nobody else', () => {
    const owners = preloads()
      .filter(({ src }) => exposes(src, 'aiOffice'))
      .map(({ app }) => app)
    expect(owners).toEqual(['shell'])
  })

  it('is not reused by an editor that has a naming bridge of its own', () => {
    for (const { app, src } of preloads()) {
      if (app === 'shell') continue
      // an editor with the naming channels has to publish them under their own
      // name; one without them has nothing to publish
      if (!src.includes('suggestFileName')) continue
      expect(exposes(src, 'aiNamingApi')).toBe(true)
    }
  })
})
