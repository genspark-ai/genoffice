import { describe, expect, it } from 'vitest'
import { readFileSync, readdirSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

/**
 * The web preloads are built with `electron` aliased to the browser shim
 * (tools/build-web-preloads.mjs), so every name an app preload imports from
 * 'electron' has to exist as a static export of the shim. A missing one does
 * not fail a unit test — it fails the web build with esbuild's "No matching
 * export", which is how the sheets preload's webFrame import was found.
 *
 * This test reads the preloads and the shim as text: the contract esbuild
 * checks is a static one, and the shim touches `window` at import time, which
 * the node test environment has no business providing.
 */

const shellRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const repoRoot = resolve(shellRoot, '..', '..')
const stubPath = join(shellRoot, 'src', 'web', 'shim', 'electron-stub.ts')

/** the names a preload pulls out of 'electron' */
function electronImports(source: string): string[] {
  const match = /import\s*\{([^}]*)\}\s*from\s*'electron'/.exec(source)
  if (!match) return []
  return match[1]!
    .split(',')
    .map((name) => name.trim())
    .filter((name) => name.length > 0)
}

/** the names the shim exports statically */
function stubExports(source: string): Set<string> {
  const names = new Set<string>()
  for (const match of source.matchAll(
    /^export\s+(?:const|let|var|function|class|async function)\s+(\w+)/gm,
  )) {
    names.add(match[1]!)
  }
  for (const match of source.matchAll(/^export\s*\{([^}]*)\}/gm)) {
    for (const part of match[1]!.split(',')) {
      const name = part
        .trim()
        .split(/\s+as\s+/)
        .pop()
      if (name) names.add(name)
    }
  }
  return names
}

describe('the browser electron shim', () => {
  it('exports every name the app preloads import from electron', () => {
    const exports = stubExports(readFileSync(stubPath, 'utf8'))
    const appsDir = join(repoRoot, 'apps')
    const missing: string[] = []
    for (const app of readdirSync(appsDir)) {
      const preload = join(appsDir, app, 'src', 'preload', 'index.ts')
      let source: string
      try {
        source = readFileSync(preload, 'utf8')
      } catch {
        continue
      }
      for (const name of electronImports(source)) {
        if (!exports.has(name)) missing.push(`${app}: ${name}`)
      }
    }
    expect(missing).toEqual([])
  })

  it('keeps the spell engine a no-op, the way Electron reports an off checker', () => {
    const source = readFileSync(stubPath, 'utf8')
    expect(source).toMatch(/export const webFrame/)
    expect(source).toMatch(/isWordMisspelled\(\): boolean\s*\{\s*return false\s*\}/)
    expect(source).toMatch(/getWordSuggestions\(\): string\[\]\s*\{\s*return \[\]\s*\}/)
  })
})
