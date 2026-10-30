/**
 * Pins every @font-face descriptor the renderer registers and every line
 * factor the pagination model answers, so a refactor of where those numbers
 * live (hand-written CSS, generated CSS, manifest) is provably value-neutral.
 * A deliberate metric change updates the fixtures with `vitest -u` and shows
 * up in review as a fixture diff, never as a silent drift.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'

import { describe, expect, it } from 'vitest'

import wordFaces from '../../../packages/font-metrics/word-faces.json'
import { lineHeightFactor, tableLineHeightFactor } from '../src/renderer/line-metrics'

const FONTS_CSS = join(__dirname, '../src/renderer/fonts/fonts.css')
const FIXTURES = join(__dirname, 'fixtures/font-metrics')

/** the stylesheet text with its relative @import lines inlined */
function cssWithImports(file: string): string {
  return readFileSync(file, 'utf8').replace(
    /^@import\s+(?:url\()?['"]([^'"]+)['"]\)?\s*;/gm,
    (_, rel: string) => cssWithImports(join(dirname(file), rel)),
  )
}

type Block = Record<string, string>

/** every @font-face block as a declaration map, grouped per family in source order */
export function fontFaceFamilies(css: string): Record<string, Block[]> {
  const out: Record<string, Block[]> = {}
  for (const m of css.matchAll(/@font-face\s*\{([^}]*)\}/g)) {
    const block: Block = {}
    for (const decl of m[1]!.replace(/\/\*[\s\S]*?\*\//g, '').split(';')) {
      const i = decl.indexOf(':')
      if (i < 0) continue
      block[decl.slice(0, i).trim()] = decl
        .slice(i + 1)
        .replace(/\s+/g, ' ')
        .trim()
    }
    const family = block['font-family']!.replace(/^'|'$/g, '')
    ;(out[family] ??= []).push(
      Object.fromEntries(Object.entries(block).sort(([a], [b]) => a.localeCompare(b))),
    )
  }
  return Object.fromEntries(Object.entries(out).sort(([a], [b]) => a.localeCompare(b)))
}

/** JSON with non-ASCII escaped: the fixtures hold CJK face names as \\u escapes */
const stable = (value: unknown) =>
  JSON.stringify(value, null, 2).replace(
    /[^\x20-\x7e\n]/g,
    (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0'),
  ) + '\n'

const testFile = readFileSync(join(__dirname, 'line-metrics.test.ts'), 'utf8')
const testNames = [...testFile.matchAll(/'([^'\n]{1,40})'|"([^"\n]{1,40})"/g)].map(
  (m) => m[1] ?? m[2]!,
)

describe('font metric baseline', () => {
  it('registers the pinned @font-face descriptor set', async () => {
    const families = fontFaceFamilies(cssWithImports(FONTS_CSS))
    await expect(stable(families)).toMatchFileSnapshot(join(FIXTURES, 'font-faces.json'))
  })

  it('answers the pinned line factor for every Word face and test name', async () => {
    const names = new Set<string>()
    for (const [key, face] of Object.entries(wordFaces.faces)) {
      names.add(key)
      names.add((face as { family: string }).family)
    }
    for (const n of testNames) names.add(n)
    const factors: Record<string, [number, number]> = {}
    for (const name of [...names].sort()) {
      factors[name] = [lineHeightFactor(name), tableLineHeightFactor(name)]
    }
    await expect(stable(factors)).toMatchFileSnapshot(join(FIXTURES, 'line-factors.json'))
  })
})
