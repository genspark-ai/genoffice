/**
 * The Formulas tab's category buttons opened the same unfiltered
 * catalog. They now pass a catalog category, so the categories they advertise
 * have to exist in the catalog (and Financial has to carry real functions).
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const read = (rel: string) => readFileSync(join(here, '..', rel), 'utf8')

const dialogSrc = read('src/renderer/InsertFunctionDialog.tsx')
const shellSrc = read('src/renderer/ExcelShell.tsx')

const catalogCategories = new Set([...dialogSrc.matchAll(/category: '([^']+)'/g)].map((m) => m[1]!))
/** Categories the ribbon buttons ask the dialog to open on */
const requestedCategories = [
  ...shellSrc.matchAll(/functionCategory\(t\('[^']+'\), '[^']*', '([^']+)'\)/g),
].map((m) => m[1]!)

describe('Formulas category buttons', () => {
  it('every ribbon category button asks for a real catalog category (or All)', () => {
    expect(requestedCategories.length).toBe(8)
    for (const cat of requestedCategories) {
      expect(cat === 'All' || catalogCategories.has(cat)).toBe(true)
    }
  })

  it('the buttons no longer all open the same unfiltered list', () => {
    const filtered = requestedCategories.filter((c) => c !== 'All')
    expect(new Set(filtered).size).toBe(filtered.length)
    expect(filtered.length).toBeGreaterThanOrEqual(6)
  })

  it('Financial resolves to actual financial functions', () => {
    expect(catalogCategories.has('Financial')).toBe(true)
    for (const fn of ['PMT', 'NPV', 'IRR']) {
      // whitespace-tolerant: the catalog is prettier-formatted one field per line
      expect(dialogSrc).toMatch(
        new RegExp(`name:\\s*'${fn}',\\s*(?:\\n\\s*)?category:\\s*'Financial'`),
      )
    }
  })

  it('every catalog category has a display label', () => {
    for (const cat of catalogCategories) {
      // keys are quoted only when they need to be ('Date & Time' vs Financial)
      expect(dialogSrc).toMatch(new RegExp(`'?${cat.replace('&', '&')}'?:\\s*'dlgFnCat`))
    }
  })
})
