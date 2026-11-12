import { afterEach, describe, expect, it } from 'vitest'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { normalizeRecentQuery, pageRecentPaths } from '../src/main/recent-files'

const dirs: string[] = []
afterEach(() => {
  for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true })
})

function files(...names: string[]): string[] {
  const dir = mkdtempSync(join(tmpdir(), 'go-search-'))
  dirs.push(dir)
  return names.map((n) => {
    const p = join(dir, n)
    writeFileSync(p, 'x')
    return p
  })
}

describe('recent search query', () => {
  it('normalizes the text filter', () => {
    expect(normalizeRecentQuery({ query: '  Budget ' }).query).toBe('budget')
    expect(normalizeRecentQuery({ query: '   ' }).query).toBeUndefined()
    expect(normalizeRecentQuery({ query: 5 }).query).toBeUndefined()
  })

  it('matches names case-insensitively and reports the filtered total', async () => {
    const paths = files('Budget 2026.xlsx', 'notes.md', 'budget-old.docx')
    const page = await pageRecentPaths(paths, { query: 'BUDGET' }, new Set())
    expect(page.entries.map((e) => e.name)).toEqual(['Budget 2026.xlsx', 'budget-old.docx'])
    expect(page.total).toBe(2)
    expect(page.totalAll).toBe(3)
  })

  it('combines with the extension filter and paging', async () => {
    const paths = files('a.docx', 'ab.docx', 'ab.xlsx')
    const page = await pageRecentPaths(
      paths,
      { query: 'ab', ext: 'docx', offset: 0, limit: 1 },
      new Set(),
    )
    expect(page.entries.map((e) => e.name)).toEqual(['ab.docx'])
    expect(page.total).toBe(1)
  })

  it('returns everything without a query', async () => {
    const paths = files('a.md', 'b.md')
    expect((await pageRecentPaths(paths, {}, new Set())).total).toBe(2)
  })
})
