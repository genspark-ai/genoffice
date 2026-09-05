import { describe, expect, it } from 'vitest'

import { matchesExtFamily, pageRecentPaths } from '../src/main/recent-files'

const NONE = new Set<string>()
const PATHS = [
  '/d/new.docx',
  '/d/legacy.doc',
  '/d/book.xlsx',
  '/d/macro.xlsm',
  '/d/old.xls',
  '/d/data.csv',
  '/d/notes.md',
  '/d/readme.markdown',
  '/d/deck.pptx',
]

describe('pageRecentPaths ext families', () => {
  it('includes legacy .doc under the docx filter', () => {
    const page = pageRecentPaths(PATHS, { ext: 'docx', limit: 50 }, NONE)
    expect(page.entries.map((e) => e.path)).toEqual(['/d/new.docx', '/d/legacy.doc'])
    expect(page.total).toBe(2)
    expect(page.totalAll).toBe(PATHS.length)
  })

  it('covers every extension the router opens under the xlsx filter', () => {
    const page = pageRecentPaths(PATHS, { ext: 'xlsx', limit: 50 }, NONE)
    expect(page.entries.map((e) => e.path)).toEqual([
      '/d/book.xlsx',
      '/d/macro.xlsm',
      '/d/old.xls',
      '/d/data.csv',
    ])
  })

  it('covers .md and .markdown under the md filter', () => {
    const page = pageRecentPaths(PATHS, { ext: 'md', limit: 50 }, NONE)
    expect(page.entries.map((e) => e.path)).toEqual(['/d/notes.md', '/d/readme.markdown'])
  })

  it('matches unknown extensions exactly', () => {
    const page = pageRecentPaths(PATHS, { ext: 'pptx', limit: 50 }, NONE)
    expect(page.entries.map((e) => e.path)).toEqual(['/d/deck.pptx'])
  })

  it('shares the family predicate with the starred list', () => {
    // The starred handler filters through matchesExtFamily so both lists agree.
    expect(matchesExtFamily('xlsx', 'csv')).toBe(true)
    expect(matchesExtFamily('md', 'markdown')).toBe(true)
    expect(matchesExtFamily('docx', 'doc')).toBe(true)
    expect(matchesExtFamily('pptx', 'ppt')).toBe(false)
    expect(matchesExtFamily(undefined, 'xls')).toBe(true)
  })
})
