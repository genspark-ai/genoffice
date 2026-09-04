import { describe, expect, it } from 'vitest'

import { pageRecentPaths } from '../src/main/recent-files'

const NONE = new Set<string>()
const PATHS = ['/d/new.docx', '/d/legacy.doc', '/d/book.xlsx', '/d/macro.xlsm', '/d/deck.pptx']

describe('pageRecentPaths ext families', () => {
  it('includes legacy .doc under the docx filter', () => {
    const page = pageRecentPaths(PATHS, { ext: 'docx', limit: 50 }, NONE)
    expect(page.entries.map((e) => e.path)).toEqual(['/d/new.docx', '/d/legacy.doc'])
    expect(page.total).toBe(2)
    expect(page.totalAll).toBe(PATHS.length)
  })

  it('keeps the xlsx family as-is', () => {
    const page = pageRecentPaths(PATHS, { ext: 'xlsx', limit: 50 }, NONE)
    expect(page.entries.map((e) => e.path)).toEqual(['/d/book.xlsx', '/d/macro.xlsm'])
  })

  it('matches unknown extensions exactly', () => {
    const page = pageRecentPaths(PATHS, { ext: 'pptx', limit: 50 }, NONE)
    expect(page.entries.map((e) => e.path)).toEqual(['/d/deck.pptx'])
  })
})
