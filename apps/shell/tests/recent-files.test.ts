import { describe, expect, it } from 'vitest'
import {
  capStatPaths,
  matchesExtFamily,
  normalizeRecentQuery,
  STAT_PATHS_MAX,
} from '../src/main/recent-files'

describe('matchesExtFamily', () => {
  it('maps sidebar filter keys onto their extension families', () => {
    expect(matchesExtFamily('doc', 'docx')).toBe(true)
    expect(matchesExtFamily('docx', 'docx')).toBe(true)
    expect(matchesExtFamily('ppt', 'pptx')).toBe(true)
    expect(matchesExtFamily('pptx', 'pptx')).toBe(true)
    expect(matchesExtFamily('markdown', 'md')).toBe(true)
    expect(matchesExtFamily('md', 'md')).toBe(true)
    expect(matchesExtFamily('csv', 'xlsx')).toBe(true)
    expect(matchesExtFamily('htm', 'html')).toBe(true)
  })

  it('still matches exact extensions and rejects outsiders', () => {
    expect(matchesExtFamily('pdf', 'pdf')).toBe(true)
    expect(matchesExtFamily('xlsx', 'docx')).toBe(false)
    expect(matchesExtFamily('png', 'xlsx')).toBe(false)
  })
})

describe('normalizeRecentQuery', () => {
  it('defaults offset/limit and normalizes the extension key', () => {
    expect(normalizeRecentQuery({ ext: '.XLSX ' })).toEqual({ offset: 0, limit: 50, ext: 'xlsx' })
    expect(normalizeRecentQuery({ offset: 3, limit: 500 })).toEqual({
      offset: 3,
      limit: 200,
    })
  })
})

describe('capStatPaths', () => {
  const paths = (n: number) => Array.from({ length: n }, (_, i) => `/f${i}.docx`)

  it('leaves a list at the cap untouched', () => {
    expect(capStatPaths(paths(0))).toEqual([])
    expect(capStatPaths(paths(STAT_PATHS_MAX - 1))).toHaveLength(STAT_PATHS_MAX - 1)
    expect(capStatPaths(paths(STAT_PATHS_MAX))).toHaveLength(STAT_PATHS_MAX)
  })

  it('truncates past the cap, keeping the first paths in order', () => {
    const capped = capStatPaths(paths(STAT_PATHS_MAX + 500))
    expect(capped).toHaveLength(STAT_PATHS_MAX)
    expect(capped[0]).toBe('/f0.docx')
    expect(capped.at(-1)).toBe(`/f${STAT_PATHS_MAX - 1}.docx`)
  })
})
