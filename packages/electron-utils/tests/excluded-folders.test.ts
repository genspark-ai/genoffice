import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  isUnderExcludedFolder,
  normalizeExcludedFolders,
  readExcludedFolders,
} from '../src/excluded-folders'

const home = '/Users/someone'

describe('readExcludedFolders / normalizeExcludedFolders', () => {
  it('reads the setting: absolute strings only, resolved, once each', () => {
    expect(
      readExcludedFolders({
        excludedFolders: [join(home, 'tmp'), 'relative', 7, null, join(home, 'tmp')],
      }),
    ).toEqual([join(home, 'tmp')])
    expect(readExcludedFolders({})).toEqual([])
    expect(readExcludedFolders({ excludedFolders: 'nope' })).toEqual([])
  })

  it('normalize caps a hand-edited list', () => {
    const many = Array.from({ length: 300 }, (_, i) => join(home, `f-${i}`))
    expect(normalizeExcludedFolders(many)).toHaveLength(200)
  })
})

describe('isUnderExcludedFolder', () => {
  const folders = [join(home, 'notes'), join(home, 'archive', 'client-x')]

  it('matches the folder itself, files inside and subfolders', () => {
    expect(isUnderExcludedFolder(join(home, 'notes'), folders)).toBe(true)
    expect(isUnderExcludedFolder(join(home, 'notes', 'a.md'), folders)).toBe(true)
    expect(isUnderExcludedFolder(join(home, 'notes', 'sub', 'b.md'), folders)).toBe(true)
    expect(isUnderExcludedFolder(join(home, 'archive', 'client-x', 'c.docx'), folders)).toBe(true)
  })

  it('the boundary is the separator, never a bare prefix', () => {
    expect(isUnderExcludedFolder(join(home, 'notes-old'), folders)).toBe(false)
    expect(isUnderExcludedFolder(join(home, 'notes-old', 'x.md'), folders)).toBe(false)
    expect(isUnderExcludedFolder(join(home, 'archive', 'client-y'), folders)).toBe(false)
  })

  it('an empty exclusion list excludes nothing', () => {
    expect(isUnderExcludedFolder(join(home, 'notes'), [])).toBe(false)
  })
})
