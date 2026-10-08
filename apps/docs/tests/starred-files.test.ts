import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  assignStarredGroup,
  dropStarredItems,
  parseStarredItems,
  readStarredItems,
  renameStarredItem,
  starredGroupMap,
  starredGroupNames,
  toggleStarredItem,
  writeStarredItems,
} from '../src/main/starred-files'

/**
 * starred.json (home-screen favorites) lives in userData and must survive app
 * upgrades: the legacy flat string[] keeps every star when read, and the
 * versioned group shape round-trips. Path is the key everywhere — toggle,
 * remove, rename and grouping all work on paths, never on positions.
 */

const LEGACY = ['/a/合同.docx', '/b/report.xlsx']

describe('parseStarredItems (the migration boundary)', () => {
  it('reads the legacy flat string[] into ungrouped entries', () => {
    expect(parseStarredItems(LEGACY)).toEqual([{ path: LEGACY[0] }, { path: LEGACY[1] }])
  })

  it('reads the versioned group shape and trims group names', () => {
    expect(
      parseStarredItems({
        version: 1,
        items: [
          { path: '/a.docx', group: ' 合同 ' },
          { path: '/b.xlsx' },
          { path: '/c.xlsx', group: '   ' },
        ],
      }),
    ).toEqual([{ path: '/a.docx', group: '合同' }, { path: '/b.xlsx' }, { path: '/c.xlsx' }])
  })

  it('degrades corrupt or foreign payloads to an empty store', () => {
    expect(parseStarredItems(null)).toEqual([])
    expect(parseStarredItems(42)).toEqual([])
    expect(parseStarredItems({ version: 1 })).toEqual([])
    expect(parseStarredItems([{ nope: true }, 7, { path: '/ok.docx' }])).toEqual([
      { path: '/ok.docx' },
    ])
  })
})

describe('file round-trip', () => {
  it('writes the versioned shape and reads it back', () => {
    const file = join(mkdtempSync(join(tmpdir(), 'genoffice-starred-')), 'starred.json')
    writeStarredItems(file, [{ path: LEGACY[0], group: '合同' }, { path: LEGACY[1] }])
    const onDisk = JSON.parse(readFileSync(file, 'utf-8')) as { version: number; items: unknown[] }
    expect(onDisk.version).toBe(1)
    expect(onDisk.items).toHaveLength(2)
    expect(readStarredItems(file)).toEqual([
      { path: LEGACY[0], group: '合同' },
      { path: LEGACY[1] },
    ])
  })

  it('migrates a legacy file on read; the next write persists the new shape', () => {
    const file = join(mkdtempSync(join(tmpdir(), 'genoffice-starred-')), 'starred.json')
    writeFileSync(file, JSON.stringify(LEGACY), 'utf-8')
    expect(readStarredItems(file)).toEqual([{ path: LEGACY[0] }, { path: LEGACY[1] }])
    const next = toggleStarredItem(readStarredItems(file), '/c/new.md')
    writeStarredItems(file, next)
    expect(readStarredItems(file)).toEqual([
      { path: LEGACY[0] },
      { path: LEGACY[1] },
      { path: '/c/new.md' },
    ])
  })
})

describe('toggle / drop / rename (path is the key)', () => {
  const items = [{ path: '/a.docx', group: '合同' }, { path: '/b.xlsx' }]

  it('toggle stars an unstarred path and unstars a starred one (group goes with it)', () => {
    expect(toggleStarredItem(items, '/c/new.md')).toEqual([...items, { path: '/c/new.md' }])
    expect(toggleStarredItem(items, '/a.docx')).toEqual([{ path: '/b.xlsx' }])
  })

  it('drop removes only the listed paths and reports nothing-to-do as null', () => {
    expect(dropStarredItems(items, ['/a.docx', '/gone'])).toEqual([{ path: '/b.xlsx' }])
    expect(dropStarredItems(items, ['/gone'])).toBeNull()
    expect(dropStarredItems(items, [])).toBeNull()
  })

  it('rename re-keys the entry and keeps its group; absent path is null', () => {
    expect(renameStarredItem(items, '/a.docx', '/moved/a.docx')).toEqual([
      { path: '/moved/a.docx', group: '合同' },
      { path: '/b.xlsx' },
    ])
    expect(renameStarredItem(items, '/nope', '/x')).toBeNull()
  })
})

describe('groups', () => {
  const items = [
    { path: '/a.docx', group: '合同' },
    { path: '/b.xlsx', group: '项目资料' },
    { path: '/c.md', group: '合同' },
    { path: '/d.docx' },
  ]

  it('assign moves files into a group and null clears it; no-change is null', () => {
    expect(assignStarredGroup(items, ['/d.docx'], '合同')).toEqual([
      { path: '/a.docx', group: '合同' },
      { path: '/b.xlsx', group: '项目资料' },
      { path: '/c.md', group: '合同' },
      { path: '/d.docx', group: '合同' },
    ])
    expect(assignStarredGroup(items, ['/a.docx'], null)).toEqual([
      { path: '/a.docx' },
      { path: '/b.xlsx', group: '项目资料' },
      { path: '/c.md', group: '合同' },
      { path: '/d.docx' },
    ])
    expect(assignStarredGroup(items, ['/a.docx'], '合同')).toBeNull()
    expect(assignStarredGroup(items, [], '合同')).toBeNull()
    // whitespace-only names are treated as "no group"
    expect(assignStarredGroup(items, ['/a.docx'], '   ')).toEqual([
      { path: '/a.docx' },
      { path: '/b.xlsx', group: '项目资料' },
      { path: '/c.md', group: '合同' },
      { path: '/d.docx' },
    ])
  })

  it('names come in first-seen order; the map covers grouped paths only', () => {
    expect(starredGroupNames(items)).toEqual(['合同', '项目资料'])
    expect(starredGroupMap(items)).toEqual(
      new Map([
        ['/a.docx', '合同'],
        ['/b.xlsx', '项目资料'],
        ['/c.md', '合同'],
      ]),
    )
    // a group vanishes with its last member — no separate cleanup step
    const withoutContracts = dropStarredItems(items, ['/a.docx', '/c.md'])
    expect(starredGroupNames(withoutContracts!)).toEqual(['项目资料'])
  })
})
