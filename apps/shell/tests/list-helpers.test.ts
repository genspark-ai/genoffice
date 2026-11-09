import { describe, expect, it } from 'vitest'
import type { RecentEntry } from '../src/shared/home-api'
import {
  isEditableTarget,
  isOpenablePath,
  nextRowIndex,
  rangeSelect,
  sortEntries,
} from '../src/renderer/src/list-helpers'

const entry = (name: string, sizeBytes: number, mtimeMs: number): RecentEntry => ({
  path: `/f/${name}`,
  name,
  ext: name.split('.').pop() ?? '',
  mtimeMs,
  sizeBytes,
  starred: false,
})

describe('rangeSelect', () => {
  const paths = ['a', 'b', 'c', 'd', 'e']
  it('selects inclusively in either direction', () => {
    expect(rangeSelect(paths, 'b', 'd')).toEqual(['b', 'c', 'd'])
    expect(rangeSelect(paths, 'd', 'b')).toEqual(['b', 'c', 'd'])
  })
  it('falls back to the target alone without a usable anchor', () => {
    expect(rangeSelect(paths, null, 'c')).toEqual(['c'])
    expect(rangeSelect(paths, 'zzz', 'c')).toEqual(['c'])
    expect(rangeSelect(paths, 'a', 'missing')).toEqual([])
  })
})

describe('nextRowIndex', () => {
  it('clamps at both ends', () => {
    expect(nextRowIndex('ArrowDown', 4, 5)).toBe(4)
    expect(nextRowIndex('ArrowUp', 0, 5)).toBe(0)
  })
  it('starts at the top from nothing focused', () => {
    expect(nextRowIndex('ArrowDown', -1, 5)).toBe(0)
    expect(nextRowIndex('ArrowUp', -1, 5)).toBe(0)
  })
  it('handles Home / End and empty lists', () => {
    expect(nextRowIndex('Home', 3, 5)).toBe(0)
    expect(nextRowIndex('End', 1, 5)).toBe(4)
    expect(nextRowIndex('ArrowDown', -1, 0)).toBe(-1)
  })
})

describe('isEditableTarget', () => {
  it('treats text fields as editable but not checkboxes', () => {
    expect(isEditableTarget({ tagName: 'INPUT', type: 'text' } as unknown as EventTarget)).toBe(
      true,
    )
    expect(isEditableTarget({ tagName: 'INPUT', type: 'checkbox' } as unknown as EventTarget)).toBe(
      false,
    )
    expect(isEditableTarget({ tagName: 'TEXTAREA' } as unknown as EventTarget)).toBe(true)
    expect(
      isEditableTarget({ tagName: 'DIV', isContentEditable: true } as unknown as EventTarget),
    ).toBe(true)
    expect(isEditableTarget({ tagName: 'DIV' } as unknown as EventTarget)).toBe(false)
    expect(isEditableTarget(null)).toBe(false)
  })
})

describe('sortEntries', () => {
  const list = [entry('b.docx', 10, 3), entry('a10.docx', 300, 2), entry('a2.docx', 20, 1)]
  it('keeps source order for recent and reverses for oldest', () => {
    expect(sortEntries(list, 'recent').map((e) => e.name)).toEqual([
      'b.docx',
      'a10.docx',
      'a2.docx',
    ])
    expect(sortEntries(list, 'oldest').map((e) => e.name)).toEqual([
      'a2.docx',
      'a10.docx',
      'b.docx',
    ])
  })
  it('sorts names naturally (a2 before a10)', () => {
    expect(sortEntries(list, 'nameAsc').map((e) => e.name)).toEqual([
      'a2.docx',
      'a10.docx',
      'b.docx',
    ])
    expect(sortEntries(list, 'nameDesc').map((e) => e.name)).toEqual([
      'b.docx',
      'a10.docx',
      'a2.docx',
    ])
  })
  it('sorts by size both ways', () => {
    expect(sortEntries(list, 'sizeDesc').map((e) => e.sizeBytes)).toEqual([300, 20, 10])
    expect(sortEntries(list, 'sizeAsc').map((e) => e.sizeBytes)).toEqual([10, 20, 300])
  })
  it('does not mutate its input', () => {
    const before = list.map((e) => e.name)
    sortEntries(list, 'nameAsc')
    expect(list.map((e) => e.name)).toEqual(before)
  })
})

describe('isOpenablePath', () => {
  it('accepts supported document types case-insensitively', () => {
    expect(isOpenablePath('C:\\x\\Report.DOCX')).toBe(true)
    expect(isOpenablePath('/a/b.md')).toBe(true)
    expect(isOpenablePath('/a/b.csv')).toBe(true)
    expect(isOpenablePath('/a/b.tsv')).toBe(true)
    expect(isOpenablePath('/a/data.json')).toBe(true)
    expect(isOpenablePath('/a/doc.xml')).toBe(true)
    expect(isOpenablePath('/a/config.yaml')).toBe(true)
    expect(isOpenablePath('/a/query.sql')).toBe(true)
    expect(isOpenablePath('/a/notes.txt')).toBe(true)
  })
  it('rejects everything else', () => {
    expect(isOpenablePath('/a/run.exe')).toBe(false)
    expect(isOpenablePath('/a/noext')).toBe(false)
  })
})
