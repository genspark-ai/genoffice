import { describe, expect, it } from 'vitest'

import {
  lazyFileSheetId,
  lazySheetMeta,
  lazySheetScreenExtent,
  type LazyWorkbookState,
} from '../src/renderer/univer-state'

function state(aliases: Array<[string, string]>): LazyWorkbookState {
  return {
    file: { sheets: [{ id: 'file-a', rowCount: 1200, columnCount: 120 }] },
    streamAliases: new Map(aliases),
    editJournal: {
      structuralOps: new Map([['dup-2', [{ kind: 'insert-rows', index: 0, count: 5 }]]]),
    },
  } as unknown as LazyWorkbookState
}

describe('stream aliases for sheets duplicated while streaming', () => {
  it('resolves a duplicate chain back to the file sheet', () => {
    const s = state([
      ['dup-1', 'file-a'],
      ['dup-2', 'dup-1'],
    ])
    expect(lazyFileSheetId(s, 'dup-2')).toBe('file-a')
    expect(lazyFileSheetId(s, 'file-a')).toBe('file-a')
    expect(lazyFileSheetId(s, 'blank-added')).toBe('blank-added')
    expect(lazySheetMeta(s, 'dup-2')?.id).toBe('file-a')
    expect(lazySheetMeta(s, 'blank-added')).toBeUndefined()
  })

  it('gives the duplicate its own screen extent over the source geometry', () => {
    const s = state([['dup-2', 'file-a']])
    // the copy's own structural ops shift its extent, not the source's
    expect(lazySheetScreenExtent(s, 'dup-2')).toEqual({ rows: 1205, columns: 120 })
    expect(lazySheetScreenExtent(s, 'file-a')).toEqual({ rows: 1200, columns: 120 })
  })

  it('tolerates a state without the alias map', () => {
    const s = { file: { sheets: [] } } as unknown as LazyWorkbookState
    expect(lazyFileSheetId(s, 'x')).toBe('x')
  })
})
