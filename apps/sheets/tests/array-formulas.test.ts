import { describe, expect, it } from 'vitest'

import { collectArrayFollowers } from '../src/renderer/univer-sync'

type Cells = Parameters<typeof collectArrayFollowers>[1]

describe('collectArrayFollowers', () => {
  it('marks every covered cell except the master', () => {
    const followers = new Set<string>()
    const cells: Cells = [
      { row: 0, column: 0, value: 7, formula: '=TRANSPOSE(B1:C1)', arrayRef: 'A1:A2' },
      { row: 1, column: 0, value: 8 },
    ]
    collectArrayFollowers(followers, cells, [])
    expect([...followers]).toEqual(['1:0'])
  })

  it('leaves single-cell array formulas alone', () => {
    const followers = new Set<string>()
    collectArrayFollowers(
      followers,
      [{ row: 0, column: 2, value: 7, formula: '=B1&""', arrayRef: 'C1' }],
      [],
    )
    expect(followers.size).toBe(0)
  })

  it('ignores unparsable and oversized refs', () => {
    const followers = new Set<string>()
    collectArrayFollowers(
      followers,
      [
        { row: 0, column: 0, value: 1, formula: '=X', arrayRef: 'A:A' },
        { row: 0, column: 1, value: 1, formula: '=X', arrayRef: 'B1:B1000000' },
      ],
      [],
    )
    expect(followers.size).toBe(0)
  })

  it('maps follower coordinates through structural ops', () => {
    const followers = new Set<string>()
    collectArrayFollowers(
      followers,
      [{ row: 3, column: 0, value: 7, formula: '=X', arrayRef: 'A2:A3' }],
      [{ kind: 'insert-rows', index: 0, count: 2 }],
    )
    expect([...followers]).toEqual(['4:0'])
  })

  it('adds no followers when the array extent was fully deleted', () => {
    const followers = new Set<string>()
    // genspark-ai/genoffice#135 op stack: file rows 6-7 are both removed, but
    // box-envelope tracking reported screen row 8 (unrelated content) as a
    // survivor and blanked it.
    collectArrayFollowers(
      followers,
      [{ row: 5, column: 0, value: 7, formula: '=X', arrayRef: 'A7:A8' }],
      [
        { kind: 'move-rows', index: 11, count: 1, before: 7 },
        { kind: 'remove-rows', index: 8, count: 1 },
        { kind: 'move-rows', index: 10, count: 2, before: 2 },
        { kind: 'remove-rows', index: 8, count: 1 },
      ],
    )
    expect(followers.size).toBe(0)
  })

  it('keeps unrelated rows a move shuffled between the survivors out', () => {
    const followers = new Set<string>()
    // The move splits the extent: file rows 0-1 land on screen rows 0 and 4,
    // with unrelated file rows 2-4 sitting on screen 1-3 in between. The
    // envelope would blank all of 1-4; only screen row 4 is a real follower.
    collectArrayFollowers(
      followers,
      [{ row: 0, column: 0, value: 7, formula: '=X', arrayRef: 'A1:A2' }],
      [{ kind: 'move-rows', index: 2, count: 3, before: 1 }],
    )
    expect([...followers]).toEqual(['4:0'])
  })
})
