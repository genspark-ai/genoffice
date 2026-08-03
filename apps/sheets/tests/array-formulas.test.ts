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
})
