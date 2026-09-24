import { describe, expect, it } from 'vitest'
import type { PmNode } from '../src/renderer/editor/convert'
import {
  MAX_VERSIONS,
  clearRollback,
  markRolledBack,
  pushVersion,
  type DocVersion,
} from '../src/renderer/ai/version-history'

/** Distinct documents: only identity matters for the version bookkeeping. */
function doc(tag: string): PmNode {
  return { type: 'doc', attrs: { tag } }
}

/** `count` versions in a row, ids 1..count, each "before turn n". */
function versions(count: number): DocVersion[] {
  let list: DocVersion[] = []
  for (let i = 1; i <= count; i++) {
    list = pushVersion(list, {
      id: i,
      label: `turn ${i}`,
      time: `10:0${i}`,
      doc: doc(`before-${i}`),
    })
  }
  return list
}

const ids = (list: DocVersion[]) => list.map((v) => v.id)
const discarded = (list: DocVersion[]) => list.filter((v) => v.discarded).map((v) => v.id)

describe('pushVersion', () => {
  it('keeps one entry per turn, oldest first, and caps the ring', () => {
    const list = versions(MAX_VERSIONS + 5)
    expect(list).toHaveLength(MAX_VERSIONS)
    // the ring keeps the newest 20: ids 6..25
    expect(ids(list)).toEqual([
      6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25,
    ])
  })

  it('drops discarded versions and closes the undo window on the next edit', () => {
    // #543: a new run starts from the current document, so a past the user has
    // already rolled back out of is unreachable — and its snapshot is dead weight.
    let list = versions(4)
    list = markRolledBack(list, 2, doc('before-rollback'))
    expect(discarded(list)).toEqual([3, 4])

    list = pushVersion(list, { id: 5, label: 'turn 5', time: '10:05', doc: doc('before-5') })
    expect(ids(list)).toEqual([1, 2, 5])
    expect(discarded(list)).toEqual([])
    expect(list.find((v) => v.id === 2)?.rolledBack).toBeFalsy()
    expect(list.find((v) => v.id === 2)?.rolledBackFrom).toBeUndefined()
  })
})

describe('markRolledBack', () => {
  it('keeps the replaced document so the roll back can be undone', () => {
    const list = markRolledBack(versions(3), 2, doc('as-it-was'))
    expect(list[1].rolledBack).toBe(true)
    expect(list[1].rolledBackFrom).toEqual(doc('as-it-was'))
  })

  it('marks later versions discarded instead of destroying them', () => {
    // The whole point of #543: the snapshots have to survive, or an undo cannot
    // bring the rollback points back.
    const list = markRolledBack(versions(4), 2, doc('as-it-was'))
    expect(discarded(list)).toEqual([3, 4])
    expect(list[2].doc).toEqual(doc('before-3'))
    expect(list[3].doc).toEqual(doc('before-4'))
    expect(list[0].discarded).toBeFalsy()
    expect(list[1].discarded).toBeFalsy()
  })

  it('ignores an unknown id', () => {
    const list = versions(2)
    expect(markRolledBack(list, 99, doc('x'))).toBe(list)
  })
})

describe('clearRollback', () => {
  it('restores the document and every rollback point it had voided', () => {
    let list = markRolledBack(versions(4), 2, doc('as-it-was'))
    list = clearRollback(list, 2)
    expect(discarded(list)).toEqual([])
    expect(list[1].rolledBack).toBeFalsy()
    expect(list[1].rolledBackFrom).toBeUndefined()
    // and all four are restorable again
    expect(list.every((v) => v.doc)).toBe(true)
  })

  it('is a no-op unless that version is the one rolled back to', () => {
    const rolled = markRolledBack(versions(3), 2, doc('as-it-was'))
    expect(clearRollback(rolled, 3)).toBe(rolled) // 3 is discarded, not rolled back
    const plain = versions(3)
    expect(clearRollback(plain, 2)).toBe(plain) // nothing was rolled back
    expect(clearRollback(plain, 99)).toBe(plain)
  })

  it('round-trips: roll back, undo, roll back again', () => {
    let list = versions(3)
    list = markRolledBack(list, 1, doc('a'))
    expect(discarded(list)).toEqual([2, 3])
    list = clearRollback(list, 1)
    expect(discarded(list)).toEqual([])
    list = markRolledBack(list, 3, doc('b'))
    expect(list[2].rolledBackFrom).toEqual(doc('b'))
    expect(discarded(list)).toEqual([])
  })
})
