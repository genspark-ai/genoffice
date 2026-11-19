/**
 * The chunked save-edit transfer store: large saves upload their edits in
 * ordered slices which the following save request consumes as one array.
 * The store must enforce ordering, declared sizes, and session ownership,
 * and must not leak abandoned uploads.
 */
import { describe, expect, it } from 'vitest'
import { SaveEditsTransferStore } from '../src/main/save-edits-transfer'
import type { WorkbookCellEdit } from '../src/shared/desktop-api'

const SESSION = '5d4f6f7a-1c2b-4e3d-9a8f-0b1c2d3e4f5a'
const OTHER_SESSION = '6e5f7a8b-2d3c-4f5e-8b9a-1c2d3e4f5a6b'
const TRANSFER = '0f9e8d7c-6b5a-4c3d-8e2f-1a0b9c8d7e6f'

function edit(row: number, sheetId = 'sheet-1'): WorkbookCellEdit {
  return { sheetId, row, column: 0, writeValue: true, value: row }
}

function edits(count: number, startRow = 0): WorkbookCellEdit[] {
  return Array.from({ length: count }, (_, index) => edit(startRow + index))
}

describe('SaveEditsTransferStore', () => {
  it('reassembles ordered chunks into one edits array', () => {
    const store = new SaveEditsTransferStore()
    store.begin({ sessionId: SESSION, transferId: TRANSFER, total: 5 })
    store.addChunk({ sessionId: SESSION, transferId: TRANSFER, seq: 0, edits: edits(3, 0) })
    store.addChunk({ sessionId: SESSION, transferId: TRANSFER, seq: 1, edits: edits(2, 3) })
    const taken = store.take(TRANSFER, SESSION)
    expect(taken.map((entry) => entry.row)).toEqual([0, 1, 2, 3, 4])
    expect(store.openTransferCount).toBe(0)
    // Consumed: a second take must fail.
    expect(() => store.take(TRANSFER, SESSION)).toThrow(/Unknown save transfer/)
  })

  it('interns repeated sheet ids to one string instance', () => {
    const store = new SaveEditsTransferStore()
    store.begin({ sessionId: SESSION, transferId: TRANSFER, total: 2 })
    // Simulate structured clone: equal but distinct string instances.
    const first = edit(0, 'Sheet1'.split('').join(''))
    const second = edit(1, 'Sheet1'.split('').join(''))
    store.addChunk({ sessionId: SESSION, transferId: TRANSFER, seq: 0, edits: [first, second] })
    const taken = store.take(TRANSFER, SESSION)
    expect(taken[0]?.sheetId).toBe(taken[1]?.sheetId)
  })

  it('rejects out-of-order chunks and drops the transfer', () => {
    const store = new SaveEditsTransferStore()
    store.begin({ sessionId: SESSION, transferId: TRANSFER, total: 4 })
    store.addChunk({ sessionId: SESSION, transferId: TRANSFER, seq: 0, edits: edits(2) })
    expect(() =>
      store.addChunk({ sessionId: SESSION, transferId: TRANSFER, seq: 2, edits: edits(2) }),
    ).toThrow(/out of order/)
    expect(store.openTransferCount).toBe(0)
  })

  it('rejects chunks that exceed the declared total', () => {
    const store = new SaveEditsTransferStore()
    store.begin({ sessionId: SESSION, transferId: TRANSFER, total: 3 })
    expect(() =>
      store.addChunk({ sessionId: SESSION, transferId: TRANSFER, seq: 0, edits: edits(4) }),
    ).toThrow(/declared size/)
  })

  it('refuses to hand out an incomplete transfer', () => {
    const store = new SaveEditsTransferStore()
    store.begin({ sessionId: SESSION, transferId: TRANSFER, total: 4 })
    store.addChunk({ sessionId: SESSION, transferId: TRANSFER, seq: 0, edits: edits(2) })
    expect(() => store.take(TRANSFER, SESSION)).toThrow(/incomplete/)
    expect(store.openTransferCount).toBe(0)
  })

  it('scopes transfers to their session', () => {
    const store = new SaveEditsTransferStore()
    store.begin({ sessionId: SESSION, transferId: TRANSFER, total: 1 })
    expect(() =>
      store.addChunk({ sessionId: OTHER_SESSION, transferId: TRANSFER, seq: 0, edits: edits(1) }),
    ).toThrow(/Unknown save transfer/)
    store.addChunk({ sessionId: SESSION, transferId: TRANSFER, seq: 0, edits: edits(1) })
    expect(() => store.take(TRANSFER, OTHER_SESSION)).toThrow(/Unknown save transfer/)
    store.discardSession(SESSION)
    expect(store.openTransferCount).toBe(0)
  })

  it('expires idle transfers on sweep', () => {
    const store = new SaveEditsTransferStore()
    const start = 1_000_000
    store.begin({ sessionId: SESSION, transferId: TRANSFER, total: 4 }, start)
    store.addChunk({ sessionId: SESSION, transferId: TRANSFER, seq: 0, edits: edits(2) }, start)
    store.sweep(start + 121_000)
    expect(store.openTransferCount).toBe(0)
  })

  it('expires idle transfers from chunk activity of other transfers', () => {
    const store = new SaveEditsTransferStore()
    const start = 1_000_000
    const other = '7f6e5d4c-3b2a-4d5e-9c8b-2d3e4f5a6b7c'
    store.begin({ sessionId: SESSION, transferId: TRANSFER, total: 4 }, start)
    store.begin({ sessionId: SESSION, transferId: other, total: 2 }, start)
    store.addChunk(
      { sessionId: SESSION, transferId: other, seq: 0, edits: edits(1) },
      start + 60_000,
    )
    // This chunk's sweep reaps TRANSFER (idle 121s) but keeps its own
    // still-active transfer (idle 61s).
    store.addChunk(
      { sessionId: SESSION, transferId: other, seq: 1, edits: edits(1) },
      start + 121_000,
    )
    expect(store.openTransferCount).toBe(1)
  })

  it('discards on request and stays silent for unknown or foreign transfers', () => {
    const store = new SaveEditsTransferStore()
    store.begin({ sessionId: SESSION, transferId: TRANSFER, total: 2 })
    store.addChunk({ sessionId: SESSION, transferId: TRANSFER, seq: 0, edits: edits(1) })
    // Wrong session and unknown ids are no-ops, not errors.
    store.discard(TRANSFER, OTHER_SESSION)
    expect(store.openTransferCount).toBe(1)
    store.discard(TRANSFER, SESSION)
    expect(store.openTransferCount).toBe(0)
    expect(() => store.discard(TRANSFER, SESSION)).not.toThrow()
  })

  it('dispose drops all transfers immediately', () => {
    const store = new SaveEditsTransferStore()
    store.begin({ sessionId: SESSION, transferId: TRANSFER, total: 2 })
    store.addChunk({ sessionId: SESSION, transferId: TRANSFER, seq: 0, edits: edits(1) })
    store.dispose()
    expect(store.openTransferCount).toBe(0)
    expect(() => store.take(TRANSFER, SESSION)).toThrow(/Unknown save transfer/)
  })

  it('caps how many transfers stay open at once', () => {
    const store = new SaveEditsTransferStore()
    for (let index = 0; index < 4; index += 1) {
      store.begin({ sessionId: SESSION, transferId: `${TRANSFER.slice(0, -1)}${index}`, total: 1 })
    }
    expect(() =>
      store.begin({ sessionId: SESSION, transferId: `${TRANSFER.slice(0, -1)}f`, total: 1 }),
    ).toThrow(/Too many open save transfers/)
  })
})
