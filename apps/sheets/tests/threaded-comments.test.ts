import { describe, expect, it } from 'vitest'
import { personIdFor } from '@genoffice/xlsx-gateway/gateway/xlsx-threaded-comments'

import {
  FALLBACK_AUTHOR,
  ThreadStore,
  neighbourThread,
  parseThreadTimestamp,
  threadTimestamp,
  type CellThread,
} from '../src/renderer/threaded-comments'

const fileThread = {
  id: '{R1}',
  personId: '{P1}',
  author: 'Ada',
  dT: '2024-01-01T00:00:00.00',
  done: false,
  text: 'Root',
  replies: [
    { id: '{R2}', personId: '{P1}', author: 'Ada', dT: '2024-01-02T00:00:00.00', text: 'Reply' },
  ],
}

function storeWithFile() {
  const store = new ThreadStore()
  const changed: string[] = []
  store.onChange = (sheetId) => changed.push(sheetId)
  store.load({
    sheets: [
      {
        id: 's1',
        comments: [
          { row: 0, column: 0, author: 'Ada', text: 'Root\n\nReply', thread: fileThread },
          { row: 3, column: 1, author: 'Bob', text: 'plain note' },
        ],
      },
    ],
  } as never)
  return { store, changed }
}

describe('ThreadStore', () => {
  it('loads only threaded comments and never marks a sheet dirty for the load', () => {
    const { store, changed } = storeWithFile()
    expect(store.sheetThreads('s1').map((thread) => [thread.row, thread.column])).toEqual([[0, 0]])
    expect(store.get('s1', 0, 0)?.text).toBe('Root')
    expect(store.get('s1', 3, 1)).toBeNull()
    expect(changed).toEqual([])
  })

  it('adds a thread under the current author with a stable person id and marks the sheet', () => {
    const { store, changed } = storeWithFile()
    store.setAuthor('  ')
    expect(store.authorName()).toBe(FALLBACK_AUTHOR)
    store.setAuthor('Grace')
    const thread = store.add('s1', 5, 5, 'Hello')
    expect(thread.author).toBe('Grace')
    expect(thread.personId).toBe(personIdFor('Grace'))
    expect(thread.id).toMatch(/^\{[0-9A-F-]{36}\}$/)
    expect(thread.done).toBe(false)
    expect(parseThreadTimestamp(thread.dT)).not.toBeNull()
    expect(changed).toEqual(['s1'])
  })

  it('replies, resolves, edits and deletes replies in place', () => {
    const { store, changed } = storeWithFile()
    store.setAuthor('Grace')
    store.reply('s1', 0, 0, 'Second')
    expect(store.get('s1', 0, 0)?.replies.map((reply) => reply.text)).toEqual(['Reply', 'Second'])
    expect(store.get('s1', 0, 0)?.replies[1]?.author).toBe('Grace')
    store.setDone('s1', 0, 0, true)
    expect(store.get('s1', 0, 0)?.done).toBe(true)
    store.editText('s1', 0, 0, '{R1}', 'Root edited')
    store.editText('s1', 0, 0, '{R2}', 'Reply edited')
    expect(store.get('s1', 0, 0)?.text).toBe('Root edited')
    expect(store.get('s1', 0, 0)?.replies[0]?.text).toBe('Reply edited')
    store.deleteReply('s1', 0, 0, '{R2}')
    expect(store.get('s1', 0, 0)?.replies.map((reply) => reply.text)).toEqual(['Second'])
    expect(store.get('s1', 0, 0)?.id).toBe('{R1}')
    expect(changed.length).toBe(5)
  })

  it('removes a thread and ignores edits on cells without one', () => {
    const { store, changed } = storeWithFile()
    expect(store.remove('s1', 0, 0)?.id).toBe('{R1}')
    expect(store.remove('s1', 0, 0)).toBeNull()
    store.reply('s1', 9, 9, 'nowhere')
    expect(store.sheetThreads('s1')).toEqual([])
    expect(changed).toEqual(['s1'])
  })

  it('notifies subscribers once per change', () => {
    const { store } = storeWithFile()
    let ticks = 0
    const unsubscribe = store.subscribe(() => (ticks += 1))
    store.add('s1', 1, 1, 'a')
    store.setDone('s1', 1, 1, true)
    unsubscribe()
    store.remove('s1', 1, 1)
    expect(ticks).toBe(2)
  })
})

describe('neighbourThread', () => {
  const at = (row: number, column: number) => ({ row, column }) as CellThread
  const threads = [at(5, 0), at(0, 2), at(0, 0)]

  it('walks reading order with wrap-around in both directions', () => {
    expect(neighbourThread(threads, 0, 0, 'next')).toEqual(at(0, 2))
    expect(neighbourThread(threads, 0, 2, 'next')).toEqual(at(5, 0))
    expect(neighbourThread(threads, 5, 0, 'next')).toEqual(at(0, 0))
    expect(neighbourThread(threads, 0, 0, 'prev')).toEqual(at(5, 0))
    expect(neighbourThread(threads, 3, 3, 'prev')).toEqual(at(0, 2))
    expect(neighbourThread([], 0, 0, 'next')).toBeNull()
  })
})

describe('threadTimestamp', () => {
  it('writes the Excel dT shape and reads it back as UTC', () => {
    const stamp = threadTimestamp(new Date(Date.UTC(2024, 4, 6, 7, 8, 9, 100)))
    expect(stamp).toBe('2024-05-06T07:08:09.10')
    expect(parseThreadTimestamp(stamp)?.toISOString()).toBe('2024-05-06T07:08:09.100Z')
    expect(parseThreadTimestamp('garbage')).toBeNull()
  })
})

describe('ThreadStore structural ops and conversion', () => {
  function storeAt(cells: [number, number][]) {
    const store = new ThreadStore()
    store.setAuthor('Ada')
    for (const [row, column] of cells) store.add('s1', row, column, `${row}:${column}`)
    const positions = () => store.sheetThreads('s1').map((thread) => [thread.row, thread.column])
    return { store, positions }
  }

  it('shifts threads on row/column inserts and drops those in removed lines', () => {
    const { store, positions } = storeAt([
      [0, 0],
      [2, 1],
      [5, 3],
    ])
    store.applyStructuralOp('s1', { kind: 'insert-rows', index: 1, count: 2 })
    expect(positions()).toEqual([
      [0, 0],
      [4, 1],
      [7, 3],
    ])
    store.applyStructuralOp('s1', { kind: 'remove-rows', index: 3, count: 2 })
    expect(positions()).toEqual([
      [0, 0],
      [5, 3],
    ])
    store.applyStructuralOp('s1', { kind: 'insert-cols', index: 1, count: 1 })
    store.applyStructuralOp('s1', { kind: 'remove-cols', index: 0, count: 1 })
    expect(positions()).toEqual([[5, 3]])
    expect(store.get('s1', 5, 3)?.text).toBe('5:3')
  })

  it('follows whole-row moves in both directions', () => {
    const { store, positions } = storeAt([
      [1, 0],
      [4, 0],
    ])
    store.applyStructuralOp('s1', { kind: 'move-rows', index: 1, count: 1, before: 6 })
    expect(positions()).toEqual([
      [3, 0],
      [5, 0],
    ])
    store.applyStructuralOp('s1', { kind: 'move-rows', index: 5, count: 1, before: 0 })
    expect(positions()).toEqual([
      [0, 0],
      [4, 0],
    ])
  })

  it('remembers the author of a converted thread until the file reloads', () => {
    const { store } = storeAt([[2, 2]])
    expect(store.convertToNote('s1', 2, 2)?.author).toBe('Ada')
    expect(store.get('s1', 2, 2)).toBeNull()
    expect(store.convertedAuthor('s1', 2, 2)).toBe('Ada')
    store.applyStructuralOp('s1', { kind: 'insert-rows', index: 0, count: 3 })
    expect(store.convertedAuthor('s1', 2, 2)).toBeNull()
    expect(store.convertedAuthor('s1', 5, 2)).toBe('Ada')
    store.applyStructuralOp('s1', { kind: 'remove-cols', index: 2, count: 1 })
    expect(store.convertedAuthor('s1', 5, 2)).toBeNull()
    store.load({ sheets: [] } as never)
    expect(store.convertedAuthor('s1', 2, 2)).toBeNull()
  })
})

describe('ThreadStore sheet copy and range moves', () => {
  it("copies a sheet's threads onto the duplicate", () => {
    const store = new ThreadStore()
    store.setAuthor('Ada')
    store.add('s1', 1, 1, 'a')
    store.copySheet('s1', 's2')
    store.copySheet('empty', 's3')
    expect(store.get('s2', 1, 1)?.text).toBe('a')
    expect(store.sheetThreads('s3')).toEqual([])
    store.reply('s2', 1, 1, 'only on the copy')
    expect(store.get('s1', 1, 1)?.replies).toEqual([])
  })

  it('moves threads with a cut range and overwrites the destination', () => {
    const store = new ThreadStore()
    store.setAuthor('Ada')
    store.add('s1', 0, 0, 'moving')
    store.add('s1', 0, 1, 'moving too')
    store.add('s1', 5, 0, 'stays')
    store.add('s1', 9, 9, 'overwritten')
    store.moveRange('s1', { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }, 's2', {
      startRow: 9,
      endRow: 9,
      startColumn: 8,
      endColumn: 9,
    })
    expect(store.sheetThreads('s1').map((thread) => thread.text)).toEqual(['stays', 'overwritten'])
    expect(
      store.sheetThreads('s2').map((thread) => [thread.row, thread.column, thread.text]),
    ).toEqual([
      [9, 8, 'moving'],
      [9, 9, 'moving too'],
    ])
    store.moveRange('s1', { startRow: 5, endRow: 5, startColumn: 0, endColumn: 0 }, 's1', {
      startRow: 9,
      endRow: 9,
      startColumn: 9,
      endColumn: 9,
    })
    expect(store.sheetThreads('s1').map((thread) => [thread.row, thread.text])).toEqual([
      [9, 'stays'],
    ])
  })

  it('clears destination threads even when the source sheet has none, and moves converted authors', () => {
    const store = new ThreadStore()
    store.setAuthor('Ada')
    store.add('s2', 3, 3, 'to be overwritten')
    store.add('s1', 0, 0, 'converted')
    store.convertToNote('s1', 0, 0)
    store.moveRange('fresh', { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, 's2', {
      startRow: 3,
      endRow: 3,
      startColumn: 3,
      endColumn: 3,
    })
    expect(store.sheetThreads('s2')).toEqual([])
    store.moveRange('s1', { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, 's2', {
      startRow: 7,
      endRow: 7,
      startColumn: 1,
      endColumn: 1,
    })
    expect(store.convertedAuthor('s1', 0, 0)).toBeNull()
    expect(store.convertedAuthor('s2', 7, 1)).toBe('Ada')
    store.copySheet('s2', 's3')
    expect(store.convertedAuthor('s3', 7, 1)).toBe('Ada')
  })

  it('keeps every thread on an overlapping same-sheet move', () => {
    const store = new ThreadStore()
    store.setAuthor('Ada')
    store.add('s1', 0, 0, 'a')
    store.add('s1', 1, 0, 'b')
    store.add('s1', 2, 0, 'c')
    store.moveRange('s1', { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }, 's1', {
      startRow: 1,
      endRow: 3,
      startColumn: 0,
      endColumn: 0,
    })
    expect(store.sheetThreads('s1').map((thread) => [thread.row, thread.text])).toEqual([
      [1, 'a'],
      [2, 'b'],
      [3, 'c'],
    ])
  })
})
