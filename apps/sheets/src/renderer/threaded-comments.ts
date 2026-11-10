/**
 * Threaded (modern Excel) comments live outside Univer's note model: notes
 * are legacy comments with a yellow marker, threads carry an author, a time,
 * replies and a resolved flag. This store is the renderer's copy of every
 * sheet's threads; edits mark the sheet note-dirty and the save snapshots
 * the store next to the live notes (collectNoteStates).
 */
import { personIdFor } from '@genoffice/xlsx-gateway/gateway/xlsx-threaded-comments'

import type { WorkbookFile } from '../shared/desktop-api'
import type { StructuralJournalOp } from './edit-journal'

export interface ThreadReply {
  readonly id: string
  readonly personId: string
  readonly author: string
  readonly dT: string
  readonly text: string
}

export interface CellThread {
  readonly row: number
  readonly column: number
  readonly id: string
  readonly personId: string
  readonly author: string
  readonly dT: string
  readonly done: boolean
  readonly text: string
  readonly replies: readonly ThreadReply[]
}

export const FALLBACK_AUTHOR = 'GenOffice User'

export interface CellRange {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

/// Excel's dT shape: ISO without zone, two fractional digits.
export function threadTimestamp(date = new Date()): string {
  return date.toISOString().slice(0, 22)
}

export function parseThreadTimestamp(dT: string): Date | null {
  const date = new Date(/[zZ]|[+-]\d\d:\d\d$/.test(dT) ? dT : `${dT}Z`)
  return Number.isNaN(date.getTime()) ? null : date
}

export const cellKey = (row: number, column: number): string => `${row}:${column}`

const byPosition = (a: CellThread, b: CellThread): number => a.row - b.row || a.column - b.column

/// Reading order with wrap-around, like Excel's Previous/Next Comment.
export function neighbourThread(
  threads: readonly CellThread[],
  row: number,
  column: number,
  direction: 'next' | 'prev',
): CellThread | null {
  const sorted = [...threads].sort(byPosition)
  if (sorted.length === 0) return null
  if (direction === 'next') {
    return (
      sorted.find((thread) => thread.row > row || (thread.row === row && thread.column > column)) ??
      sorted[0]!
    )
  }
  return (
    [...sorted]
      .reverse()
      .find((thread) => thread.row < row || (thread.row === row && thread.column < column)) ??
    sorted[sorted.length - 1]!
  )
}

type Listener = () => void

export class ThreadStore {
  private threads = new Map<string, Map<string, CellThread>>()
  /// Authors of threads converted to notes this session, so the save can
  /// split the "Author:\n" marker off notes that never came from the file.
  private converted = new Map<string, Map<string, string>>()
  private listeners = new Set<Listener>()
  private version = 0
  private author = FALLBACK_AUTHOR
  /// Called with the sheet id after every edit (journal marking).
  onChange: ((sheetId: string) => void) | null = null

  subscribe = (listener: Listener): (() => void) => {
    this.listeners.add(listener)
    return () => void this.listeners.delete(listener)
  }

  getVersion = (): number => this.version

  setAuthor(name: string): void {
    const trimmed = name.trim()
    this.author = trimmed.length > 0 ? trimmed : FALLBACK_AUTHOR
  }

  authorName(): string {
    return this.author
  }

  /// Replaces the whole store with the file's threads (open / reload).
  load(file: Pick<WorkbookFile, 'sheets'>): void {
    this.threads = new Map()
    this.converted = new Map()
    for (const sheet of file.sheets) {
      for (const comment of sheet.comments) {
        if (!comment.thread) continue
        this.sheet(sheet.id).set(cellKey(comment.row, comment.column), {
          row: comment.row,
          column: comment.column,
          ...comment.thread,
        })
      }
    }
    this.bump(null)
  }

  sheetThreads(sheetId: string | null): CellThread[] {
    if (sheetId === null) return []
    return [...(this.threads.get(sheetId)?.values() ?? [])].sort(byPosition)
  }

  get(sheetId: string, row: number, column: number): CellThread | null {
    return this.threads.get(sheetId)?.get(cellKey(row, column)) ?? null
  }

  add(sheetId: string, row: number, column: number, text: string): CellThread {
    const thread: CellThread = {
      row,
      column,
      id: newGuid(),
      personId: personIdFor(this.author),
      author: this.author,
      dT: threadTimestamp(),
      done: false,
      text,
      replies: [],
    }
    this.sheet(sheetId).set(cellKey(row, column), thread)
    this.bump(sheetId)
    return thread
  }

  reply(sheetId: string, row: number, column: number, text: string): void {
    this.update(sheetId, row, column, (thread) => ({
      ...thread,
      replies: [
        ...thread.replies,
        {
          id: newGuid(),
          personId: personIdFor(this.author),
          author: this.author,
          dT: threadTimestamp(),
          text,
        },
      ],
    }))
  }

  setDone(sheetId: string, row: number, column: number, done: boolean): void {
    this.update(sheetId, row, column, (thread) => ({ ...thread, done }))
  }

  /// Edits the root (entryId = thread id) or one reply in place.
  editText(sheetId: string, row: number, column: number, entryId: string, text: string): void {
    this.update(sheetId, row, column, (thread) =>
      entryId === thread.id
        ? { ...thread, text }
        : {
            ...thread,
            replies: thread.replies.map((reply) =>
              reply.id === entryId ? { ...reply, text } : reply,
            ),
          },
    )
  }

  deleteReply(sheetId: string, row: number, column: number, replyId: string): void {
    this.update(sheetId, row, column, (thread) => ({
      ...thread,
      replies: thread.replies.filter((reply) => reply.id !== replyId),
    }))
  }

  remove(sheetId: string, row: number, column: number): CellThread | null {
    const sheet = this.threads.get(sheetId)
    const key = cellKey(row, column)
    const thread = sheet?.get(key) ?? null
    if (!sheet || !thread) return null
    sheet.delete(key)
    this.bump(sheetId)
    return thread
  }

  /// Converts a thread to a plain note: the thread leaves the store and the
  /// author is remembered for the note snapshot.
  convertToNote(sheetId: string, row: number, column: number): CellThread | null {
    const thread = this.remove(sheetId, row, column)
    if (!thread) return null
    let sheet = this.converted.get(sheetId)
    if (!sheet) {
      sheet = new Map()
      this.converted.set(sheetId, sheet)
    }
    sheet.set(cellKey(row, column), thread.author)
    return thread
  }

  convertedAuthor(sheetId: string, row: number, column: number): string | null {
    return this.converted.get(sheetId)?.get(cellKey(row, column)) ?? null
  }

  /// Duplicate Sheet: the copy starts with the source's threads (Excel
  /// copies them; the save mirrors the file parts the same way).
  copySheet(sourceId: string, targetId: string): void {
    const converted = this.converted.get(sourceId)
    if (converted && converted.size > 0) this.converted.set(targetId, new Map(converted))
    const source = this.threads.get(sourceId)
    if (!source || source.size === 0) return
    this.threads.set(targetId, new Map(source))
    this.bump(targetId)
  }

  /// Cut/paste or drag-move: threads inside the source range travel with
  /// their cells (Univer moves the notes itself); threads already at the
  /// destination are overwritten like the cell contents.
  moveRange(fromSheetId: string, from: CellRange, toSheetId: string, to: CellRange): void {
    const convertedFrom = this.converted.get(fromSheetId) ?? new Map<string, string>()
    const convertedTo =
      toSheetId === fromSheetId ? convertedFrom : (this.converted.get(toSheetId) ?? new Map())
    moveEntries(
      convertedFrom,
      convertedTo,
      from,
      to,
      (key) => key,
      () => null,
    )
    this.converted.set(toSheetId, convertedTo)
    if (toSheetId !== fromSheetId) this.converted.set(fromSheetId, convertedFrom)

    const source = this.threads.get(fromSheetId) ?? new Map<string, CellThread>()
    const target = this.sheet(toSheetId)
    const changed = moveEntries(
      source,
      target,
      from,
      to,
      (key) => key,
      (thread, row, column) => ({ ...thread, row, column }),
    )
    if (!changed) return
    this.bump(fromSheetId)
    if (toSheetId !== fromSheetId) this.bump(toSheetId)
  }

  removeSheet(sheetId: string): void {
    const had = this.threads.delete(sheetId)
    this.converted.delete(sheetId)
    if (had) this.bump(null)
  }

  /// Keeps threads on their cells through row/column inserts, deletes and
  /// whole-row moves (Univer moves the notes itself; the store mirrors it).
  applyStructuralOp(sheetId: string, op: StructuralJournalOp): void {
    const converted = this.converted.get(sheetId)
    if (converted) {
      const shifted = new Map<string, string>()
      for (const [key, author] of converted) {
        const [row, column] = key.split(':').map(Number) as [number, number]
        const next = shiftThread({ row, column } as CellThread, op)
        if (next) shifted.set(cellKey(next.row, next.column), author)
      }
      this.converted.set(sheetId, shifted)
    }
    const sheet = this.threads.get(sheetId)
    if (!sheet || sheet.size === 0) return
    const shifted = new Map<string, CellThread>()
    let changed = false
    for (const thread of sheet.values()) {
      const next = shiftThread(thread, op)
      if (next === null) {
        changed = true
        continue
      }
      if (next !== thread) changed = true
      shifted.set(cellKey(next.row, next.column), next)
    }
    if (!changed) return
    this.threads.set(sheetId, shifted)
    this.bump(sheetId)
  }

  private update(
    sheetId: string,
    row: number,
    column: number,
    change: (thread: CellThread) => CellThread,
  ): void {
    const sheet = this.threads.get(sheetId)
    const key = cellKey(row, column)
    const thread = sheet?.get(key)
    if (!sheet || !thread) return
    sheet.set(key, change(thread))
    this.bump(sheetId)
  }

  private sheet(sheetId: string): Map<string, CellThread> {
    let sheet = this.threads.get(sheetId)
    if (!sheet) {
      sheet = new Map()
      this.threads.set(sheetId, sheet)
    }
    return sheet
  }

  private bump(sheetId: string | null): void {
    this.version += 1
    if (sheetId !== null) this.onChange?.(sheetId)
    for (const listener of this.listeners) listener()
  }
}

/// Relocates the entries of a cell-keyed map inside `from` to `to` (clearing
/// `to` first); `relocate` returns the entry at its new cell, or null to keep
/// the value as is. True when anything changed.
function moveEntries<T>(
  source: Map<string, T>,
  target: Map<string, T>,
  from: CellRange,
  to: CellRange,
  keyOf: (key: string) => string,
  relocate: (value: T, row: number, column: number) => T | null,
): boolean {
  // Lift the moving entries first: on a same-sheet move the two maps are
  // one and the ranges may overlap.
  const moving: [number, number, T][] = []
  for (const [key, value] of [...source]) {
    const [row, column] = keyOf(key).split(':').map(Number) as [number, number]
    if (row < from.startRow || row > from.endRow) continue
    if (column < from.startColumn || column > from.endColumn) continue
    source.delete(key)
    moving.push([row, column, value])
  }
  let changed = moving.length > 0
  for (let row = to.startRow; row <= to.endRow; row += 1) {
    for (let column = to.startColumn; column <= to.endColumn; column += 1) {
      changed = target.delete(cellKey(row, column)) || changed
    }
  }
  for (const [row, column, value] of moving) {
    const nextRow = row - from.startRow + to.startRow
    const nextColumn = column - from.startColumn + to.startColumn
    target.set(cellKey(nextRow, nextColumn), relocate(value, nextRow, nextColumn) ?? value)
  }
  return changed
}

function shiftThread(thread: CellThread, op: StructuralJournalOp): CellThread | null {
  switch (op.kind) {
    case 'insert-rows':
      return thread.row >= op.index ? { ...thread, row: thread.row + op.count } : thread
    case 'insert-cols':
      return thread.column >= op.index ? { ...thread, column: thread.column + op.count } : thread
    case 'remove-rows':
      if (thread.row < op.index) return thread
      return thread.row < op.index + op.count ? null : { ...thread, row: thread.row - op.count }
    case 'remove-cols':
      if (thread.column < op.index) return thread
      return thread.column < op.index + op.count
        ? null
        : { ...thread, column: thread.column - op.count }
    case 'move-rows': {
      const { index, count, before } = op
      const row = thread.row
      if (row >= index && row < index + count) {
        return {
          ...thread,
          row: before > index ? row + (before - index - count) : row - (index - before),
        }
      }
      if (before > index && row >= index + count && row < before)
        return { ...thread, row: row - count }
      if (before < index && row >= before && row < index) return { ...thread, row: row + count }
      return thread
    }
    default:
      return thread
  }
}

function newGuid(): string {
  const uuid =
    typeof crypto !== 'undefined' && 'randomUUID' in crypto
      ? crypto.randomUUID()
      : Array.from({ length: 36 }, (_, i) =>
          i === 8 || i === 13 || i === 18 || i === 23
            ? '-'
            : Math.floor(Math.random() * 16).toString(16),
        ).join('')
  return `{${uuid.toUpperCase()}}`
}

export const threadStore = new ThreadStore()

/// The thread body as the sidecar / CLI display it: root then replies.
export function threadBodyText(thread: CellThread): string {
  return [thread.text, ...thread.replies.map((reply) => reply.text)].join('\n\n')
}

/// Pane visibility and the cell whose composer should open, shared between
/// the ribbon / context menu (writers) and the pane component (reader).
export interface ThreadPaneState {
  readonly open: boolean
  readonly draft: { readonly sheetId: string; readonly row: number; readonly column: number } | null
  readonly focus: { readonly sheetId: string; readonly row: number; readonly column: number } | null
}

class ThreadPaneController {
  private state: ThreadPaneState = { open: false, draft: null, focus: null }
  private listeners = new Set<Listener>()

  subscribe = (listener: Listener): (() => void) => {
    this.listeners.add(listener)
    return () => void this.listeners.delete(listener)
  }

  getState = (): ThreadPaneState => this.state

  toggle(): void {
    this.set({ open: !this.state.open, draft: null, focus: null })
  }

  close(): void {
    this.set({ open: false, draft: null, focus: null })
  }

  /// Opens the pane on a cell: a composer when it has no thread yet,
  /// otherwise scrolls that thread into view with its reply box focused.
  openAt(sheetId: string, row: number, column: number, compose: boolean): void {
    const cell = { sheetId, row, column }
    this.set({ open: true, draft: compose ? cell : null, focus: compose ? null : cell })
  }

  clearDraft(): void {
    this.set({ ...this.state, draft: null })
  }

  private set(state: ThreadPaneState): void {
    this.state = state
    for (const listener of this.listeners) listener()
  }
}

export const threadPane = new ThreadPaneController()
