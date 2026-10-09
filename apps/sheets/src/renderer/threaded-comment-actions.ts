/**
 * Review ▸ Comments / Notes ribbon commands and their context-menu twins.
 * Threads come from the thread store; Previous / Next walk threads and
 * notes together like Excel, and Delete falls back to the note at the
 * selection when the cell has no thread.
 */
import { t } from './i18n/locale'
import type { RibbonCommandContext } from './ribbon-actions'
import { neighbourThread, threadPane, threadStore, type CellThread } from './threaded-comments'
import { revealCellBelowFreeze } from './univer-sync'

interface Position {
  readonly row: number
  readonly column: number
}

export function handleThreadedCommentCommand(
  ctx: RibbonCommandContext,
  command: string,
  fallback: (command: string) => void,
): void {
  const runtime = ctx.univerRef.current
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  const active = workbook?.getActiveRange()
  if (!runtime || !workbook || !sheet || !active) return
  const sheetId = sheet.getSheetId()
  const row = active.getRow()
  const column = active.getColumn()
  const thread = threadStore.get(sheetId, row, column)
  const notes = sheet.getNotes() as readonly { row: number; col: number; show?: boolean }[]
  const noteAt = (position: Position) =>
    notes.find((note) => note.row === position.row && note.col === position.column)

  switch (command) {
    case 'comment-new':
      if (thread) {
        threadPane.openAt(sheetId, row, column, false)
      } else if (noteAt({ row, column })) {
        ctx.setMessage(t('appThreadCellHasNote'))
      } else {
        threadPane.openAt(sheetId, row, column, true)
      }
      return
    case 'comment-reply':
      if (!thread) {
        ctx.setMessage(t('appThreadNoneAtSelection'))
        return
      }
      threadPane.openAt(sheetId, row, column, false)
      return
    case 'comment-resolve':
      if (!thread) {
        ctx.setMessage(t('appThreadNoneAtSelection'))
        return
      }
      threadStore.setDone(sheetId, row, column, !thread.done)
      return
    case 'comment-delete':
      if (thread) threadStore.remove(sheetId, row, column)
      else fallback('note-delete')
      return
    case 'comments-pane-toggle':
      threadPane.toggle()
      return
    case 'comment-prev':
    case 'comment-next': {
      const threads = threadStore.sheetThreads(sheetId)
      const stops: CellThread[] = [
        ...threads,
        ...notes
          .filter((note) => !threadStore.get(sheetId, note.row, note.col))
          .map((note) => ({ row: note.row, column: note.col }) as CellThread),
      ]
      const target = neighbourThread(
        stops,
        row,
        column,
        command === 'comment-next' ? 'next' : 'prev',
      )
      if (!target) {
        ctx.setMessage(t('appNoNotesOnSheet'))
        return
      }
      sheet.getRange(target.row, target.column, 1, 1).activate()
      void revealCellBelowFreeze(sheet, target.row, target.column)
      if (threadStore.get(sheetId, target.row, target.column)) {
        threadPane.openAt(sheetId, target.row, target.column, false)
      } else {
        void runtime.univerAPI.executeCommand('sheet.operation.add-note-popup')
      }
      return
    }
    case 'note-show-all': {
      // Toggle: pin every hidden note; when all are pinned, unpin them all.
      const hidden = notes.filter((note) => !note.show)
      const targets = hidden.length > 0 ? hidden : notes
      if (targets.length === 0) {
        ctx.setMessage(t('appNoNotesOnSheet'))
        return
      }
      const unitId = workbook.getId()
      for (const note of targets) {
        void runtime.univerAPI.executeCommand('sheet.mutation.toggle-note-popup', {
          unitId,
          sheetId,
          row: note.row,
          col: note.col,
        })
      }
      return
    }
    case 'notes-convert': {
      const threads = threadStore.sheetThreads(sheetId)
      if (threads.length === 0) {
        ctx.setMessage(t('appNoCommentsOnSheet'))
        return
      }
      for (const entry of threads) {
        threadStore.convertToNote(sheetId, entry.row, entry.column)
        const body = [entry.text, ...entry.replies.map((reply) => reply.text)].join('\n\n')
        // Same "Author:\n" marker as encodeNoteText; the save splits it back.
        sheet.getRange(entry.row, entry.column).createOrUpdateNote({
          id: `note-${sheetId}-${entry.row}-${entry.column}`,
          row: entry.row,
          col: entry.column,
          width: 220,
          height: 90,
          note: `${entry.author}:\n${body}`,
        })
      }
      ctx.setMessage(t('appThreadConverted', { count: threads.length }))
      return
    }
    default:
      return
  }
}
