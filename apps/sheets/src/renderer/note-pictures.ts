/**
 * Notes that carry a picture, listed apart from the notes themselves.
 *
 * Univer draws a note as a yellow popup whose body is the note's text, and
 * that popup is not something a host can put an image into. Threaded comments
 * get away with a hover card because a thread is not a note to Univer at all;
 * a note is. So the picture is offered here instead: a pane listing every
 * note in the workbook that has one, which is also the only view that stays
 * useful when the notes themselves are dozens of cells apart.
 */
import { useSyncExternalStore } from 'react'

import type { WorkbookFile } from '../shared/desktop-api'

type Listener = () => void

interface PaneState {
  readonly open: boolean
}

class NotePicturesPaneController {
  private state: PaneState = { open: false }
  private listeners = new Set<Listener>()

  subscribe = (listener: Listener): (() => void) => {
    this.listeners.add(listener)
    return () => void this.listeners.delete(listener)
  }

  getState = (): PaneState => this.state

  toggle(): void {
    this.set({ open: !this.state.open })
  }

  close(): void {
    this.set({ open: false })
  }

  private set(state: PaneState): void {
    this.state = state
    for (const listener of this.listeners) listener()
  }
}

export const notePicturesPane = new NotePicturesPaneController()

export function useNotePicturesPaneOpen(): boolean {
  return useSyncExternalStore(notePicturesPane.subscribe, notePicturesPane.getState).open
}

/** One note that is showing a picture. */
export interface NotePicture {
  readonly sheetId: string
  readonly sheetName: string
  readonly row: number
  readonly column: number
  readonly author: string
  readonly text: string
  /** The picture's id in the file; absent for a picture not written yet. */
  readonly imageId?: string
  /**
   * The picture's bytes as a data URL, for a picture put into a note this
   * session and not written to the file yet.
   */
  readonly dataUrl?: string
}

/**
 * A picture put into a note this session, before the save writes it.
 *
 * `dataUrl: null` is a picture taken away: the note keeps its text and the
 * pane stops showing the picture the file had.
 */
export interface PendingNotePicture {
  readonly sheetId: string
  readonly row: number
  readonly column: number
  readonly dataUrl: string | null
}

/// Every note in the file with a picture, in sheet then cell order. A note
/// without one is not in here at all, which is what keeps the pane from
/// becoming a second list of all comments.
///
/// A picture put into a note this session replaces the one the file had for
/// that cell: it is the same note, and the pane shows what the save will
/// write.
export function collectNotePictures(
  file: WorkbookFile | null,
  pending: readonly PendingNotePicture[] = [],
): NotePicture[] {
  const pictures: NotePicture[] = []
  const byCell = new Map<string, PendingNotePicture>()
  for (const entry of pending) {
    byCell.set(`${entry.sheetId}:${entry.row}:${entry.column}`, entry)
  }
  const covered = new Set<string>()
  for (const sheet of file?.sheets ?? []) {
    for (const comment of sheet.comments ?? []) {
      const key = `${sheet.id}:${comment.row}:${comment.column}`
      const entry = byCell.get(key)
      // a picture taken away this session is not shown any more
      if (entry?.dataUrl === null) {
        covered.add(key)
        continue
      }
      if (entry) {
        covered.add(key)
        pictures.push({
          sheetId: sheet.id,
          sheetName: sheet.name,
          row: comment.row,
          column: comment.column,
          author: comment.author,
          text: comment.text,
          dataUrl: entry.dataUrl ?? undefined,
        })
        continue
      }
      if (!comment.image) continue
      pictures.push({
        sheetId: sheet.id,
        sheetName: sheet.name,
        row: comment.row,
        column: comment.column,
        author: comment.author,
        text: comment.text,
        imageId: comment.image.id,
      })
    }
  }
  // the notes that had no picture in the file but have one now
  for (const entry of pending) {
    if (!entry.dataUrl) continue
    const key = `${entry.sheetId}:${entry.row}:${entry.column}`
    if (covered.has(key)) continue
    const sheet = file?.sheets?.find((candidate) => candidate.id === entry.sheetId)
    pictures.push({
      sheetId: entry.sheetId,
      sheetName: sheet?.name ?? entry.sheetId,
      row: entry.row,
      column: entry.column,
      author: '',
      text: '',
      dataUrl: entry.dataUrl,
    })
  }
  return pictures
}
