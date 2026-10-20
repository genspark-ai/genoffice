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
  readonly imageId: string
}

/// Every note in the file with a picture, in sheet then cell order. A note
/// without one is not in here at all, which is what keeps the pane from
/// becoming a second list of all comments.
export function collectNotePictures(file: WorkbookFile | null): NotePicture[] {
  if (!file) return []
  const pictures: NotePicture[] = []
  for (const sheet of file.sheets ?? []) {
    for (const comment of sheet.comments ?? []) {
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
  return pictures
}
