/**
 * The host card that rides under Univer's note popup: Univer draws a note as
 * a yellow text-only popup (see note-pictures.ts), and a note that carries a
 * picture shows it here instead of nowhere. The trigger is Univer's own note
 * popup service, so the card lives and dies with the popup — hover-temp and
 * pinned popups surface it (Show All draws notes as static annotations, not
 * popups, so it does not), and the card never fights the popup for the same
 * pixels (it hangs below the popup's box).
 */
import { SheetsNotePopupService } from '@univerjs/sheets-note-ui'

import type { NotePicture } from './note-pictures'
import type { UniverRuntime } from './univer-state'

/** The note popup that is up and carries a picture. */
export interface NotePopupCard {
  readonly sheetId: string
  readonly row: number
  readonly column: number
  readonly picture: NotePicture
}

export function installNotePopupCard(
  runtime: UniverRuntime,
  options: {
    /** The note's picture, if the file or this session's journal has one. */
    readonly picturesAt: (sheetId: string, row: number, column: number) => NotePicture | undefined
    readonly onCard: (card: NotePopupCard | null) => void
  },
): { dispose(): void } {
  const service = runtime.univer.__getInjector().get(SheetsNotePopupService)
  const subscription = service.activePopup$.subscribe((popup) => {
    if (!popup) {
      options.onCard(null)
      return
    }
    const picture = options.picturesAt(popup.subUnitId, popup.row, popup.col)
    options.onCard(
      picture ? { sheetId: popup.subUnitId, row: popup.row, column: popup.col, picture } : null,
    )
  })
  return {
    dispose() {
      subscription.unsubscribe()
    },
  }
}
