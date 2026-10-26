/**
 * Putting a picture into a note.
 *
 * Univer draws a note as a yellow popup whose body is its text, and that
 * popup is not something a host can put an image into. So the picture the
 * reader picks here cannot be shown inside the note while the app is open: it
 * is journalled instead, the save paints it onto the note's shape, and the
 * Note pictures pane shows it in the meantime.
 *
 * The bytes travel with the note they belong to, not as a separate edit: a
 * note carries one picture, and a note that says nothing about its picture
 * keeps the one the file already had.
 */
import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'

import { journalSize, recordNoteChange, recordNotePicture } from './edit-journal'
import { t } from './i18n/locale'
import { pictureFileProblem, pictureMediaType } from './picture-paste'
import type { RibbonCommandContext } from './ribbon-actions'
import { threadStore } from './threaded-comments'
import type { UniverRuntime } from './univer-state'

const MEDIA_TYPES = new Set(['image/png', 'image/jpeg', 'image/gif'])

/// Picks a picture and puts it into the note at the selection.
export function handleInsertNotePicture(ctx: RibbonCommandContext, runtime: UniverRuntime): void {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  const active = workbook?.getActiveRange()
  if (!sheet || !active) return
  const row = active.getRow()
  const column = active.getColumn()
  const sheetId = sheet.getSheetId()
  // a cell holds either a thread or a note (Excel), and a thread has no shape
  // to paint a picture on
  if (threadStore.get(sheetId, row, column)) {
    ctx.setMessage(t('appNoteCellHasThread'))
    return
  }
  const input = document.createElement('input')
  input.type = 'file'
  input.accept = 'image/png,image/jpeg,image/gif'
  input.onchange = () => {
    const file = input.files?.[0]
    if (file) void insertNotePictureFile(ctx, runtime, sheetId, row, column, file)
  }
  input.click()
}

/// Takes the picture off a note, leaving its text alone.
export function handleRemoveNotePicture(
  ctx: RibbonCommandContext,
  sheetId: string,
  row: number,
  column: number,
): void {
  const state = ctx.lazyWorkbookRef.current
  if (!state) return
  recordNotePicture(state.editJournal, { sheetId, row, column, picture: null })
  // the note itself did not change, so nothing else would mark the sheet
  recordNoteChange(state.editJournal, sheetId)
  ctx.setPendingEdits(journalSize(state.editJournal))
  ctx.setMessage(t('appNotePictureRemoved', { cell: formatAddress(row, column) }))
}

async function insertNotePictureFile(
  ctx: RibbonCommandContext,
  runtime: UniverRuntime,
  sheetId: string,
  row: number,
  column: number,
  file: File,
): Promise<void> {
  const problem = pictureFileProblem(file)
  if (problem) {
    ctx.setMessage(t(problem === 'too-large' ? 'appPictureTooLarge' : 'appPictureBadType'))
    return
  }
  const mediaType = pictureMediaType(file.type)
  if (!MEDIA_TYPES.has(mediaType)) {
    ctx.setMessage(t('appPictureBadType'))
    return
  }
  const state = ctx.lazyWorkbookRef.current
  if (!state) {
    ctx.setMessage(t('appPictureNeedsFile'))
    return
  }
  const dataUrl = await readAsDataUrl(file)
  if (!dataUrl) return
  const size = await measureImage(dataUrl)
  // The picture is painted onto the note's shape, so the note has to be there.
  // A note with nothing but a picture is what Excel's own picture-in-a-popup
  // is, so an empty note is made rather than refused.
  await ensureNote(ctx, runtime, sheetId, row, column)
  recordNotePicture(state.editJournal, {
    sheetId,
    row,
    column,
    picture: {
      mediaType: mediaType as 'image/png' | 'image/jpeg' | 'image/gif',
      dataUrl,
      widthPx: size.width,
      heightPx: size.height,
      name: file.name,
    },
  })
  recordNoteChange(state.editJournal, sheetId)
  ctx.setPendingEdits(journalSize(state.editJournal))
  ctx.setMessage(t('appNotePictureInserted', { cell: formatAddress(row, column) }))
}

/// Creates the note at the cell when there is none, so a picture can be the
/// note's whole content.
async function ensureNote(
  ctx: RibbonCommandContext,
  runtime: UniverRuntime,
  sheetId: string,
  row: number,
  column: number,
): Promise<void> {
  const sheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(sheetId)
  if (!sheet) return
  const hasNote = sheet.getNotes().some((note) => note.row === row && note.col === column)
  if (hasNote) return
  await ctx.runOps(
    [{ op: 'set_note', sheetId, address: formatAddress(row, column), text: '' }],
    null,
  )
}

function readAsDataUrl(file: File): Promise<string | null> {
  return new Promise((resolve) => {
    const reader = new FileReader()
    reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : null)
    reader.onerror = () => resolve(null)
    reader.readAsDataURL(file)
  })
}

/// The picture's own size: the note's shape is given it, so a photograph is
/// shown at its own proportions rather than stretched.
function measureImage(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const image = new Image()
    image.onload = () =>
      resolve({ width: image.naturalWidth || 480, height: image.naturalHeight || 320 })
    image.onerror = () => resolve({ width: 480, height: 320 })
    image.src = dataUrl
  })
}
