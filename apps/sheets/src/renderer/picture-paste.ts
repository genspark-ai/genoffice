/**
 * Clipboard / drag-and-drop bitmaps become floating pictures anchored at a
 * cell, the way Excel handles a pasted screenshot. Univer's own paste path
 * (hidden editor -> SheetPasteShortKeyCommand -> drawing-ui float image)
 * would otherwise insert a drawing this app never journals or saves, so the
 * window-capture listener claims image-only transfers before it runs.
 */
import { isGridKeyTarget } from './clear-selection-keyboard'

export const PICTURE_MAX_BYTES = 20 * 1024 * 1024
const PICTURE_MEDIA_TYPES = new Set(['image/png', 'image/jpeg', 'image/gif'])

export interface PictureAnchor {
  readonly row: number
  readonly column: number
}

export interface TransferItemLike {
  readonly kind: string
  readonly type: string
  getAsFile(): File | null
}

export interface TransferLike {
  readonly items: ArrayLike<TransferItemLike>
  getData(format: string): string
}

/// Whether the transfer carries an image file at all (drag-over exposes item
/// types but not the files themselves).
export function transferHasImage(data: Pick<TransferLike, 'items'>): boolean {
  return Array.from(data.items).some(
    (item) => item.kind === 'file' && item.type.startsWith('image/'),
  )
}

/// The image to insert when the transfer is a picture rather than cell data:
/// Excel/Numbers copy a bitmap of the cells next to their table markup, and
/// text without an <img> counterpart is still text, so both stay cell pastes.
export function pickTransferImage(data: TransferLike): File | null {
  if (!transferHasImage(data)) return null
  const html = data.getData('text/html')
  if (/<(table|tr|td|th)\b/i.test(html)) return null
  if (data.getData('text/plain').trim() && !/<img\b/i.test(html)) return null
  for (const item of Array.from(data.items)) {
    if (item.kind !== 'file' || !item.type.startsWith('image/')) continue
    const file = item.getAsFile()
    if (file) return file
  }
  return null
}

export function pictureMediaType(fileType: string): string {
  return fileType === 'image/jpg' ? 'image/jpeg' : fileType
}

export function pictureFileProblem(
  file: Pick<File, 'size' | 'type'>,
): 'too-large' | 'bad-type' | null {
  if (file.size > PICTURE_MAX_BYTES) return 'too-large'
  if (!PICTURE_MEDIA_TYPES.has(pictureMediaType(file.type))) return 'bad-type'
  return null
}

/// Frame size in grid cells for a picture of the given natural pixel size:
/// ~80px per column, ~22px per row, scaled down to a <=480px-wide frame.
export function fitPictureFrame(
  naturalWidth: number,
  naturalHeight: number,
): { columns: number; rows: number } {
  const scale = Math.min(1, 480 / Math.max(1, naturalWidth))
  return {
    columns: Math.min(16, Math.max(2, Math.round((naturalWidth * scale) / 80))),
    rows: Math.min(40, Math.max(2, Math.round((naturalHeight * scale) / 22))),
  }
}

export function pointInRect(
  x: number,
  y: number,
  rect: { left: number; top: number; right: number; bottom: number },
): boolean {
  return x >= rect.left && x < rect.right && y >= rect.top && y < rect.bottom
}

export interface PictureTransferHost {
  isCellEditing(): boolean
  cellAtPoint(clientX: number, clientY: number): PictureAnchor | null
  insert(file: File, anchor: PictureAnchor | null): void
}

export function installPictureTransfer(
  gridHost: HTMLElement,
  host: PictureTransferHost,
): () => void {
  const onPaste = (event: ClipboardEvent): void => {
    if (event.defaultPrevented || !event.clipboardData) return
    if (host.isCellEditing() || !isGridKeyTarget(event.target)) return
    const file = pickTransferImage(event.clipboardData)
    if (!file) return
    event.preventDefault()
    event.stopImmediatePropagation()
    host.insert(file, null)
  }
  const onDragOver = (event: DragEvent): void => {
    if (host.isCellEditing()) return
    if (!event.dataTransfer || !transferHasImage(event.dataTransfer)) return
    event.preventDefault()
    event.dataTransfer.dropEffect = 'copy'
  }
  const onDrop = (event: DragEvent): void => {
    if (host.isCellEditing()) return
    const file = event.dataTransfer ? pickTransferImage(event.dataTransfer) : null
    if (!file) return
    event.preventDefault()
    event.stopPropagation()
    host.insert(file, host.cellAtPoint(event.clientX, event.clientY))
  }
  window.addEventListener('paste', onPaste, true)
  gridHost.addEventListener('dragover', onDragOver, true)
  gridHost.addEventListener('drop', onDrop, true)
  return () => {
    window.removeEventListener('paste', onPaste, true)
    gridHost.removeEventListener('dragover', onDragOver, true)
    gridHost.removeEventListener('drop', onDrop, true)
  }
}
