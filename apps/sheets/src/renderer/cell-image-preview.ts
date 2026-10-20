/**
 * Double-click preview for pictures embedded in cells. Univer's own
 * double-click starts editing the host cell; when the pointer lands on a
 * cell whose content is a picture, the viewer opens instead and the edit
 * is swallowed.
 */
import { isGridKeyTarget } from './clear-selection-keyboard'
import type { PictureAnchor } from './picture-paste'

export interface CellImagePreviewHost {
  /// True while a cell editor is open; double-clicks must keep editing.
  isCellEditing(): boolean
  /// The grid cell under a viewport point, or null off the sheet canvas.
  cellAtPoint(clientX: number, clientY: number): PictureAnchor | null
  /// Opens the viewer when the cell holds a picture; false passes the
  /// double-click through to Univer.
  openPreview(row: number, column: number): boolean
}

export function installCellImagePreview(
  gridHost: HTMLElement,
  host: CellImagePreviewHost,
): () => void {
  const onDoubleClick = (event: MouseEvent): void => {
    if (event.defaultPrevented || event.button !== 0) return
    if (host.isCellEditing() || !isGridKeyTarget(event.target)) return
    const cell = host.cellAtPoint(event.clientX, event.clientY)
    if (!cell || !host.openPreview(cell.row, cell.column)) return
    // The viewer took the gesture; Univer must not start editing the cell.
    event.preventDefault()
    event.stopImmediatePropagation()
  }
  gridHost.addEventListener('dblclick', onDoubleClick, true)
  return () => gridHost.removeEventListener('dblclick', onDoubleClick, true)
}
