/**
 * Placement for the fixed-position folder context menu.
 *
 * A right-click menu opens to the RIGHT of the pointer; a "..." button menu stays
 * right-aligned under its button (align 'end'). Either one flips above the anchor
 * when there is no room below, and both axes are clamped to the viewport with a
 * margin, so the menu is never cut off at any edge.
 */

export interface FolderMenuAnchor {
  /** preferred left edge of the menu (viewport coords) */
  x: number
  /** preferred top edge of the menu (viewport coords) */
  y: number
  /** bottom edge of the anchor; used to flip the menu above when it does not fit below */
  above: number
  /**
   * 'start' (default): `x` is the menu's left edge, so it opens to the right of the pointer.
   * 'end': `x` is the menu's right edge, so a "..." button menu stays right-aligned under its button.
   */
  align?: 'start' | 'end'
}

export interface FolderMenuSize {
  width: number
  height: number
}

export interface FolderMenuViewport {
  width: number
  height: number
}

export interface FolderMenuPosition {
  left: number
  top: number
}

/** gap kept between the viewport edges and the menu */
export const FOLDER_MENU_MARGIN = 8

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), Math.max(min, max))
}

export function placeFolderMenu(
  anchor: FolderMenuAnchor,
  menu: FolderMenuSize,
  viewport: FolderMenuViewport,
  margin: number = FOLDER_MENU_MARGIN,
): FolderMenuPosition {
  const preferredLeft = anchor.align === 'end' ? anchor.x - menu.width : anchor.x
  const left = clamp(preferredLeft, margin, viewport.width - menu.width - margin)
  const fitsBelow = anchor.y + menu.height <= viewport.height - margin
  const preferredTop = fitsBelow ? anchor.y : anchor.above - menu.height
  const top = clamp(preferredTop, margin, viewport.height - menu.height - margin)
  return { left, top }
}
