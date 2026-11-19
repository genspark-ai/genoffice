import type { SectionSettings } from '@genoffice/docx-engine'

export type SideMargins = Pick<SectionSettings, 'marginLeft' | 'marginRight'>

/**
 * Side margins (twips) of the page whose displayed number is `pageNo`. With
 * w:mirrorMargins, w:left (the gutter already folded in by the parser) is the
 * inside margin and w:right the outside one: odd pages read them as left/right,
 * even pages swap them. Parity follows the displayed number like Word's
 * odd/even headers (w:pgNumType start and even/odd section starts included).
 */
export function pageMargins(
  set: SideMargins,
  pageNo: number,
  mirror: boolean,
): { left: number; right: number } {
  return mirror && pageNo % 2 === 0
    ? { left: set.marginRight, right: set.marginLeft }
    : { left: set.marginLeft, right: set.marginRight }
}

/** Horizontal shift (px) of a page's text column from the section's odd-page placement. */
export function mirrorShiftPx(set: SideMargins, pageNo: number, mirror: boolean): number {
  return ((pageMargins(set, pageNo, mirror).left - set.marginLeft) / 1440) * 96
}
