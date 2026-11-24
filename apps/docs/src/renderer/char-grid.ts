import type { DocDefaults, DocGrid, StyleInfo } from '@genoffice/docx-engine'

/** the grids that give every line a pitch */
export const LINE_GRIDS: ReadonlyArray<DocGrid['type']> = ['lines', 'linesAndChars', 'snapToChars']

/** the grids that give every character a pitch */
export const CHAR_GRIDS: ReadonlyArray<DocGrid['type']> = ['linesAndChars', 'snapToChars']

const CHAR_SPACE_PER_PT = 4096

/** Word probe 2026-10-22: a file that states no size anywhere snaps to cells of 12pt + charSpace */
const UNSTATED_SIZE_HALF_POINTS = 24

/** size (pt) of the default paragraph style: what the character pitch is counted from */
export function gridBaseSizePt(
  styles: ReadonlyMap<string, StyleInfo> | undefined,
  docDefaults: DocDefaults | undefined,
): number {
  let normal: StyleInfo | undefined
  for (const s of styles?.values() ?? [])
    if (s.type === 'paragraph' && s.isDefault) {
      normal = s
      break
    }
  return (
    (normal?.display?.sizeHalfPoints ?? docDefaults?.sizeHalfPoints ?? UNSTATED_SIZE_HALF_POINTS) /
    2
  )
}

/** w:charSpace is the pitch's distance from the base size, in 1/4096 pt */
export const charPitchPt = (basePt: number, charSpace: number | undefined): number =>
  basePt + (charSpace ?? 0) / CHAR_SPACE_PER_PT

/** rounded down, so the count the pitch was made for still fits the line */
export const charSpaceForPitch = (basePt: number, pitchPt: number): number =>
  Math.floor((pitchPt - basePt) * CHAR_SPACE_PER_PT + 1e-6)

export const charsPerLine = (textWidthPt: number, pitchPt: number): number =>
  Math.max(1, Math.floor(textWidthPt / pitchPt + 1e-6))

/**
 * What a paragraph with w:adjustRightInd on adds to its right indent in a
 * document below compatibilityMode 15: the part of the column that a whole
 * number of cells leaves over (Word probes 2026-11-22). Word 2013 documents
 * lay out as if the setting were off.
 */
export function gridRightIndentPt(columnWidthPt: number, pitchPt: number): number {
  if (!(columnWidthPt > 0) || !(pitchPt > 0)) return 0
  const rest = columnWidthPt - Math.floor(columnWidthPt / pitchPt + 1e-6) * pitchPt
  // a hair less, so the last whole cell still fits once widths are rounded to pixels
  return rest >= 0.05 ? (Math.floor(rest * 100) - 2) / 100 : 0
}

/** cells narrower than this share of the base size are not offered */
const MIN_PITCH_RATIO = 10 / 11

export const maxCharsPerLine = (textWidthPt: number, basePt: number): number =>
  charsPerLine(textWidthPt, basePt * MIN_PITCH_RATIO)

/** the w:charSpace the dialog may write for a section this wide */
export function charSpaceInRange(charSpace: number, textWidthPt: number, basePt: number): boolean {
  const pitch = charPitchPt(basePt, charSpace)
  return (
    Number.isInteger(charSpace) && pitch >= basePt * MIN_PITCH_RATIO - 1e-3 && pitch <= textWidthPt
  )
}

/**
 * letter-spacing of a full-width character snapped to cells of `pitchPt`
 * (Word probe 2026-10-22, snapToChars): it advances by whole cells, as many
 * as its own size needs. The value is kept as text so 1em resolves on the
 * run that uses it.
 */
export const charCellSpacingCss = (pitchPt: number): string =>
  `calc(round(up, 1em - 0.01px, ${Math.round(pitchPt * 10000) / 10000}pt) - 1em)`
