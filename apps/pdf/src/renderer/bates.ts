import type { TextInsertInput } from '../shared/ipc'
import { viewToPdf } from './annotations'

/**
 * Bates numbering as plain text inserts: one real, searchable text object per
 * page through the existing insert pipeline, so the numbers survive every save
 * like any other content. Geometry works in display space (the page as the
 * reader sees it, after its /Rotate) and maps back to user space, because the
 * insert matrix only counter-rotates glyphs — the origin itself stays in the
 * page's unrotated coordinate system.
 */

export type BatesPosition =
  'bottom-left' | 'bottom-center' | 'bottom-right' | 'top-left' | 'top-center' | 'top-right'

export interface BatesConfig {
  prefix: string
  start: number
  digits: number
  fontSize: number
  color: [number, number, number]
  position: BatesPosition
  /** 1-based first page to stamp */
  firstPage: number
  /** 1-based last page (inclusive); null runs to the end */
  lastPage: number | null
}

export const BATES_MARGIN = 28
export const BATES_MIN_FONT = 6
export const BATES_MAX_FONT = 36

/** `ACME-` + 41 at 6 digits → `ACME-000041`; a number that outgrows its field uses every digit it needs */
export function formatBates(prefix: string, n: number, digits: number): string {
  return `${prefix}${String(n).padStart(digits, '0')}`
}

export type BatesError = 'batesErrStart' | 'batesErrDigits' | 'batesErrFontSize' | 'batesErrRange'

export function validateBates(cfg: Omit<BatesConfig, 'color' | 'position'>): BatesError | null {
  if (!Number.isInteger(cfg.start) || cfg.start < 0) return 'batesErrStart'
  if (!Number.isInteger(cfg.digits) || cfg.digits < 1 || cfg.digits > 15) return 'batesErrDigits'
  if (cfg.fontSize < BATES_MIN_FONT || cfg.fontSize > BATES_MAX_FONT) return 'batesErrFontSize'
  if (!Number.isInteger(cfg.firstPage) || cfg.firstPage < 1) return 'batesErrRange'
  if (cfg.lastPage !== null && (!Number.isInteger(cfg.lastPage) || cfg.lastPage < cfg.firstPage)) {
    return 'batesErrRange'
  }
  return null
}

/**
 * One insert per stamped page. `measure` is the renderer's canvas text width —
 * center/right alignment travels as `lineXOffsets`, the same convention the
 * interactive insert flow uses, because the writer has no glyph metrics.
 */
export function buildBatesInserts(
  pages: readonly { width: number; height: number; rotate: number; x0?: number; y0?: number }[],
  cfg: BatesConfig,
  measure: (text: string, fontSize: number) => number,
): TextInsertInput[] {
  const out: TextInsertInput[] = []
  const last = cfg.lastPage ?? pages.length
  for (let index = cfg.firstPage - 1; index < Math.min(last, pages.length); index++) {
    // display corners are crop-relative: the user-space origin lands on the cropped page
    const { width, height, rotate, x0 = 0, y0 = 0 } = pages[index]!
    const landscape = (((rotate % 360) + 360) % 360) % 180 === 90
    const displayWidth = landscape ? height : width
    const displayHeight = landscape ? width : height
    const text = formatBates(cfg.prefix, cfg.start + (index - (cfg.firstPage - 1)), cfg.digits)
    const align = cfg.position.endsWith('center')
      ? 'center'
      : cfg.position.endsWith('right')
        ? 'right'
        : 'left'
    const offset =
      align === 'left'
        ? 0
        : align === 'center'
          ? -measure(text, cfg.fontSize) / 2
          : -measure(text, cfg.fontSize)
    const xDisplay =
      align === 'left'
        ? BATES_MARGIN
        : align === 'center'
          ? displayWidth / 2
          : displayWidth - BATES_MARGIN
    const yDisplay = cfg.position.startsWith('top')
      ? BATES_MARGIN + cfg.fontSize
      : displayHeight - BATES_MARGIN
    // The one display→user transform the app owns: a hand-rolled copy agreed
    // with it only for 0°, and viewToPdf also carries the CropBox origin and
    // /UserUnit.
    const [ux, uy] = viewToPdf({ pw: width, ph: height, rot: rotate, x0, y0 }, xDisplay, yDisplay)
    out.push({
      pageIndex: index,
      origin: [ux, uy],
      text,
      fontSize: cfg.fontSize,
      color: cfg.color,
      lineXOffsets: [offset],
      align,
      rotate,
    })
  }
  return out
}
