/**
 * Slide size presets and units for Design › Slide Size (genoffice#2191).
 * Sizes are EMU, the unit of `<p:sldSz>`; the paper presets are the real sheet
 * sizes so a deck prints edge to edge on that paper.
 */

export const EMU_PER_UNIT = { mm: 36000, cm: 360000, in: 914400 } as const
export type SizeUnit = keyof typeof EMU_PER_UNIT

/** ECMA-376 ST_SlideSizeCoordinate: 1 inch to 56 inches per side */
export const MIN_SLIDE_EMU = 914400
export const MAX_SLIDE_EMU = 51206400

export type SlideSizePreset = '16:9' | '4:3' | 'A4' | 'A3'

/** landscape orientation of each preset */
export const SLIDE_SIZE_PRESETS: Record<SlideSizePreset, { cx: number; cy: number }> = {
  '16:9': { cx: 12192000, cy: 6858000 },
  '4:3': { cx: 9144000, cy: 6858000 },
  A4: { cx: 297 * 36000, cy: 210 * 36000 },
  A3: { cx: 420 * 36000, cy: 297 * 36000 },
}

/** decks saved elsewhere round sizes slightly differently; half a millimetre still counts */
const MATCH_TOLERANCE_EMU = 18000

export const isPortrait = (size: { cx: number; cy: number }): boolean => size.cy > size.cx

/** the preset a deck size is, in either orientation, or null for a custom size */
export function matchSlideSizePreset(size: { cx: number; cy: number }): SlideSizePreset | null {
  const long = Math.max(size.cx, size.cy)
  const short = Math.min(size.cx, size.cy)
  for (const [key, p] of Object.entries(SLIDE_SIZE_PRESETS) as [
    SlideSizePreset,
    { cx: number; cy: number },
  ][]) {
    if (
      Math.abs(p.cx - long) <= MATCH_TOLERANCE_EMU &&
      Math.abs(p.cy - short) <= MATCH_TOLERANCE_EMU
    )
      return key
  }
  return null
}

export function presetSize(key: SlideSizePreset, portrait: boolean): { cx: number; cy: number } {
  const p = SLIDE_SIZE_PRESETS[key]
  return portrait ? { cx: p.cy, cy: p.cx } : { ...p }
}

/** a typed length in `unit` as EMU, or null when it is not a number or outside the slide limits */
export function lengthToEmu(value: string, unit: SizeUnit): number | null {
  const n = Number(value.trim().replace(',', '.'))
  if (!value.trim() || !Number.isFinite(n)) return null
  const emu = Math.round(n * EMU_PER_UNIT[unit])
  return emu >= MIN_SLIDE_EMU && emu <= MAX_SLIDE_EMU ? emu : null
}

/** EMU shown in `unit`, without trailing zeros (2 decimals; 1 for mm) */
export function emuToLength(emu: number, unit: SizeUnit): string {
  const digits = unit === 'mm' ? 1 : 2
  return String(Number((emu / EMU_PER_UNIT[unit]).toFixed(digits)))
}

const DEFAULT_KEY = 'ai-slides-default-size'

const inLimits = (n: unknown): n is number =>
  typeof n === 'number' && Number.isInteger(n) && n >= MIN_SLIDE_EMU && n <= MAX_SLIDE_EMU

/** "Use for New Presentations": the size new blank decks start at; null keeps the 16:9 template. Per install, like the default shape. */
export function defaultNewSlideSize(): { cx: number; cy: number } | null {
  try {
    const v = JSON.parse(localStorage.getItem(DEFAULT_KEY) ?? 'null') as {
      cx?: unknown
      cy?: unknown
    } | null
    if (v && inLimits(v.cx) && inLimits(v.cy)) return { cx: v.cx, cy: v.cy }
  } catch {
    /* unreadable entry: new decks keep the template size */
  }
  return null
}

export function saveDefaultNewSlideSize(size: { cx: number; cy: number }): void {
  try {
    localStorage.setItem(DEFAULT_KEY, JSON.stringify({ cx: size.cx, cy: size.cy }))
  } catch {
    /* quota: the default is a convenience, not document data */
  }
}
