/**
 * ODF length parsing. All conversions hub on EMU (914400 per inch), the one
 * unit shared exactly by both target formats: docx twips are EMU/635 and pptx
 * offsets are EMU directly, so a centimetre never drifts through two float
 * paths.
 */

export const EMU_PER_INCH = 914_400
export const EMU_PER_PT = 12_700
/** docx lengths are twentieths of a point */
export const TWIPS_PER_INCH = 1440

const EMU_PER_UNIT: Record<string, number> = {
  cm: 360_000,
  mm: 36_000,
  in: EMU_PER_INCH,
  pt: EMU_PER_PT,
  pc: 12 * EMU_PER_PT,
  px: Math.round(EMU_PER_INCH / 96),
}

/**
 * `"2.5cm"`, `"0.5in"`, `"10mm"`, `"12pt"`, `"2pc"` → EMU. Percentages and
 * bare non-zero numbers have no provable unit and return null; a bare `0` is
 * accepted (producers write it for a zero length).
 */
export function odfLengthToEmu(text: string | undefined): number | null {
  if (!text) return null
  const match = /^(-?\d*\.?\d+)(cm|mm|in|pt|pc|px)?$/.exec(text.trim())
  if (!match) return null
  const value = Number(match[1])
  if (!Number.isFinite(value)) return null
  const unit = match[2]
  if (!unit) return value === 0 ? 0 : null
  return Math.round(value * EMU_PER_UNIT[unit])
}

export const emuToPt = (emu: number): number => emu / EMU_PER_PT
export const emuToTwips = (emu: number): number => Math.round(emu / 635)

/** `#RRGGBB` / `#RGB` (ODF colors) → uppercase hex without `#`; null otherwise. */
export function odfColorToHex(text: string | undefined): string | null {
  if (!text) return null
  const hex = text.trim().replace(/^#/, '').toUpperCase()
  if (/^[0-9A-F]{6}$/.test(hex)) return hex
  if (/^[0-9A-F]{3}$/.test(hex)) {
    return hex
      .split('')
      .map((c) => c + c)
      .join('')
  }
  return null
}

/** Image mime from magic bytes; falls back to the path extension. */
export function sniffImageMime(
  bytes: Uint8Array,
  path: string,
): 'image/png' | 'image/jpeg' | 'image/gif' | null {
  if (
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47
  ) {
    return 'image/png'
  }
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
    return 'image/jpeg'
  }
  if (bytes.length >= 6 && bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
    return 'image/gif'
  }
  if (/\.png$/i.test(path)) return 'image/png'
  if (/\.(jpe?g)$/i.test(path)) return 'image/jpeg'
  if (/\.gif$/i.test(path)) return 'image/gif'
  return null
}
