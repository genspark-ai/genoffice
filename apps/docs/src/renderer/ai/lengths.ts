const EMU_PER: Record<string, number> = {
  emu: 1,
  twip: 635,
  twips: 635,
  pt: 12700,
  px: 9525,
  in: 914400,
  cm: 360000,
  mm: 36000,
}

const LENGTH = /^\s*(-?\d+(?:\.\d+)?)\s*(emu|twips?|pt|px|in|cm|mm)\s*$/i

/**
 * Largest magnitude a parsed length may have (~35in / ~89cm in EMU).
 * AI-supplied lengths are untrusted: uncapped, "99999999in" becomes ~1e14
 * twips and breaks layout + save round-trip.
 */
export const MAX_LENGTH_EMU = 32_000_000

function boundEmu(emu: number): number | undefined {
  if (!Number.isFinite(emu)) return undefined
  const rounded = Math.round(emu)
  return Math.abs(rounded) <= MAX_LENGTH_EMU ? rounded : undefined
}

/** "2.54cm" / "1in" / "72pt" / "96px" / a bare number in `bareUnit` -> EMU; undefined when unparseable */
export function parseEmu(
  value: unknown,
  bareUnit: keyof typeof EMU_PER = 'emu',
): number | undefined {
  if (typeof value === 'number')
    return Number.isFinite(value) ? boundEmu(value * EMU_PER[bareUnit]!) : undefined
  if (typeof value !== 'string') return undefined
  const m = LENGTH.exec(value)
  if (!m) return undefined
  return boundEmu(Number(m[1]) * EMU_PER[m[2]!.toLowerCase()]!)
}

/** length -> twips; bare numbers are twips */
export function parseTwips(value: unknown): number | undefined {
  const emu = parseEmu(value, 'twip')
  return emu === undefined ? undefined : Math.round(emu / 635)
}

/** length -> points; bare numbers are points */
export function parsePoints(value: unknown): number | undefined {
  const emu = parseEmu(value, 'pt')
  return emu === undefined ? undefined : emu / 12700
}

export const emuToPx = (emu: number): number => Math.round(emu / 9525)
