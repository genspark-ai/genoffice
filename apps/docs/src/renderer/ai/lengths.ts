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

/** Word's largest "Exactly" line height, 1584pt */
export const MAX_EXACT_LINE_TWIPS = 31680

/** exact line height: a length with a unit ("30pt", "0.5in") -> positive twips; bare numbers are multiples, not lengths */
export function parseExactLineTwips(value: unknown): number | undefined {
  if (typeof value !== 'string') return undefined
  const twips = parseTwips(value)
  return twips !== undefined && twips > 0 && twips <= MAX_EXACT_LINE_TWIPS ? twips : undefined
}

export const emuToPx = (emu: number): number => Math.round(emu / 9525)
