/**
 * The Latin face a CJK-named declared face lays its Latin out in, and the
 * autospace gap Word takes from the Latin face's xAvgCharWidth
 * (INVENTORY-fonts G1-G3).
 */
import { isEastAsianFaceName as isCjkNamedFont } from '../face-info'
import { resolveDeclaredFace } from '../../font-resolve'
import { isBundledFont, isFontAvailable } from './availability'
import { latinSlotTarget } from './css-chains'
import { latinFace, legacyFaceAlias, lineHeightFactor } from './line-factor'

/** OS/2 xAvgCharWidth (em) of a declared Latin name's Word layout face; null for East Asian or unknown faces */
function latinAvgCharWidth(font: string): number | null {
  const face = latinFace(font)
  return face && !face.metrics.eastAsian ? face.metrics.xAvg : null
}
/** a Latin name Word lacks lays out in Cambria (fallback: the 1/4 em gap, doubled) */
const cambriaAvgCharWidth = () => latinAvgCharWidth('Cambria') ?? 0.5

const latinGapCache = new Map<string, number>()
/**
 * Word probe 2026-10-20: the autospace gap between a Latin letter/digit and
 * the CJK letter after it is half the Latin face's OS/2 xAvgCharWidth
 * (Cambria 0.308, Calibri 0.260, Aptos 0.280, Arial 0.221, Times New Roman
 * 0.200, Garamond 0.194 em) whatever the EA face or size; the CJK -> Latin
 * gap and CJK-capable ascii faces (Batang, Malgun, Yu Gothic) stay 1/4 em.
 * A Latin name Word lacks lays out in Cambria, as does the Latin of a missing
 * KR superfamily name (the 'KR Theme Latin GO' alias).
 */
export function latinAutospaceGapEm(font: string | null | undefined): number {
  if (!font) return 0.25
  const cached = latinGapCache.get(font)
  if (cached !== undefined) return cached
  // a hangul-named face resolves to its EA face; its Latin slot is the KR theme rule below
  const avg = isCjkNamedFont(font) && !legacyFaceAlias(font) ? null : latinAvgCharWidth(font)
  let gap: number
  if (avg !== null) gap = avg / 2
  else if (isCjkNamedFont(font)) {
    const latin = latinSlotFace(font)
    gap = latin ? (latinAvgCharWidth(latin) ?? 0.5) / 2 : 0.25
  } else if (!isBundledFont(font) && !isFontAvailable(font)) gap = cambriaAvgCharWidth() / 2
  else gap = 0.25
  latinGapCache.set(font, gap)
  return gap
}

/**
 * Latin face a CJK-named face's Latin lays out in: the split Latin slot's
 * target, or the Latin face Word substitutes wholesale (altName Arial); the
 * same face the CSS chain leads with. Null when the declared face draws its
 * own Latin or the slot lands on an East Asian face (SimSun), whose box the
 * EA rule owns. Any CJK-lettered name qualifies, not only the known families:
 * the resolver decides whether it was substituted and onto what
 */
export function latinSlotFace(font: string): string | null {
  if (!isCjkNamedFont(font)) return null
  const target = latinSlotTarget(font)
  if (target) {
    const face = latinFace(target)
    return face && !face.metrics.eastAsian ? target : null
  }
  const face = resolveDeclaredFace(font)
  return face?.substituted && !face.metrics.eastAsian ? face.family : null
}

/** Latin-only lines of a missing KR superfamily name lay out in Word's Cambria substitute (1.172,
 *  prod corpus 20261020 sample 023 URL lines), not in the hangul face's 1.3029 */
export function latinLineFactor(font: string): number {
  const latin = latinSlotFace(font)
  return latin ? lineHeightFactor(latin) : lineHeightFactor(font)
}
