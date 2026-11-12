/**
 * Which face each rFonts slot lays out with under an East Asian docDefaults
 * language (INVENTORY-fonts S1-S4).
 */
import { isEastAsianLangTag } from '@genoffice/docx-engine'
import { isEastAsianFaceName } from '../face-info'
import { resolveDeclaredFace } from '../../font-resolve'
import { openDocumentText } from '../open-document'

// docDefaults w:lang w:val naming an East Asian language (Google Docs exports
// of Korean documents write lang=ko with Arial in every slot) makes Word lay
// out every run's Latin text with the run's eastAsia face, run-level lang
// overrides included; an eastAsia face without East Asian coverage (Arial,
// Courier New, Times New Roman, Calibri) is replaced by SimSun for Latin
// (0.5em half-width, 1.3 line) and Batang for hangul (probe 2026-11-02; a
// run-level lang=ko without the docDefaults one changes nothing).
export function latinViaEastAsiaSlotFor(parsed: { docDefaults?: { lang?: string } }): boolean {
  return isEastAsianLangTag(parsed.docDefaults?.lang)
}
export function latinViaEastAsiaSlotOn(): boolean {
  return openDocumentText().latinViaEaSlot
}
/** the document's inherited East Asian face: the eastAsia slot of runs that declare none */
function docEaSlotFace(): string | null {
  return openDocumentText().bodyEaSlot
}

/** faces a run's slots lay out with: its own ascii / eastAsia pair, or under
 *  the East Asian docDefaults language the eastAsia face (SimSun when that
 *  face draws no East Asian script) for both */
export function wordSlotFaces<T extends string | null | undefined>(
  ascii: T,
  eastAsia: T,
  on: boolean = latinViaEastAsiaSlotOn(),
): { ascii: T | string; eastAsia: T | string } {
  if (!on) return { ascii, eastAsia }
  const face = eastAsia || docEaSlotFace() || ascii
  if (!face) return { ascii, eastAsia }
  const laid = eastAsianCoverage(face) ? face : 'SimSun'
  return { ascii: laid, eastAsia: laid }
}

/** the face draws East Asian script itself: a CJK / Korean name (either name
 *  table), or a face Word ships whose OS/2 code pages say so; the same bit on
 *  installed Latin macOS faces (Arial, Courier New) is a data quirk */
function eastAsianCoverage(face: string): boolean {
  if (isEastAsianFaceName(face)) return true
  const r = resolveDeclaredFace(face)
  if (!r || r.substituted || !r.metrics.eastAsian) return false
  return r.source !== 'installed' && r.source !== 'system'
}

/** face a paragraph mark lays its line with under the rule: the eastAsia docDefaults
 *  face when it draws East Asian script (Calibri/Malgun 10pt empty and space-only
 *  lines 17.28 under w:val=ko-KR, probe 2026-11-03); a Latin eastAsia face leaves the
 *  mark at its ascii height (Arial everywhere, probe 2026-11-02) */
export function inheritedMarkFace<T extends string | null | undefined>(ascii: T): T | string {
  const slot = docEaSlotFace()
  if (!latinViaEastAsiaSlotOn() || !slot || !eastAsianCoverage(slot)) return ascii
  return slot
}

/** face a run without any rFonts lays its Latin out with under the rule, else null */
export function inheritedSlotFace(): string | null {
  const slot = docEaSlotFace()
  if (!latinViaEastAsiaSlotOn() || !slot) return null
  return eastAsianCoverage(slot) ? slot : 'SimSun'
}

/** wordSlotFaces of a run's attrs: font is eastAsia-first and copies the Latin
 *  name when the run declares no eastAsia (a parsed run then carries
 *  eastAsiaFont), so only a declared eastAsia counts as the slot */
export function runSlotFaces(attrs: {
  font?: string | null
  fontAscii?: string | null
  eastAsiaFont?: string | null
  rawRPr?: string | null
}): { ascii: string | null; eastAsia: string | null; declaredEa: string | null } {
  const font = attrs.font ?? null
  const ascii = attrs.fontAscii ?? null
  const declaredEa = attrs.eastAsiaFont ?? (!attrs.rawRPr && font !== ascii ? font : null)
  if (!latinViaEastAsiaSlotOn()) return { ascii, eastAsia: font, declaredEa }
  const slot = wordSlotFaces(ascii, declaredEa)
  return { ascii: slot.ascii, eastAsia: slot.eastAsia ?? slot.ascii, declaredEa }
}
