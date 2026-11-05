/**
 * Document-level text layout facts, derived once per parsed document. Every
 * rule that needs "the document's East Asian face", "is the docDefaults
 * language East Asian" or a compat switch reads them here instead of walking
 * the style chain itself.
 *
 * The body East Asian face is still chosen by five policies (one per rule
 * family) because they read the chain differently (empty theme slot, w:lang
 * backfill, a Korean face in every slot). They are named selectors over one
 * `BodyFaces` record so the differences are visible in one place; the
 * snapshot tool records all five, and a Word probe retires each difference.
 */
import type { DocDefaults, ParsedDocFull, StyleDisplay } from '@genoffice/docx-engine'
import {
  EMPTY_EA_THEME_FONT,
  isEastAsianLangTag,
  themeLangEaSlotFont,
} from '@genoffice/docx-engine'
import { isEaLayoutFace, isKoreanLayoutFace } from '../line-metrics'
import { isEastAsianFaceName, isJapaneseFaceName, isKoreanFaceName } from './face-info'

export interface BodyFaces {
  /** the w:default="1" paragraph style's display */
  normal: StyleDisplay | undefined
  dd: DocDefaults | undefined
  /** Normal's own East Asian face: a filled eastAsia slot, or a name differing from the ascii slot */
  normalEa: string | null
  /** Normal's EA-first name when it differs from the ascii slot (an explicit eastAsia slot
   *  naming the Latin face does not count; the line-factor and autospace policies read this) */
  normalEaDistinct: string | null
  /** Normal's EA slot is an empty theme slot (its face is a themeFontLang backfill) */
  normalEaSlotEmpty: boolean
  /** docDefaults East Asian face as parsed (declared, or backfilled for an empty theme slot) */
  ddEa: string | null
  ddEaSlotEmpty: boolean
  ddEaFromLang: boolean
  ddEaThemeSlot: boolean
  /** themeFontLang face of an empty docDefaults theme slot */
  ddThemeLangFace: string | null
  /** theme minor Latin face */
  themeMinorLatin: string | null
}

export interface DocumentTextContext {
  body: BodyFaces
  /** the East Asian face Word lays the body out with (three-way rule below) */
  bodyEaFace: BodyEastAsianFace
  /** docDefaults w:lang w:val names an East Asian language */
  eastAsianLang: boolean
  /** Latin body face (Normal ascii, then docDefaults, then Normal's shared name) */
  bodyAscii: string | null
  /** bodyEaFace.face (docBodyEaFont) */
  bodyEa: string | null
  /** the eastAsia slot as declared (Normal, then docDefaults), Latin names included: the
   *  Latin-via-eastAsia rule and empty marks read the slot, not the face Word falls back to */
  bodyEaSlot: string | null
  /** docDefaults w:lang lays every run's Latin out with its eastAsia face (line-metrics F8) */
  latinViaEaSlot: boolean
  /** balanceSingleByteDoubleByteWidth widens spaces: the flag plus useFELayout or a CJK body face */
  balancedSpaces: boolean
}

function defaultParaDisplay(parsed: ParsedDocFull): StyleDisplay | undefined {
  for (const info of parsed.styles.values()) {
    if (info.isDefault && info.type === 'paragraph' && info.display) return info.display
  }
  return undefined
}

function bodyFaces(parsed: ParsedDocFull): BodyFaces {
  const normal = defaultParaDisplay(parsed)
  const dd = parsed.docDefaults
  const normalEa =
    normal?.eastAsiaFont ??
    (normal?.font && normal.font !== normal.fontAscii ? normal.font : null) ??
    null
  return {
    normal,
    dd,
    normalEa,
    normalEaDistinct: normal?.font && normal.font !== normal.fontAscii ? normal.font : null,
    normalEaSlotEmpty: !!normal?.eaSlotEmpty,
    ddEa: dd?.eastAsiaFont ?? null,
    ddEaSlotEmpty: !!dd?.eaSlotEmpty,
    ddEaFromLang: !!dd?.eaFromLang,
    ddEaThemeSlot: !!dd?.eaThemeSlot,
    ddThemeLangFace:
      !dd?.eastAsiaFont && dd?.eaThemeSlot && parsed.themeFonts
        ? (themeLangEaSlotFont(parsed.themeFonts, 'minorEastAsia') ?? null)
        : null,
    themeMinorLatin: parsed.themeFonts?.minor ?? null,
  }
}

const cache = new WeakMap<object, DocumentTextContext>()

export function documentTextContext(parsed: ParsedDocFull): DocumentTextContext {
  let ctx = cache.get(parsed)
  if (ctx) return ctx
  const body = bodyFaces(parsed)
  const bodyEaFace = resolveBodyEastAsianFace(body, parsed)
  const bodyEa = bodyEaFace.face
  const balance = parsed.balanceDbcsSpacing === true
  ctx = {
    body,
    bodyEaFace,
    bodyEaSlot: body.normal?.eastAsiaFont ?? body.normalEaDistinct ?? body.ddEa,
    eastAsianLang: isEastAsianLangTag(parsed.docDefaults?.lang),
    bodyAscii: body.normal?.fontAscii ?? body.dd?.asciiFont ?? body.normal?.font ?? null,
    bodyEa,
    latinViaEaSlot: isEastAsianLangTag(parsed.docDefaults?.lang),
    balancedSpaces: balance && (parsed.useFELayout === true || bodyEaFace.declared),
  }
  cache.set(parsed, ctx)
  return ctx
}

// ─── the body East Asian face ────────────────────────────────────────────────
// Word probes 2026-11-04 (e1-e20) and 2026-10-20 (truth 429): the eastAsia slot
// of Normal, then docDefaults, decides in three ways. An East Asian-named face
// is the face. An empty theme slot takes the themeFontLang face (the theme
// script table, else ko Malgun Gothic / ja MS Mincho / zh or none SimSun). A
// Latin-named face in the slot (declared, or one name in every slot) is not an
// East Asian face at all: Word falls back per glyph to SimSun for Han/kana and
// Batang for hangul, both 1.30, and themeFontLang does not apply. Hangul under a
// non-Korean face goes to Malgun Gothic (Batang under SimSun and the Yu faces).

export interface BodyEastAsianFace {
  /** face Word lays Han and kana with */
  face: string
  /** face Word lays hangul with */
  hangulFace: string
  source: 'normal' | 'docDefaults' | 'emptySlot' | 'latinSlot' | 'none'
  /** an East Asian face was declared (sources normal / docDefaults) */
  declared: boolean
}

const SIMSUN = 'SimSun'
const BATANG = 'Batang'
const MALGUN = 'Malgun Gothic'

function hangulFallback(face: string): string {
  if (isKoreanFaceName(face) || isKoreanLayoutFace(face)) return face
  if (/yu (gothic|mincho)|yugoth|\u6e38/i.test(face.normalize('NFKC'))) return BATANG
  return isJapaneseFaceName(face) ? MALGUN : BATANG
}

function emptySlotFace(parsed: ParsedDocFull): string {
  const theme = parsed.themeFonts
  const fromLang = theme ? themeLangEaSlotFont(theme, 'minorEastAsia') : undefined
  if (fromLang) return fromLang
  // zh-CN without a script table keeps the parser's DengXian reading (probe F1); no language → SimSun (truth 429, e1/e6)
  return /^zh/i.test(theme?.eaLang ?? '') ? EMPTY_EA_THEME_FONT : SIMSUN
}

function resolveBodyEastAsianFace(body: BodyFaces, parsed: ParsedDocFull): BodyEastAsianFace {
  const done = (face: string, source: BodyEastAsianFace['source'], declared: boolean) => ({
    face,
    hangulFace: hangulFallback(face),
    source,
    declared,
  })
  // Normal's slot
  if (body.normal) {
    if (body.normalEaSlotEmpty) return done(emptySlotFace(parsed), 'emptySlot', false)
    // an explicit eastAsia slot (also one Latin name written into every slot);
    // a style naming only its Latin slots inherits the eastAsia slot below
    const slot = body.normal.eastAsiaFont ?? body.normalEaDistinct
    if (slot) {
      if (isEastAsianFaceName(slot) || isEaLayoutFace(slot)) return done(slot, 'normal', true)
      return { face: SIMSUN, hangulFace: BATANG, source: 'latinSlot', declared: false }
    }
    // one East Asian name in every slot (Meiryo / Malgun everywhere) is an East Asian choice
    const shared = body.normal.font
    if (
      shared &&
      shared === body.normal.fontAscii &&
      (isEastAsianFaceName(shared) || isEaLayoutFace(shared))
    )
      return done(shared, 'normal', true)
  }
  // docDefaults' slot
  if (body.ddEaSlotEmpty || (body.ddEaThemeSlot && !body.ddEa))
    return done(emptySlotFace(parsed), 'emptySlot', false)
  if (body.ddEa) {
    if (isEastAsianFaceName(body.ddEa) || isEaLayoutFace(body.ddEa))
      return done(body.ddEa, 'docDefaults', true)
    return { face: SIMSUN, hangulFace: BATANG, source: 'latinSlot', declared: false }
  }
  return { face: SIMSUN, hangulFace: BATANG, source: 'none', declared: false }
}

// ─── selectors over the unified face ────────────────────────────────────────

/** --doc-line-factor-cjk face */
export function eaFaceForCjkFactor(ctx: DocumentTextContext): string {
  return ctx.bodyEaFace.face
}

/** autospace document switch: the declared face that gates it (a bare docDefaults lets Normal gate).
 *  Reads the slots directly: the Noto superfamily gate is about what the document names, not what Word draws */
export function eaFaceForAutospace(ctx: DocumentTextContext): string | null {
  const { body } = ctx
  const normalEa =
    body.normal?.font &&
    !body.normalEaSlotEmpty &&
    (body.normalEaDistinct || isKoreanFaceName(body.normal.font))
      ? body.normal.font
      : null
  const ddDeclared = body.ddEa && !body.ddEaSlotEmpty && !body.ddEaFromLang ? body.ddEa : null
  return ddDeclared ?? (body.ddEaThemeSlot ? null : normalEa)
}

/** --doc-line-factor-kr face: the face Word lays hangul with */
export function eaFaceForKrFactor(ctx: DocumentTextContext): string {
  return ctx.bodyEaFace.hangulFace
}

/** typed-grid strut face: only a declared East Asian face (or a w:lang backfill) leads the strut */
export function eaFaceForGridStrut(ctx: DocumentTextContext): string | undefined {
  const { bodyEaFace, body } = ctx
  if (bodyEaFace.declared) return bodyEaFace.face
  if (bodyEaFace.source === 'emptySlot' && body.ddEaFromLang && body.ddEa) return body.ddEa
  return undefined
}
