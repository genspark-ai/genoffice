/**
 * Excel coerces text numbers with the session's own separators: in a pt-BR
 * session "15,68" is 15.68 and "=C6/5" computes, in an en-US session the same
 * text is not a number at all. Univer's engine parses text numbers with the
 * en-US rules only — a comma never reads as a decimal separator, so a
 * comma-decimal workbook's formulas refuse to compute at all ("15,68" →
 * #VALUE! everywhere; #2883).
 *
 * The wrap re-parses the strict comma-decimal shapes by the app language's
 * decimal separator before the engine sees the value. Every other shape —
 * dot decimals, the en "1,234" thousands group, "12,5%" — keeps the engine's
 * own path, so a dot-decimal session computes exactly as before.
 */
import { NumberValueObject, StringValueObject } from '@univerjs/engine-formula'

import type { Lang } from '@genoffice/i18n'

/** languages whose number formatting uses a comma decimal separator (Excel's
    locale table; conservative — ar/he conventions vary by region and stay out) */
const COMMA_DECIMAL_LANGS = new Set<Lang>([
  'pt',
  'fr',
  'de',
  'es',
  'it',
  'nl',
  'ru',
  'pl',
  'cs',
  'hu',
  'tr',
  'vi',
  'id',
  'ms',
])

/** strict comma-decimal shapes: an integer with a comma fraction ("5,47"),
    optionally with dotted thousands groups in front ("1.234,56"). A lone
    "1,234" is ambiguous and stays with the engine's own parse — under a
    comma-decimal session Excel reads it as 1.234, but every existing
    comma-grouped file the users already have would flip meaning the moment
    they switch the app language; the strict shapes carry the fix without
    touching that. */
const COMMA_DECIMAL = /^(?:\d{1,3}(?:\.\d{3})+|\d+),\d+$/

/** the numeric value of a strict comma-decimal string ("15,68" → 15.68) */
export function parseCommaDecimal(raw: string): number {
  return Number(raw.replace(/\./g, '').replace(',', '.'))
}

export function localeUsesCommaDecimal(lang: Lang): boolean {
  return COMMA_DECIMAL_LANGS.has(lang)
}

/**
 * Arms the locale-aware coercion for the session. `lang` is read per call, so
 * a language switch mid-session takes effect without a reinstall. A no-op on
 * dot-decimal languages.
 */
export function installFormulaLocaleCoercion(lang: () => Lang): { dispose(): void } {
  if (!localeUsesCommaDecimal(lang())) return { dispose() {} }
  const proto = StringValueObject.prototype as {
    convertToNumberObjectValue(): unknown
  }
  const original = proto.convertToNumberObjectValue
  proto.convertToNumberObjectValue = function (this: StringValueObject) {
    const raw = this.getValue()
    if (typeof raw === 'string') {
      const candidate = raw.trim()
      if (COMMA_DECIMAL.test(candidate))
        return NumberValueObject.create(parseCommaDecimal(candidate))
    }
    return original.call(this)
  }
  return {
    dispose() {
      proto.convertToNumberObjectValue = original
    },
  }
}
