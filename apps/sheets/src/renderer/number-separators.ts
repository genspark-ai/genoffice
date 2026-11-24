/**
 * Region number separators, the way Excel follows them: a decimal-comma region
 * shows and accepts `5,47` and `=SUM(A1;B1)*1,5`, while the model, the file and
 * every non-UI consumer keep the canonical `5.47` and `=SUM(A1,B1)*1.5`.
 * Everything here is text-to-text, so the canonical parsers stay the only ones
 * that decide what a number or a formula is.
 */
import { numfmt } from '@univerjs/core'

export interface NumberSeparators {
  decimal: '.' | ','
  group: string
}

export const CANONICAL_SEPARATORS: NumberSeparators = { decimal: '.', group: ',' }

export const usesDecimalComma = (sep: NumberSeparators): boolean => sep.decimal === ','

const HOST_NUMFMT_LOCALE = 'zz-ZZ'
let current = CANONICAL_SEPARATORS
let suspended = 0
const listeners = new Set<() => void>()

/** Separators the grid shows and the editors take; canonical inside withCanonicalNumbers. */
export const getNumberSeparators = (): NumberSeparators =>
  suspended > 0 ? CANONICAL_SEPARATORS : current

/** numfmt `locale` option for display text: English names, the region's separators. */
export const numfmtDisplayLocale = (): string =>
  usesDecimalComma(getNumberSeparators()) ? HOST_NUMFMT_LOCALE : 'en'

export function setNumberSeparators(sep: NumberSeparators): void {
  current = sep
  if (usesDecimalComma(sep)) {
    numfmt.addLocale(
      { ...numfmt.getLocale('en'), decimal: sep.decimal, group: sep.group },
      HOST_NUMFMT_LOCALE,
    )
  }
  for (const listener of listeners) listener()
}

/** Fires whenever the text a number displays as may have changed (display caches drop). */
export function onNumberSeparatorsChange(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

/** Runs `fn` with canonical display text: file exports must not follow the region. */
export function withCanonicalNumbers<T>(fn: () => T): T {
  const notify = usesDecimalComma(current) && suspended === 0
  suspended += 1
  if (notify) for (const listener of listeners) listener()
  try {
    return fn()
  } finally {
    suspended -= 1
    if (notify) for (const listener of listeners) listener()
  }
}

/** Separators of a BCP 47 region tag; anything but a decimal comma stays canonical. */
export function separatorsForLocale(tag: string | null | undefined): NumberSeparators {
  if (!tag) return CANONICAL_SEPARATORS
  try {
    const parts = new Intl.NumberFormat(tag, { numberingSystem: 'latn' }).formatToParts(1234567.5)
    if (parts.find((part) => part.type === 'decimal')?.value !== ',') return CANONICAL_SEPARATORS
    const group = parts.find((part) => part.type === 'group')?.value ?? '.'
    return { decimal: ',', group: /\s/.test(group) ? ' ' : group }
  } catch {
    return CANONICAL_SEPARATORS
  }
}

const escapeRe = (text: string): string => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

const localNumberRes = new Map<string, RegExp>()

function localNumberRe(sep: NumberSeparators): RegExp {
  let re = localNumberRes.get(sep.group)
  if (!re) {
    // a space-grouping region types any space, and Excel takes the dot too
    const group = /\s/.test(sep.group) ? '[\\s.]' : escapeRe(sep.group)
    const affix = '[\\s+\\-()%\\p{Sc}]*'
    re = new RegExp(
      `^(${affix})(\\d{1,3}(?:${group}\\d{3})+|\\d*)(?:,(\\d*))?(e[+-]?\\d+)?(${affix})$`,
      'iu',
    )
    localNumberRes.set(sep.group, re)
  }
  return re
}

/**
 * A typed or pasted number in the region's notation, rewritten in canonical
 * notation (`1.234,5` → `1,234.5`); null when the text is not one, so dates,
 * times and plain text pass through untouched.
 */
export function localNumberToCanonical(text: string, sep: NumberSeparators): string | null {
  if (!usesDecimalComma(sep) || text.length > 64) return null
  const match = localNumberRe(sep).exec(text)
  if (!match) return null
  const [, lead = '', whole = '', fraction, exponent = '', tail = ''] = match
  if (!whole && !fraction) return null
  const digits = whole.replace(/\D/g, ',')
  const canonical = `${lead}${digits}${fraction === undefined ? '' : `.${fraction}`}${exponent}${tail}`
  return canonical === text ? null : canonical
}

const CANONICAL_NUMBER_RE = /^-?\d*\.?\d+(?:e[+-]?\d+)?%?$/i

/** The editor's text for a numeric cell value (`5.47` → `5,47`). */
export function canonicalNumberToLocal(value: string | number, sep: NumberSeparators): string {
  const text = String(value)
  if (!usesDecimalComma(sep) || !CANONICAL_NUMBER_RE.test(text)) return text
  return text.replace('.', ',')
}

const isDigit = (ch: string | undefined): boolean => ch !== undefined && ch >= '0' && ch <= '9'
const isNameChar = (ch: string | undefined): boolean =>
  ch !== undefined && /[\p{L}\p{N}_.$\\?]/u.test(ch)

/**
 * Walks a formula outside string literals, quoted sheet names and bracketed
 * references, handing each number literal and separator to `rewrite`.
 */
function rewriteFormula(
  formula: string,
  number: (literal: string, next: string | undefined, joined: boolean) => string | null,
  separator: (ch: string, inArray: boolean) => string,
  numberBody: RegExp,
): string {
  if (!formula.startsWith('=')) return formula
  let out = ''
  let braces = 0
  let brackets = 0
  let i = 0
  while (i < formula.length) {
    const ch = formula[i]!
    if (ch === '"' || ch === "'") {
      let end = i + 1
      while (end < formula.length) {
        if (formula[end] === ch) {
          if (formula[end + 1] !== ch) break
          end += 1
        }
        end += 1
      }
      out += formula.slice(i, end + 1)
      i = end + 1
      continue
    }
    if (ch === '[') brackets += 1
    else if (ch === ']') brackets = Math.max(0, brackets - 1)
    if (brackets > 0) {
      out += ch
      i += 1
      continue
    }
    if (ch === '{') braces += 1
    else if (ch === '}') braces = Math.max(0, braces - 1)
    // inside an array constant a backslash is the column separator, not a name character
    const joins = (other: string | undefined): boolean =>
      isNameChar(other) && !(braces > 0 && other === '\\')
    if (isDigit(ch) && !joins(formula[i - 1])) {
      numberBody.lastIndex = i
      const literal = numberBody.exec(formula)?.[0] ?? ch
      const next = formula[i + literal.length]
      const rewritten = number(literal, next, joins(next))
      if (rewritten !== null) {
        out += rewritten
        i += literal.length
        continue
      }
      // not a number after all (a reference such as 1:5): copy the digits as they are
      const digits = /^\d+/.exec(literal)![0]
      out += digits
      i += digits.length
      continue
    }
    out += separator(ch, braces > 0)
    i += 1
  }
  return out
}

const CANONICAL_LITERAL = /\d+(?:\.\d+)?(?:e[+-]?\d+)?/iy
const LOCAL_LITERAL = /\d+(?:,\d+)?(?:e[+-]?\d+)?/iy

/** `=SUM(A1,B1)*1.5` → `=SUM(A1;B1)*1,5`; array constants `{1,2;3,4}` → `{1\2;3\4}`. */
export function formulaToLocal(formula: string, sep: NumberSeparators): string {
  if (!usesDecimalComma(sep)) return formula
  return rewriteFormula(
    formula,
    (literal, next, joined) => (joined && next !== '.' ? null : literal.replace('.', ',')),
    (ch, inArray) => (ch === ',' ? (inArray ? '\\' : ';') : ch),
    CANONICAL_LITERAL,
  )
}

/**
 * `=SUM(A1;B1)*1,5` → `=SUM(A1,B1)*1.5`. A comma that is not inside a number
 * is left as the canonical separator it already is, so a formula pasted in
 * canonical notation still works.
 */
export function formulaToCanonical(formula: string, sep: NumberSeparators): string {
  if (!usesDecimalComma(sep)) return formula
  return rewriteFormula(
    formula,
    (literal, _next, joined) => (joined ? null : literal.replace(',', '.')),
    (ch, inArray) => {
      if (ch === ';') return inArray ? ';' : ','
      if (ch === '\\' && inArray) return ','
      return ch
    },
    LOCAL_LITERAL,
  )
}
