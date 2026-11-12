/**
 * Sensitive-information patterns for Search & Redact (Acrobat's Find Text ▸ Patterns):
 * phone numbers, email addresses, credit card numbers, US Social Security numbers,
 * dates and IBANs. Ported from pdfcraft's crates/redact/src/patterns.rs. Each matcher
 * is deterministic and returns [start, end) offsets into the page text, which the
 * caller maps onto text-item geometry.
 */

export type RedactPattern = 'phone' | 'email' | 'credit-card' | 'ssn' | 'date' | 'iban'

export const REDACT_PATTERNS: readonly RedactPattern[] = [
  'phone',
  'email',
  'credit-card',
  'ssn',
  'date',
  'iban',
]

type S = string

const isDigit = (c: string | undefined): boolean => c !== undefined && c >= '0' && c <= '9'
const ALNUM_RE = /^[\p{Alphabetic}\p{N}]$/u
const ALPHA_RE = /^\p{Alphabetic}$/u
const isAlnum = (c: string | undefined): boolean => c !== undefined && ALNUM_RE.test(c)
const isAlpha = (c: string | undefined): boolean => c !== undefined && ALPHA_RE.test(c)
const isAsciiAlpha = (c: string | undefined): boolean => c !== undefined && /^[A-Za-z]$/.test(c)
const isAsciiAlnum = (c: string | undefined): boolean => c !== undefined && /^[A-Za-z0-9]$/.test(c)
const at = (s: S, i: number): string | undefined => (i >= 0 && i < s.length ? s[i] : undefined)

/** Digits and the separators allowed between them, from `start`: [end, digit count]. */
function digitRun(s: S, start: number, seps: string, maxLen: number): [number, number] {
  let i = start
  let n = 0
  while (i < s.length && i - start < maxLen) {
    const c = s[i]!
    if (isDigit(c)) {
      n++
    } else {
      const joins =
        seps.includes(c) &&
        (isDigit(at(s, i + 1)) || at(s, i + 1) === '(') &&
        // A slash only ever joins two digits ("0426/123456").
        (c !== '/' || (isDigit(at(s, i + 1)) && isDigit(at(s, i - 1))))
      const paren = (c === '(' || c === ')') && seps.includes(c)
      if (!joins && !paren) break
    }
    i++
  }
  while (i > start && !isDigit(s[i - 1])) i--
  return [i, n]
}

/** Whether the separated digit run goes on past `k` in direction `step` (±1). */
function runContinues(s: S, k: number, step: 1 | -1, seps: string): boolean {
  for (let j = k + step; ; j += step) {
    const c = at(s, j)
    if (c === undefined) return false
    if (isDigit(c)) return true
    if (!seps.includes(c)) return false
  }
}

const PHONE_LABELS = new Set([
  'tel',
  'tél',
  'tele',
  'telefon',
  'telefono',
  'teléfono',
  'telephone',
  'téléphone',
  'phone',
  'ph',
  'cell',
  'cel',
  'cellulare',
  'cellphone',
  'mobile',
  'mob',
  'mobil',
  'móvil',
  'movil',
  'handy',
  'fax',
  'tlf',
  'telf',
  'gsm',
])

/** Whether the text before `i` is a phone label, with one to four spaces, colons or dots between. */
function labelled(s: S, i: number): boolean {
  let j = i
  while (i - j < 4 && ' :.\u00a0'.includes(at(s, j - 1) ?? 'x')) j--
  const end = j
  while (end - j < 12 && isAlpha(at(s, j - 1))) j--
  if (end === i || isAlnum(at(s, j - 1))) return false
  return PHONE_LABELS.has(s.slice(j, end).toLowerCase())
}

function phone(s: S, i: number): number | null {
  const SEPS = ' -.()'
  // A slash joins the area code to the number in Italian and German landlines ("0426/123456").
  const RUN_SEPS = ' -.()/'
  const c0 = s[i]!
  const startOk = ((c0 === '+' || c0 === '(') && isDigit(at(s, i + 1))) || isDigit(c0)
  if (!startOk || isAlnum(at(s, i - 1)) || at(s, i - 1) === '+') return null
  // The tail of a longer grouped number (an IBAN, a company or social security number) is not a phone number.
  if (runContinues(s, i, -1, SEPS)) return null
  const plus = c0 === '+'
  const first = plus ? i + 1 : i
  const [end, n] = digitRun(s, first, RUN_SEPS, 28)
  if (end <= first) return null
  const run = s.slice(first, end)
  const seps = [...run].filter((c) => !isDigit(c)).length
  if ((seps === 0 && !labelled(s, i)) || isDigit(at(s, end)) || runContinues(s, end - 1, 1, SEPS))
    return null
  const slash = run.includes('/')
  // Italian geographic numbers keep their 0 after +39; only a labelled number takes a slash there.
  let slashRun = run
  if (plus && slash) {
    if (!labelled(s, i) || n < 8 || n > 13 || !run.startsWith('39')) return null
    slashRun = run.slice(2)
    if (slashRun.startsWith(' ')) slashRun = slashRun.slice(1)
  }
  // Only an area code beginning with 0 takes a slash, and only one, right after it.
  const slashes = [...run].filter((c) => c === '/').length
  if (
    slashes > 1 ||
    (slash &&
      (slashRun[0] !== '0' ||
        !isDigit(slashRun[1]) ||
        ![...slashRun.slice(0, slashRun.indexOf('/'))].every(isDigit)))
  )
    return null
  let ok: boolean
  if (plus) {
    // E.164: country code and up to 15 digits in all.
    ok = n >= 8 && n <= 15
  } else {
    // North American: 10 digits, 11 with a country code.
    const nanp = n >= 10 && n <= 11 && !slash
    // National with a trunk 0; the first group needs two digits so decimals don't
    // match, and day-first dates are left to the date pattern.
    const digitsFrom = run.replace(/^[^0-9]+/, '')
    const trunkZero = digitsFrom[0] === '0'
    const lead = /^[0-9]*/.exec(digitsFrom)![0].length
    const trunk = trunkZero && n >= 7 && n <= 15 && lead >= 2 && date(s, i) === null
    ok = nanp || trunk
  }
  return ok ? end - i : null
}

/** Scripts written without spaces between words: their letters never join an address. */
function spaceless(c: string): boolean {
  const u = c.codePointAt(0)!
  return (
    (u >= 0x0e00 && u <= 0x0eff) ||
    (u >= 0x1000 && u <= 0x109f) ||
    (u >= 0x1100 && u <= 0x11ff) ||
    (u >= 0x1780 && u <= 0x17ff) ||
    (u >= 0x2e80 && u <= 0x9fff) ||
    (u >= 0xa960 && u <= 0xa97f) ||
    (u >= 0xac00 && u <= 0xd7ff) ||
    (u >= 0xf900 && u <= 0xfaff) ||
    (u >= 0xff00 && u <= 0xffef)
  )
}

/** An address letter or digit: ASCII, or of a script written with spaces (IDN / RFC 6531). */
const addressChar = (c: string | undefined): boolean =>
  c !== undefined && (isAsciiAlnum(c) || (isAlnum(c) && !spaceless(c)))

function email(s: S, i: number): number | null {
  const local = (c: string | undefined): boolean =>
    c !== undefined && (addressChar(c) || '._%+-'.includes(c))
  if (!local(s[i]) || local(at(s, i - 1))) return null
  let j = i
  while (j < s.length && local(s[j])) j++
  if (s[j] !== '@' || ![...s.slice(i, j)].some(addressChar)) return null
  const d0 = j + 1
  let k = d0
  while (
    k < s.length &&
    (addressChar(s[k]) || s[k] === '-' || (s[k] === '.' && addressChar(at(s, k + 1))))
  )
    k++
  const domain = s.slice(d0, k)
  const tld = domain.slice(domain.lastIndexOf('.') + 1)
  return domain.includes('.') && [...tld].length >= 2 && [...tld].every(isAlpha) ? k - i : null
}

export function luhn(digits: readonly number[]): boolean {
  let sum = 0
  for (let k = 0; k < digits.length; k++) {
    let d = digits[digits.length - 1 - k]!
    if (k % 2 === 1) {
      d *= 2
      if (d > 9) d -= 9
    }
    sum += d
  }
  return sum % 10 === 0
}

function creditCard(s: S, i: number): number | null {
  if (!isDigit(s[i]) || isDigit(at(s, i - 1))) return null
  const [end, n] = digitRun(s, i, ' -', 29)
  if (end === i || runContinues(s, end - 1, 1, ' -')) return null
  // The run may carry a trailing expiry, year or CVV: try it, then each shorter run
  // ending at a group boundary, dropping at most four digits.
  let count = n
  for (let e = end; e > i; e--) {
    if (n - count > 4) break
    if (!isDigit(s[e - 1])) continue
    const boundary = e === end ? !isDigit(at(s, end)) : s[e] === ' ' || s[e] === '-'
    if (boundary && count >= 13 && count <= 19) {
      const digits = [...s.slice(i, e)].filter(isDigit).map(Number)
      if (luhn(digits)) return e - i
    }
    count = Math.max(0, count - 1)
  }
  return null
}

function ssn(s: S, i: number): number | null {
  if (isDigit(at(s, i - 1))) return null
  const g = (from: number, n: number): boolean => {
    for (let k = from; k < from + n; k++) if (!isDigit(at(s, k))) return false
    return true
  }
  const sep = (k: number): boolean => s[k] === '-' || s[k] === ' '
  if (
    !(g(i, 3) && sep(i + 3) && g(i + 4, 2) && sep(i + 6) && g(i + 7, 4)) ||
    isDigit(at(s, i + 11))
  )
    return null
  const area = s.slice(i, i + 3)
  return area !== '000' && area !== '666' && area[0] !== '9' ? 11 : null
}

const MONTHS = [
  'january february march april may june july august september october november december',
  'ocak şubat mart nisan mayıs haziran temmuz ağustos eylül ekim kasım aralık',
  'januar februar märz april mai juni juli august september oktober november dezember',
  'janvier février mars avril mai juin juillet août septembre octobre novembre décembre',
  'gennaio febbraio marzo aprile maggio giugno luglio agosto settembre ottobre novembre dicembre',
].flatMap((l) => l.split(' '))

const MONTH_ABBREVIATIONS = [
  'jan feb mar apr may jun jul aug sep oct nov dec',
  'oca şub mar nis may haz tem ağu eyl eki kas ara',
].flatMap((l) => l.split(' '))

/** One lowercase character per character, Turkish dotted and dotless i folded to i. */
function fold(c: string): string {
  if (c === 'I' || c === '\u0130' || c === '\u0131') return 'i'
  return c.toLowerCase()[0] ?? c
}

/** Characters `name` covers at `i` as a whole word (any case). */
function wordAt(s: S, i: number, name: string): number | null {
  for (let n = 0; n < name.length; n++) {
    const c = at(s, i + n)
    if (c === undefined || fold(c) !== fold(name[n]!)) return null
  }
  return isAlpha(at(s, i + name.length)) ? null : name.length
}

function month(s: S, i: number): number | null {
  if (isAlnum(at(s, i - 1))) return null
  let best: number | null = null
  for (const m of MONTHS) {
    const n = wordAt(s, i, m)
    if (n !== null && (best === null || n > best)) best = n
  }
  if (best !== null) return best
  for (const m of MONTH_ABBREVIATIONS) {
    const n = wordAt(s, i, m)
    if (n !== null) return s[i + n] === '.' ? n + 1 : n
  }
  return null
}

function number(s: S, i: number, min: number, max: number): number | null {
  let n = 0
  while (n <= max && isDigit(at(s, i + n))) n++
  return n >= min && n <= max ? n : null
}

function date(s: S, i: number): number | null {
  if (isAlnum(at(s, i - 1))) return null
  const spaces = (k: number): number => {
    let n = 0
    while (s[k + n] === ' ') n++
    return n
  }
  const a = number(s, i, 1, 4)
  if (a !== null) {
    // Numeric: m/d/yy(yy), m-d-yyyy, d.m.yyyy, yyyy-mm-dd.
    const sep = at(s, i + a)
    if (sep === '-' || sep === '.' || sep === '/') {
      const b = number(s, i + a + 1, 1, 2)
      if (b !== null && s[i + a + 1 + b] === sep) {
        const c = number(s, i + a + b + 2, a === 4 ? 1 : 2, a === 4 ? 2 : 4)
        if (c !== null && (a <= 2 || a === 4) && !isDigit(at(s, i + a + b + 2 + c)))
          return a + b + c + 2
      }
    }
    // "5 January 2024", "5 Ekim 2024", "5. Oktober 2024".
    const day = i + a + (s[i + a] === '.' ? 1 : 0)
    const k = day + spaces(day)
    if (a <= 2 && k > day) {
      const m = month(s, k)
      if (m !== null) {
        const k2 = k + m + spaces(k + m)
        const y = number(s, k2, 4, 4)
        if (y !== null) return k2 + y - i
      }
    }
    return null
  }
  // "January 5, 2024" / "Jan. 5 2024".
  const m = month(s, i)
  if (m === null) return null
  const k = i + m + spaces(i + m)
  const d = number(s, k, 1, 2)
  if (d === null) return null
  let k2 = k + d
  if (s[k2] === ',') k2++
  k2 += spaces(k2)
  const y = number(s, k2, 4, 4)
  return y === null ? null : k2 + y - i
}

/** Exact IBAN length per country, SWIFT IBAN registry release 99 (Dec 2024). */
export const IBAN_LENGTHS: Readonly<Record<string, number>> = Object.fromEntries(
  (
    'AD24 AE23 AL28 AT20 AZ28 BA20 BE16 BG22 BH22 BI27 BR29 BY28 CH21 CR22 CY28 CZ24 DE22 DJ27 DK18 ' +
    'DO28 EE20 EG29 ES24 FI18 FK18 FO18 FR27 GB22 GE22 GI23 GL18 GR27 GT28 HN28 HR21 HU28 IE22 IL23 ' +
    'IQ23 IS26 IT27 JO30 KW30 KZ20 LB28 LC32 LI21 LT20 LU20 LV21 LY25 MC27 MD24 ME22 MK19 MN20 MR27 ' +
    'MT31 MU30 NI28 NL18 NO15 OM23 PK24 PL28 PS29 PT25 QA29 RO24 RS22 RU33 SA24 SC31 SD18 SE24 SI19 ' +
    'SK24 SM27 SO23 ST25 SV28 TL23 TN24 TR26 UA29 VA22 VG24 XK20 YE30'
  )
    .split(' ')
    .map((e) => [e.slice(0, 2), Number(e.slice(2))]),
)

/** ISO 13616 mod-97: move the first four characters to the end, letters A=10..Z=35. */
export function ibanCheck(chars: string): boolean {
  let rem = 0
  for (const c of chars.slice(4) + chars.slice(0, 4)) {
    rem = isDigit(c) ? (rem * 10 + Number(c)) % 97 : (rem * 100 + c.charCodeAt(0) - 55) % 97
  }
  return rem === 1
}

function iban(s: S, i: number): number | null {
  if (!(
    isAsciiAlpha(s[i]) &&
    isAsciiAlpha(at(s, i + 1)) &&
    isDigit(at(s, i + 2)) &&
    isDigit(at(s, i + 3))
  ))
    return null
  if (isAlnum(at(s, i - 1))) return null
  const len = IBAN_LENGTHS[s.slice(i, i + 2).toUpperCase()]
  if (!len) return null
  let chars = ''
  let j = i
  let afterSpace = false
  // A single space is accepted only after a full group of four.
  while (chars.length < len) {
    const c = at(s, j)
    if (isAsciiAlnum(c)) {
      chars += c!.toUpperCase()
      afterSpace = false
    } else if (c === ' ' && chars.length % 4 === 0 && !afterSpace) {
      afterSpace = true
    } else {
      return null
    }
    j++
  }
  if (isAlnum(at(s, j))) return null
  return ibanCheck(chars) ? j - i : null
}

const MATCHERS: Record<RedactPattern, (s: S, i: number) => number | null> = {
  phone,
  email,
  'credit-card': creditCard,
  ssn,
  date,
  iban,
}

/** Non-overlapping [start, end) ranges of `text` matching `pattern`, left to right. */
export function findPattern(pattern: RedactPattern, text: string): [number, number][] {
  const f = MATCHERS[pattern]
  const out: [number, number][] = []
  for (let i = 0; i < text.length;) {
    const n = f(text, i)
    if (n !== null && n > 0) {
      out.push([i, i + n])
      i += n
    } else {
      i++
    }
  }
  return out
}
