/**
 * Parses the LLM JSON output of generate_deck / planDeckOutline.
 * Models often put unescaped quotes, trailing commas and smart quotes inside brief/title,
 * making JSON.parse fail outright.
 */

export type OutlineJson = {
  core_hook?: unknown
  pages?: unknown
  [key: string]: unknown
}

/** Strip markdown fences and slice the object between the first and last braces. */
export function extractJsonObject(text: string): string {
  let s = String(text ?? '').trim()
  s = s.replace(/^```[a-zA-Z]*\s*/, '').replace(/```\s*$/, '').trim()
  const st = s.indexOf('{')
  const en = s.lastIndexOf('}')
  if (st >= 0 && en > st) s = s.slice(st, en + 1)
  return s
}

/** Remove trailing commas in objects/arrays: {"a":1,} / [1,] */
function stripTrailingCommas(s: string): string {
  return s.replace(/,\s*([}\]])/g, '$1')
}

/** Curly quotes (CJK and Western) → straight quotes, so parse doesn't blow up. */
function normalizeSmartQuotes(s: string): string {
  return s
    .replace(/[\u201C\u201D\u201E\u201F\u2033\u2036]/g, '"')
    .replace(/[\u2018\u2019\u201A\u201B\u2032\u2035]/g, "'")
}

/** `, "key":` right after a quote: the string ended and the next member starts. */
const KEY_BOUNDARY = /^\s*,\s*"[^"\\\n]{1,64}"\s*:/
/** A swallowed member boundary inside a string: `…", "key": …`. */
const SWALLOWED_MEMBER = /",\s*"[\w.-]{1,64}"\s*:/

/**
 * A repair that parsed but merged two members into one, or turned value text into a key, is
 * wrong data; failing lets the caller ask the model again instead.
 */
function looksMisrepaired(value: unknown): boolean {
  if (typeof value === 'string') return SWALLOWED_MEMBER.test(value)
  if (Array.isArray(value)) return value.some(looksMisrepaired)
  if (value && typeof value === 'object') {
    return Object.entries(value).some(([k, v]) => /[",]/.test(k) || looksMisrepaired(v))
  }
  return false
}

/**
 * Escape unescaped " inside string values, and write bare newlines/tabs as JSON escapes.
 * Heuristic: if the character after " (skipping whitespace) isn't , } ] : or the end, treat it
 * as an in-value quote.
 */
export function escapeBrokenStringLiterals(json: string, lowQuotes = true): string {
  let out = ''
  let i = 0
  let inString = false
  // A low-9 opening quote („, U+201E) is often closed with a plain "; while one is open, the next
  // " is its closer and is escaped even when , or : follows it (`a „hónap terméke": …`).
  let lowQuoteOpen = false
  while (i < json.length) {
    const c = json[i]!
    if (!inString) {
      out += c
      if (c === '"') {
        inString = true
        lowQuoteOpen = false
      }
      i++
      continue
    }
    if (c === '\\') {
      out += c + (json[i + 1] ?? '')
      i += 2
      continue
    }
    if (c === '\u201E') lowQuoteOpen = lowQuotes
    else if (c === '\u201D' || c === '\u201C') lowQuoteOpen = false
    // A " that starts the next key (`", "key":`) is never a low-9 closer.
    if (c === '"' && lowQuoteOpen && !KEY_BOUNDARY.test(json.slice(i + 1, i + 80))) {
      lowQuoteOpen = false
      out += '\\"'
      i++
      continue
    }
    if (c === '"') {
      lowQuoteOpen = false
      let j = i + 1
      while (j < json.length && /\s/.test(json[j]!)) j++
      const next = json[j]
      if (next === undefined || next === ',' || next === '}' || next === ']' || next === ':') {
        out += '"'
        inString = false
      } else {
        out += '\\"'
      }
      i++
      continue
    }
    if (c === '\n') { out += '\\n'; i++; continue }
    if (c === '\r') { out += '\\r'; i++; continue }
    if (c === '\t') { out += '\\t'; i++; continue }
    out += c
    i++
  }
  return out
}

function tryParse(s: string): OutlineJson | null {
  try {
    const obj = JSON.parse(s) as OutlineJson
    return obj && typeof obj === 'object' ? obj : null
  } catch {
    return null
  }
}

/**
 * Multi-layer fallback parsing of outline JSON.
 * Returns the object on success; null on failure (caller composes the error message).
 */
export function parseOutlineJson(text: string): OutlineJson | null {
  const raw = extractJsonObject(text)
  if (!raw) return null

  const candidates = [
    raw,
    stripTrailingCommas(raw),
    normalizeSmartQuotes(raw),
    stripTrailingCommas(normalizeSmartQuotes(raw)),
  ]
  // Quote repairs guess where strings end; a guess that merged members is rejected, not returned.
  const repaired = [
    escapeBrokenStringLiterals(raw),
    escapeBrokenStringLiterals(raw, false),
    escapeBrokenStringLiterals(normalizeSmartQuotes(raw)),
    stripTrailingCommas(escapeBrokenStringLiterals(normalizeSmartQuotes(raw))),
  ]

  const seen = new Set<string>()
  for (const c of candidates) {
    if (seen.has(c)) continue
    seen.add(c)
    const obj = tryParse(c)
    if (obj) return obj
  }
  for (const c of repaired) {
    if (seen.has(c)) continue
    seen.add(c)
    const obj = tryParse(c)
    if (obj && !looksMisrepaired(obj)) return obj
  }
  return null
}
