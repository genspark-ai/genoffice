/**
 * JSON.parse for model-written specs, with one repair pass when the strict parse fails: an
 * unescaped " inside a string value is escaped when the next non-whitespace character is not
 * structural. A low-9 opening quote („, U+201E: Hungarian, German, Czech, Polish…) is often closed
 * with a plain " (`a „hónap terméke": …`); while one is open, the next " is its closer and is escaped
 * even when , or : follows it. The strict parse error is the one reported.
 */
export function parseLenientJson(
  text: string,
): { ok: true; value: unknown } | { ok: false; error: string } {
  try {
    return { ok: true, value: JSON.parse(text) }
  } catch (e) {
    for (const lowQuotes of [true, false]) {
      try {
        const value: unknown = JSON.parse(escapeStrayQuotes(text, lowQuotes))
        if (!looksMisrepaired(value)) return { ok: true, value }
      } catch {
        /* try the next repair; the strict error is reported below */
      }
    }
    return { ok: false, error: e instanceof Error ? e.message : String(e) }
  }
}

/** `, "key":` right after a quote: the string ended and the next member starts. */
const KEY_BOUNDARY = /^\s*,\s*"[^"\\\n]{1,64}"\s*:/
/** A swallowed member boundary inside a string: `…", "key": …`. */
const SWALLOWED_MEMBER = /",\s*"[\w.-]{1,64}"\s*:/

/**
 * A repair that parsed but merged two members into one, or turned value text into a key, is
 * wrong data; failing lets the caller ask the model again instead.
 */
export function looksMisrepaired(value: unknown): boolean {
  if (typeof value === 'string') return SWALLOWED_MEMBER.test(value)
  if (Array.isArray(value)) return value.some(looksMisrepaired)
  if (value && typeof value === 'object') {
    return Object.entries(value).some(([k, v]) => /[",]/.test(k) || looksMisrepaired(v))
  }
  return false
}

export function escapeStrayQuotes(json: string, lowQuotes = true): string {
  let out = ''
  let inStr = false
  let lowQuoteOpen = false
  for (let i = 0; i < json.length; i++) {
    const c = json[i]!
    if (!inStr) {
      if (c === '"') {
        inStr = true
        lowQuoteOpen = false
      }
      out += c
      continue
    }
    if (c === '\\') {
      out += c + (json[++i] ?? '')
      continue
    }
    if (c === '„') lowQuoteOpen = lowQuotes
    else if (c === '”' || c === '“') lowQuoteOpen = false
    if (c === '"') {
      // A " that starts the next key (`", "key":`) is never a low-9 closer.
      if (lowQuoteOpen && !KEY_BOUNDARY.test(json.slice(i + 1, i + 80))) {
        lowQuoteOpen = false
        out += '\\"'
        continue
      }
      lowQuoteOpen = false
      let j = i + 1
      while (j < json.length && ' \n\r\t'.includes(json[j]!)) j++
      const next = json[j]
      if (next === undefined || ',}]:'.includes(next)) {
        inStr = false
        out += c
      } else {
        out += '\\"'
      }
      continue
    }
    out += c
  }
  return out
}
