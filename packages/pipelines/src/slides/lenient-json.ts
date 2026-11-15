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
        return { ok: true, value: JSON.parse(escapeStrayQuotes(text, lowQuotes)) }
      } catch {
        /* try the next repair; the strict error is reported below */
      }
    }
    return { ok: false, error: e instanceof Error ? e.message : String(e) }
  }
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
      if (lowQuoteOpen) {
        lowQuoteOpen = false
        out += '\\"'
        continue
      }
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
