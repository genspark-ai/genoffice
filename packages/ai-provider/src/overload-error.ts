/**
 * Classifies capacity/rate-limit failures (HTTP 429/503/529, gateway
 * "engine overloaded" notices, provider rate limits…) so the apps can show a
 * localized "the AI service is busy, try again shortly" message (errorCode
 * 'overloaded') instead of the raw HTTP body dump.
 */

// Status markers embedded by the protocol layers ("HTTP 429: …", "Claude HTTP 529: …")
// plus the notice texts the gateways/providers put in error bodies.
const OVERLOADED_PATTERN = new RegExp(
  [
    '\\bHTTP (429|503|529)\\b',
    'overload', // "overloaded", "engine_overloaded_error", Anthropic's "Overloaded"
    'rate.?limit',
    'too many requests',
    'resource.{0,12}exhausted', // Gemini RESOURCE_EXHAUSTED / "Resource has been exhausted"
    'quota exceeded',
  ].join('|'),
  'i',
)

// Credits-exhausted notices ("Your Genspark credits have been exhausted…") are
// a different failure class (errorCode 'credits', "top up" message) — never
// misreport them as a transient capacity problem.
const CREDITS_PATTERN = /credit|pricing/i

// "credits per minute exceeded" is a rate limit; only the exhausted/insufficient wording is billing.
const CREDITS_EXHAUSTED_PATTERN =
  /(credits?[^\n]{0,40}(exhausted|insufficient)|(exhausted|insufficient)[^\n]{0,40}credits?)/i

// A bare 429/503/529 status marker makes it transient even when credits are named.
const HTTP_STATUS_PATTERN = /\bHTTP (429|503|529)\b/i

function matches(text: string): boolean {
  if (!OVERLOADED_PATTERN.test(text)) return false
  if (CREDITS_EXHAUSTED_PATTERN.test(text)) return false
  if (CREDITS_PATTERN.test(text) && !HTTP_STATUS_PATTERN.test(text)) return false
  return true
}

/**
 * True when the error (an Error, its `cause` chain, or a plain error string)
 * looks like a transient capacity/rate-limit failure that a later retry can
 * resolve. Works on message text so it covers both thrown HTTP errors and
 * error notices delivered inside a 200 SSE stream.
 */
export function isAiOverloadedError(err: unknown): boolean {
  let current: unknown = err
  for (let depth = 0; current && depth < 5; depth++) {
    if (typeof current === 'string') return matches(current)
    const e = current as { message?: unknown; cause?: unknown }
    if (typeof e.message === 'string' && matches(e.message)) return true
    current = e.cause
  }
  return false
}
