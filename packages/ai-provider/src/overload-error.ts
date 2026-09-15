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

// Credits / quota-exhausted notices are a different failure class (errorCode
// 'credits', "top up" message) — never misreport them as a transient capacity
// problem, because retrying cannot help. Providers phrase this in different
// ways (and languages):
//   Genspark : "Your Genspark credits have been exhausted…"
//   OpenAI   : "You exceeded your current quota, please check your plan and billing details"
//   Anthropic: "Your credit balance is too low"
//   DeepSeek : "Insufficient Balance"
//   MiniMax  : "已达到 Token Plan 用量上限：请升级 Token Plan 套餐或购买积分补充用量。"
//   Moonshot : "余额不足" / 通用中文网关: "额度不足" / "欠费" / "购买积分"
// The pattern matches both English and CJK phrasing so a Chinese 429 no
// longer falls through to the generic "busy, retry shortly" message.
const CREDITS_PATTERN = new RegExp(
  [
    'credit',
    'pricing',
    'insufficient.?quota',
    'quota.{0,12}(exceed|exhaust)',
    'exceed.{0,12}quota',
    'insufficient.?balance',
    'balance.{0,8}(too low|insufficient)',
    'billing',
    '充值',
    '余额不足',
    '额度不足',
    '用量上限',
    '欠费',
    '购买积分',
    '补充用量',
    '套餐',
  ].join('|'),
  'i',
)

function matches(text: string): boolean {
  return OVERLOADED_PATTERN.test(text) && !CREDITS_PATTERN.test(text)
}

/**
 * True when the error text names a quota / balance / credit exhaustion rather
 * than a transient capacity problem. Callers use this to surface a "top up"
 * message (errorCode 'credits') instead of "the service is busy, retry".
 *
 * Deliberately separate from `isAiOverloadedError` so the two are not
 * entangled: a 429 whose body says "rate_limit_error" but also carries a
 * quota notice is a quota problem, not a burst to ride out.
 */
export function isAiQuotaExhaustedError(err: unknown): boolean {
  let current: unknown = err
  for (let depth = 0; current && depth < 5; depth++) {
    if (typeof current === 'string') return CREDITS_PATTERN.test(current)
    const e = current as { message?: unknown; cause?: unknown }
    if (typeof e.message === 'string' && CREDITS_PATTERN.test(e.message)) return true
    current = e.cause
  }
  return false
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
