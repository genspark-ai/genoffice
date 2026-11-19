/**
 * Classifies "this provider does not serve you here" refusals — HTTP 403s a
 * vendor issues from the request's origin (OpenAI's unsupported-territory
 * code, Gemini's location precondition, the terms-of-service wording OpenRouter
 * relays for Anthropic routes) — so the apps can say "pick another model or
 * change your network" (errorCode 'region') instead of a bare 403 dump. The
 * model is listed in the catalog because the vendor decides at call time.
 */
const REGION_BLOCKED_PATTERN = new RegExp(
  [
    'violation of provider terms of service',
    'unsupported_country_region_territory',
    'country, region, or territory not supported',
    'user location is not supported',
    'not (?:currently )?(?:available|supported) (?:in|for) your (?:country|region|location)',
  ].join('|'),
  'i',
)

/** True when the error (an Error, its `cause` chain, or an error string) is a region/territory refusal. */
export function isAiRegionBlockedError(err: unknown): boolean {
  let current: unknown = err
  for (let depth = 0; current && depth < 5; depth++) {
    if (typeof current === 'string') return REGION_BLOCKED_PATTERN.test(current)
    const e = current as { message?: unknown; cause?: unknown }
    if (typeof e.message === 'string' && REGION_BLOCKED_PATTERN.test(e.message)) return true
    current = e.cause
  }
  return false
}
