import { describe, expect, it } from 'vitest'
import { isAiOverloadedError } from '../src/overload-error'

describe('isAiOverloadedError', () => {
  it('matches the Genspark gateway 429 engine-overloaded body', () => {
    expect(
      isAiOverloadedError(
        new Error(
          'HTTP 429: {"error":{"message":"The engine is currently overloaded, please try again later","type":"engine_overloaded_error"}}',
        ),
      ),
    ).toBe(true)
  })

  it('matches provider-prefixed capacity/rate-limit HTTP statuses', () => {
    expect(isAiOverloadedError(new Error('Claude HTTP 529: overloaded_error'))).toBe(true)
    expect(isAiOverloadedError(new Error('Gemini HTTP 429: resource exhausted'))).toBe(true)
    expect(isAiOverloadedError(new Error('HTTP 503: upstream temporarily unavailable'))).toBe(true)
  })

  it('matches in-stream error notices that carry no HTTP status', () => {
    expect(isAiOverloadedError(new Error('Overloaded'))).toBe(true)
    expect(isAiOverloadedError(new Error('Rate limit exceeded, retry in 30s'))).toBe(true)
    expect(isAiOverloadedError(new Error('Resource has been exhausted (check quota).'))).toBe(true)
    expect(isAiOverloadedError('Too Many Requests')).toBe(true)
  })

  it('walks the cause chain', () => {
    const cause = new Error('HTTP 429: too many requests')
    expect(isAiOverloadedError(new Error('stream failed', { cause }))).toBe(true)
  })

  it('never claims credits-exhausted notices (a separate error class)', () => {
    expect(
      isAiOverloadedError(
        new Error('Your Genspark credits have been exhausted. Visit genspark.ai/pricing to top up'),
      ),
    ).toBe(false)
  })

  it('still exempts a credits-exhausted body that carries an HTTP status', () => {
    // credits are only typed on the HTTP 200 JSON path, so this arrives on a non-2xx
    // as a plain body — classifying it as "service busy, retry" is wrong advice for a
    // billing notice, so the exhausted wording has to win over the status marker.
    expect(
      isAiOverloadedError(
        new Error('HTTP 429: {"error":{"message":"Your Genspark credits have been exhausted"}}'),
      ),
    ).toBe(false)
    expect(
      isAiOverloadedError(new Error('HTTP 503: insufficient credits, top up to continue')),
    ).toBe(false)
  })

  it('keeps a rate limit transient when it only names the credits bucket', () => {
    // "credits per minute" is a rate limit, not an exhausted balance: the status
    // marker is what lifts the credits exemption, and main already declined the
    // no-status variant (a bare credits mention stays out of the retry path).
    expect(
      isAiOverloadedError(new Error('HTTP 429: rate limit: credits per minute exceeded')),
    ).toBe(true)
    expect(isAiOverloadedError(new Error('rate limit reached for your credit tier'))).toBe(false)
  })

  it('does not match other HTTP errors or generic failures', () => {
    expect(isAiOverloadedError(new Error('Claude HTTP 401: bad key'))).toBe(false)
    expect(isAiOverloadedError(new Error('HTTP 500: internal error'))).toBe(false)
    expect(isAiOverloadedError(new Error('fetch failed cause=ECONNRESET'))).toBe(false)
    expect(isAiOverloadedError(new Error('The model returned no content'))).toBe(false)
    expect(isAiOverloadedError(null)).toBe(false)
    expect(isAiOverloadedError(undefined)).toBe(false)
  })
})
