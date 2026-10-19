import { describe, expect, it } from 'vitest'
import { isAiRegionBlockedError } from '../src/region-error'

describe('isAiRegionBlockedError', () => {
  it('matches the terms-of-service 403 OpenRouter relays for Anthropic routes', () => {
    expect(
      isAiRegionBlockedError(
        'HTTP 403: {"error":{"message":"The request is prohibited due to a violation of provider Terms Of Service.","code":403}}',
      ),
    ).toBe(true)
  })

  it('matches vendor territory codes and location preconditions', () => {
    expect(isAiRegionBlockedError(new Error('403 unsupported_country_region_territory'))).toBe(true)
    expect(isAiRegionBlockedError('User location is not supported for the API use.')).toBe(true)
    expect(isAiRegionBlockedError('This model is not available in your region')).toBe(true)
  })

  it('walks the cause chain', () => {
    const err = new Error('request failed', {
      cause: new Error('Country, region, or territory not supported'),
    })
    expect(isAiRegionBlockedError(err)).toBe(true)
  })

  it('leaves other 403s and unrelated errors alone', () => {
    expect(isAiRegionBlockedError('HTTP 403: invalid api key')).toBe(false)
    expect(isAiRegionBlockedError('HTTP 429: rate limit exceeded')).toBe(false)
    expect(isAiRegionBlockedError(new Error('fetch failed'))).toBe(false)
    expect(isAiRegionBlockedError(undefined)).toBe(false)
  })
})
