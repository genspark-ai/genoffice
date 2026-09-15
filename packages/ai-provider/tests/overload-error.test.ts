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

  it('does not match other HTTP errors or generic failures', () => {
    expect(isAiOverloadedError(new Error('Claude HTTP 401: bad key'))).toBe(false)
    expect(isAiOverloadedError(new Error('HTTP 500: internal error'))).toBe(false)
    expect(isAiOverloadedError(new Error('fetch failed cause=ECONNRESET'))).toBe(false)
    expect(isAiOverloadedError(new Error('The model returned no content'))).toBe(false)
    expect(isAiOverloadedError(null)).toBe(false)
    expect(isAiOverloadedError(undefined)).toBe(false)
  })
})

describe('isAiQuotaExhaustedError', () => {
  it('matches MiniMax 429 quota notice (Chinese) that also carries rate_limit_error', async () => {
    const { isAiQuotaExhaustedError } = await import('../src/overload-error')
    const body =
      '{"type":"error","error":{"type":"rate_limit_error","message":"已达到 Token Plan 用量上限：请升级 Token Plan 套餐或购买积分补充用量。 (2056)"},"http_code":"429"}'
    expect(isAiQuotaExhaustedError(new Error(body))).toBe(true)
    // …and the same body must NOT be classified as a transient capacity blip
    expect(isAiOverloadedError(new Error(body))).toBe(false)
  })

  it('matches English quota / balance notices across providers', async () => {
    const { isAiQuotaExhaustedError } = await import('../src/overload-error')
    expect(
      isAiQuotaExhaustedError(
        new Error(
          'HTTP 429: {"error":{"code":"insufficient_quota","message":"You exceeded your current quota, please check your plan and billing details"}}',
        ),
      ),
    ).toBe(true)
    expect(isAiQuotaExhaustedError(new Error('Your credit balance is too low'))).toBe(true)
    expect(isAiQuotaExhaustedError(new Error('DeepSeek HTTP 402: Insufficient Balance'))).toBe(true)
  })

  it('matches other CJK quota phrasing', async () => {
    const { isAiQuotaExhaustedError } = await import('../src/overload-error')
    expect(isAiQuotaExhaustedError(new Error('账户余额不足,请充值'))).toBe(true)
    expect(isAiQuotaExhaustedError(new Error('API 额度不足'))).toBe(true)
    expect(isAiQuotaExhaustedError(new Error('账户欠费'))).toBe(true)
  })

  it('walks the cause chain and accepts plain strings', async () => {
    const { isAiQuotaExhaustedError } = await import('../src/overload-error')
    expect(
      isAiQuotaExhaustedError(new Error('stream failed', { cause: new Error('余额不足') })),
    ).toBe(true)
    expect(isAiQuotaExhaustedError('用量上限')).toBe(true)
  })

  it('does not claim plain capacity errors', async () => {
    const { isAiQuotaExhaustedError } = await import('../src/overload-error')
    expect(isAiQuotaExhaustedError(new Error('HTTP 429: too many requests'))).toBe(false)
    expect(isAiQuotaExhaustedError(new Error('Overloaded'))).toBe(false)
    expect(isAiQuotaExhaustedError(new Error('HTTP 503: upstream unavailable'))).toBe(false)
    expect(isAiQuotaExhaustedError(null)).toBe(false)
  })
})
