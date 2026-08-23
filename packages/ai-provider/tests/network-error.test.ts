import { describe, expect, it } from 'vitest'
import { isAiNetworkError } from '../src/network-error'

describe('isAiNetworkError', () => {
  it('matches the wrapped provider fetch failure with the cause code in the message', () => {
    expect(isAiNetworkError(new Error('Claude fetch failed: fetch failed cause=ECONNRESET'))).toBe(
      true,
    )
    expect(
      isAiNetworkError(new Error('Claude fetch failed: fetch failed cause=ECONNREFUSED')),
    ).toBe(true)
  })

  it('matches a raw undici fetch TypeError via the cause chain', () => {
    const cause = Object.assign(new Error('connect ECONNREFUSED 127.0.0.1:443'), {
      code: 'ECONNREFUSED',
    })
    expect(isAiNetworkError(new TypeError('fetch failed', { cause }))).toBe(true)
  })

  it('matches Chromium net-stack errors from the net.fetch rescue path', () => {
    expect(isAiNetworkError(new Error('net::ERR_INTERNET_DISCONNECTED'))).toBe(true)
    expect(isAiNetworkError(new Error('net::ERR_PROXY_CONNECTION_FAILED'))).toBe(true)
    expect(isAiNetworkError(new Error('net::ERR_CONNECTION_RESET'))).toBe(true)
    expect(isAiNetworkError(new Error('net::ERR_NAME_NOT_RESOLVED'))).toBe(true)
  })

  it('matches DNS and unreachable-host failures', () => {
    expect(isAiNetworkError(new Error('getaddrinfo ENOTFOUND api.anthropic.com'))).toBe(true)
    expect(isAiNetworkError(new Error('getaddrinfo EAI_AGAIN llm.genspark.ai'))).toBe(true)
    expect(isAiNetworkError(new Error('connect ENETUNREACH 1.2.3.4:443'))).toBe(true)
  })

  it('does not match HTTP, credits, or content errors', () => {
    expect(isAiNetworkError(new Error('Claude HTTP 500: internal error'))).toBe(false)
    expect(isAiNetworkError(new Error('Your Genspark credits have been exhausted.'))).toBe(false)
    expect(isAiNetworkError(new Error('Claude returned no content (empty stream)'))).toBe(false)
    expect(isAiNetworkError(new Error('A custom provider requires a Base URL'))).toBe(false)
    expect(isAiNetworkError(null)).toBe(false)
    expect(isAiNetworkError(undefined)).toBe(false)
  })
})
