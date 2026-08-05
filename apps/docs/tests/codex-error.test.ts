import { describe, expect, it } from 'vitest'
import { safeCodexStreamError } from '../src/main/codex-error'

describe('safeCodexStreamError', () => {
  it.each([
    [401, 'ChatGPT authorization was rejected (HTTP 401). Sign in again.'],
    [403, 'ChatGPT access was denied (HTTP 403). Check your ChatGPT access.'],
    [404, 'ChatGPT Codex endpoint or selected model is unavailable (HTTP 404).'],
    [429, 'ChatGPT Codex rate limit reached (HTTP 429). Try again later.'],
    [503, 'ChatGPT Codex service error (HTTP 503). Try again later.'],
  ])('categorizes HTTP %s without exposing its body', (status, expected) => {
    const error = new Error(`Codex HTTP ${status}: bearer secret-token and backend body`)

    expect(safeCodexStreamError(error)).toBe(expected)
    expect(safeCodexStreamError(error)).not.toContain('secret-token')
  })

  it('categorizes timeout, network, and protocol failures without raw details', () => {
    expect(safeCodexStreamError(new Error('request timed out at https://secret.example'))).toBe(
      'ChatGPT Codex request timed out. Try again.',
    )
    expect(safeCodexStreamError(new Error('fetch failed: ENOTFOUND private-host'))).toBe(
      'ChatGPT Codex network error. Check your connection and try again.',
    )
    expect(
      safeCodexStreamError(new Error('Codex stream malformed event: {"token":"secret"}')),
    ).toBe('ChatGPT Codex returned an invalid stream response. Try again.')
    expect(
      safeCodexStreamError(new Error('model gpt-private is not available for this account')),
    ).toBe('ChatGPT Codex selected model is unavailable. Choose another model and try again.')
  })

  it('surfaces only a bounded safe 400 code and detail', () => {
    const error = Object.assign(new Error('Codex HTTP 400'), {
      name: 'CodexHttpError',
      status: 400,
      code: 'invalid_model',
      detail: 'The selected model is unavailable.',
    })

    expect(safeCodexStreamError(error)).toBe(
      'ChatGPT Codex rejected the request (HTTP 400; invalid_model: The selected model is unavailable.).',
    )
  })

  it('uses a generic safe message for unknown errors', () => {
    expect(safeCodexStreamError(new Error('unexpected token secret'))).toBe(
      'ChatGPT Codex request failed. Try again.',
    )
  })
})
