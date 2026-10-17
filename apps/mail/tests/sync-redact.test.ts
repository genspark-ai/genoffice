import { describe, expect, it } from 'vitest'
import { containsPii, sanitizeError } from '../src/main/sync/redact'

describe('PII redaction (spec case: pii-redaction)', () => {
  it('strips addresses, tokens, and long opaque runs from errors', () => {
    const dirty =
      'send failed for reader@example.com with bearer abcdefghij1234567890 and payload aGVsbG8td29ybGQtdGhpcy1pcy1hLWxvbmdib2JmaW5zdHJpbmdmb3J0ZXN0aW5nMTIz s\nsecond line body text here'
    const clean = sanitizeError(new Error(dirty))
    expect(clean).not.toContain('reader@example.com')
    expect(clean).not.toContain('\n')
    expect(containsPii(clean)).toBe(false)
  })

  it('falls back to a safe message for empty input', () => {
    expect(sanitizeError(undefined)).toBe('sync failed')
    expect(sanitizeError(new Error(''))).toBe('sync failed')
  })

  it('detects PII shapes and stays stable across calls', () => {
    expect(containsPii('nothing sensitive here')).toBe(false)
    expect(containsPii('mail reader@example.com now')).toBe(true)
    // Global-regex state must not flip answers on repeat calls.
    expect(containsPii('mail reader@example.com now')).toBe(true)
    expect(containsPii('nothing sensitive here')).toBe(false)
    expect(containsPii('nothing sensitive here')).toBe(false)
  })
})
