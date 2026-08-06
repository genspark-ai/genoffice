import { describe, expect, it } from 'vitest'
import { sanitizeAgentPayload } from '../src'

describe('sanitizeAgentPayload', () => {
  it('masks API keys with known secret prefixes', () => {
    const input = 'use sk-abcdef1234567890abcdef and ghp_0123456789abcdef0123456789abcdef0123'
    const out = sanitizeAgentPayload(input)
    expect(out).toBe('use [REDACTED_API_KEY] and [REDACTED_API_KEY]')
  })

  it('masks long project-scoped keys completely, leaving no tail', () => {
    const key = `sk-proj-${'A1b2C3d4'.repeat(18)}` // 152 chars, like real sk-proj-/sk-ant- keys
    const out = sanitizeAgentPayload(`key=${key} done`)
    expect(out).toBe('key=[REDACTED_API_KEY] done')
    expect(out).not.toContain('A1b2C3d4')
  })

  it('does not treat prefixes embedded in ordinary words as keys', () => {
    const input = 'the task-management-dashboard-redesign spec'
    expect(sanitizeAgentPayload(input)).toBe(input)
  })

  it('masks userinfo credentials in URLs', () => {
    const input = 'connect to postgres://admin:hunter2@db.example.com:5432/app'
    expect(sanitizeAgentPayload(input)).toBe(
      'connect to postgres://admin:[REDACTED_CREDENTIALS]@db.example.com:5432/app',
    )
  })

  it('leaves colon-at text that is not a URL untouched', () => {
    const input = 'meet at 10:30@office, mapping a:b@c stays as typed'
    expect(sanitizeAgentPayload(input)).toBe(input)
  })

  it('masks quoted password/key assignments while keeping the separator', () => {
    const input = `password: "hunter2" and PRIVATE_KEY = 'abc123'`
    expect(sanitizeAgentPayload(input)).toBe(
      `password: "[REDACTED_SECURE_TOKEN]" and PRIVATE_KEY = "[REDACTED_SECURE_TOKEN]"`,
    )
  })

  it('returns unrelated prose unchanged', () => {
    const input = 'Summarize the quarterly report and draft an email to the team.'
    expect(sanitizeAgentPayload(input)).toBe(input)
  })
})
