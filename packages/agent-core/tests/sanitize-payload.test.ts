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

  it('masks AWS access key ids', () => {
    const input = 'AWS_ACCESS_KEY_ID=AKIAIOSFODNN7EXAMPLE and ASIAJEXAMPLEXEG2JICE' // public-hygiene: fixture
    expect(sanitizeAgentPayload(input)).toBe(
      'AWS_ACCESS_KEY_ID=[REDACTED_API_KEY] and [REDACTED_API_KEY]',
    )
  })

  it('masks Slack tokens', () => {
    const bot = ['xoxb', '1234567890', '1234567890123', 'AbCdEfGhIjKlMnOpQrStUvWx'].join('-')
    const user = ['xoxp', '12345', '67890', 'abcdef'].join('-')
    const input = `bot ${bot} user ${user}`
    expect(sanitizeAgentPayload(input)).toBe('bot [REDACTED_API_KEY] user [REDACTED_API_KEY]')
  })

  it('masks PEM private key blocks, terminated or truncated', () => {
    const pem = [
      '-----BEGIN RSA PRIVATE KEY-----', // public-hygiene: fixture
      'MIIEowIBAAKCAQEA0Z3VS5JJcds3xfn/ygWyF8PbnGy0AIXdGXSp7pJnhn5ITVMr',
      '-----END RSA PRIVATE KEY-----',
    ].join('\n')
    expect(sanitizeAgentPayload(`key:\n${pem}\nrotate it`)).toBe(
      'key:\n[REDACTED_PRIVATE_KEY]\nrotate it',
    )
    const cut = [
      '-----BEGIN OPENSSH PRIVATE KEY-----', // public-hygiene: fixture
      'b3BlbnNzaC1rZXktdjEAAAAABG5vbmUAAAAEbm9uZQ==',
      'please rotate',
    ].join('\n')
    expect(sanitizeAgentPayload(cut)).toBe('[REDACTED_PRIVATE_KEY]\nplease rotate')
  })

  it('masks unquoted password assignments', () => {
    const input =
      'password=abc123 DB_PASSWORD: hunter2! passwd = s3cret-value db.password=p4ss.word'
    expect(sanitizeAgentPayload(input)).toBe(
      'password=[REDACTED_SECURE_TOKEN] DB_PASSWORD: [REDACTED_SECURE_TOKEN] passwd = [REDACTED_SECURE_TOKEN] db.password=[REDACTED_SECURE_TOKEN]',
    )
  })

  it('leaves prose after a password label untouched', () => {
    const input =
      'password: is stored in the vault, see /etc/passwd: root and password: (see attached)'
    expect(sanitizeAgentPayload(input)).toBe(input)
  })

  it('returns unrelated prose unchanged', () => {
    const input = 'Summarize the quarterly report and draft an email to the team.'
    expect(sanitizeAgentPayload(input)).toBe(input)
  })
})
