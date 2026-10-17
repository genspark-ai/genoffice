/**
 * Gmail send limits plus credential-hygiene gates: per-driver caps are
 * enforced before transmit and surfaced for compose, the OAuth scope is
 * gmail.modify only, and the driver source never mentions safeStorage,
 * refresh tokens, client secrets, or the full-mail scope.
 */

import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  GMAIL_MAX_MESSAGE_BYTES,
  GMAIL_MAX_RECIPIENTS_PER_MESSAGE,
  GMAIL_MODIFY_SCOPE,
  GMAIL_WARN_BYTES,
  GmailDriver,
  GmailSendLimitError,
  buildRawMessage,
  checkSendLimits,
  estimateSendBytes,
  type SendRequest,
} from '../src/gmail'
import { FakeGmailTransport, fakeAuth } from './gmail-fake'

const small = (over: Partial<SendRequest> = {}): SendRequest => ({
  to: ['bob@example.com'],
  subject: 'Hello',
  text: 'Small message',
  ...over,
})

const staged = (size: number) => ({
  filename: 'big.bin',
  mime: 'application/octet-stream',
  size,
  path: '/none',
})

describe('send-limit verdicts', () => {
  it('passes small sends', () => {
    expect(checkSendLimits(small())).toBe('ok')
    expect(estimateSendBytes(small())).toBeLessThan(GMAIL_WARN_BYTES)
  })

  it('warns above 25 MB without rejecting', () => {
    expect(checkSendLimits(small({ attachments: [staged(26 * 1024 * 1024)] }))).toBe('warn')
  })

  it('rejects above the 35 MB Gmail cap', () => {
    expect(checkSendLimits(small({ attachments: [staged(36 * 1024 * 1024)] }))).toBe('too-large')
  })

  it('rejects above the per-message recipient cap', () => {
    const to = Array.from(
      { length: GMAIL_MAX_RECIPIENTS_PER_MESSAGE + 1 },
      (_, i) => `u${i}@example.com`,
    )
    expect(checkSendLimits(small({ to }))).toBe('too-many-recipients')
  })
})

describe('driver send enforcement', () => {
  it('throws GmailSendLimitError before transmit on oversize sends', async () => {
    const transport = new FakeGmailTransport()
    const driver = new GmailDriver({ auth: fakeAuth(), transport, fromAddress: 'me@example.com' })
    await driver.connect()
    await expect(
      driver.send(small({ attachments: [staged(36 * 1024 * 1024)] })),
    ).rejects.toMatchObject({
      name: 'GmailSendLimitError',
      code: 'too-large',
    })
    expect(transport.sentRaw).toHaveLength(0)
    await driver.disconnect()
  })

  it('throws before transmit on too many recipients', async () => {
    const transport = new FakeGmailTransport()
    const driver = new GmailDriver({ auth: fakeAuth(), transport, fromAddress: 'me@example.com' })
    await driver.connect()
    const to = Array.from(
      { length: GMAIL_MAX_RECIPIENTS_PER_MESSAGE + 1 },
      (_, i) => `u${i}@example.com`,
    )
    const err = await driver.send(small({ to })).catch((e: unknown) => e)
    expect(err).toBeInstanceOf(GmailSendLimitError)
    expect((err as GmailSendLimitError).code).toBe('too-many-recipients')
    expect(transport.sentRaw).toHaveLength(0)
    await driver.disconnect()
  })

  it('transmits staged attachments from disk', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'gmail-stage-'))
    const path = join(dir, 'note.txt')
    writeFileSync(path, 'staged bytes')
    const transport = new FakeGmailTransport()
    const driver = new GmailDriver({ auth: fakeAuth(), transport, fromAddress: 'me@example.com' })
    await driver.connect()
    const result = await driver.send(
      small({ attachments: [{ filename: 'note.txt', mime: 'text/plain', size: 12, path }] }),
    )
    expect(result.outboxId).toMatch(/^gmail-outbox-/)
    expect(transport.sentRaw).toHaveLength(1)
    expect(GMAIL_MAX_MESSAGE_BYTES).toBe(35 * 1024 * 1024)
    await driver.disconnect()
  })

  it('builds multipart raw messages with attachments', () => {
    const raw = buildRawMessage({
      from: 'me@example.com',
      to: ['bob@example.com'],
      subject: 'Files',
      text: 'See attached',
      attachments: [
        { filename: 'a.txt', mime: 'text/plain', bytes: new TextEncoder().encode('data') },
      ],
    })
    expect(raw).toContain('multipart/mixed')
    expect(raw).toContain('filename="a.txt"')
  })
})

describe('credential hygiene', () => {
  it('requests gmail.modify and nothing broader', () => {
    expect(GMAIL_MODIFY_SCOPE).toBe('https://www.googleapis.com/auth/gmail.modify')
  })

  it('keeps secrets and full scopes out of the driver source', () => {
    const file = readFileSync(fileURLToPath(new URL('../src/gmail.ts', import.meta.url)), 'utf8')
    const code = file
      .replace(/\/\*[\s\S]*?\*\//g, '')
      .split('\n')
      .filter((line) => !line.trim().startsWith('//'))
      .join('\n')
    expect(code).not.toMatch(/safeStorage/)
    expect(code).not.toMatch(/refresh_token/)
    expect(code).not.toMatch(/client_secret/)
    expect(code).not.toMatch(/setUsePlainTextEncryption/)
    expect(code).not.toMatch(/https:\/\/mail\.google\.com\//)
  })
})
