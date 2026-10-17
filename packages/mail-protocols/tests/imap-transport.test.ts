/**
 * Production-transport wiring checks that need no network:
 * - NodemailerSmtpAdapter builds real MIME through nodemailer's offline
 *   stream transport (memory + disk-path attachments).
 * - ImapFlowAdapter constructs without connecting and reports capabilities.
 */
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { unlinkSync, writeFileSync } from 'node:fs'
import nodemailer from 'nodemailer'
import { describe, expect, it } from 'vitest'
import { ImapFlowAdapter, NodemailerSmtpAdapter } from '../src/imap-transport.js'

const IMAP_OPTIONS = {
  host: 'imap.example.local',
  port: 993,
  secure: true,
  auth: { user: 'test-user', password: 'test-pass' },
}

const SMTP_OPTIONS = {
  host: 'smtp.example.local',
  port: 465,
  secure: true,
  auth: { user: 'test-user', password: 'test-pass' },
  from: 'test-user@example.local',
}

describe('NodemailerSmtpAdapter (offline stream transport)', () => {
  it('renders MIME with memory and disk-path attachments', async () => {
    const staged = join(tmpdir(), `mail-lane-b-transport-${Date.now()}.txt`)
    writeFileSync(staged, 'staged attachment bytes')
    try {
      const stream = nodemailer.createTransport({ streamTransport: true, buffer: true })
      const adapter = new NodemailerSmtpAdapter(SMTP_OPTIONS, stream)
      const result = await adapter.send({
        from: SMTP_OPTIONS.from,
        to: ['friend@example.local'],
        subject: 'Streamed',
        text: 'offline MIME render',
        attachments: [
          {
            filename: 'note.txt',
            contentType: 'text/plain',
            content: new TextEncoder().encode('hi'),
          },
          { filename: 'staged.txt', contentType: 'text/plain', path: staged },
        ],
      })
      expect(result.messageId).toContain('@')
      await adapter.close()
    } finally {
      unlinkSync(staged)
    }
  })
})

describe('ImapFlowAdapter (no connection)', () => {
  it('constructs and answers capability queries without connecting', () => {
    const adapter = new ImapFlowAdapter(IMAP_OPTIONS)
    expect(typeof adapter.supports('IDLE')).toBe('boolean')
  })
})
