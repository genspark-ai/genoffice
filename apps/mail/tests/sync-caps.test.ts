import { describe, expect, it } from 'vitest'
import {
  capsForKind,
  checkOutboxCaps,
  countOutboxRecipients,
  estimateOutboxBytes,
} from '../src/main/sync/caps'
import { testSendRequest } from './sync-fixtures'

describe('outbox cap enforcement (spec cases: outbox-cap-enforcement)', () => {
  it('locks per-driver caps', () => {
    expect(capsForKind('imap')).toMatchObject({
      maxTotalBytes: 50 * 1024 * 1024,
      maxAttachmentBytes: 25 * 1024 * 1024,
      maxRecipients: 100,
    })
    expect(capsForKind('gmail')).toMatchObject({
      maxTotalBytes: 35 * 1024 * 1024,
      maxRecipients: 500,
    })
    expect(capsForKind('graph')).toMatchObject({
      maxTotalBytes: 4 * 1024 * 1024,
      maxAttachmentBytes: 3 * 1024 * 1024,
      maxRecipients: 500,
    })
  })

  it('accepts a small send on every driver', () => {
    const req = testSendRequest()
    for (const kind of ['imap', 'gmail', 'graph'] as const) {
      expect(checkOutboxCaps(kind, req).verdict).toBe('ok')
    }
  })

  it('rejects too many recipients with per-driver messages', () => {
    const many = Array.from({ length: 101 }, (_, i) => `r${i}@example.com`)
    expect(checkOutboxCaps('imap', testSendRequest({ to: many })).verdict).toBe(
      'too-many-recipients',
    )
    // Gmail and Graph allow 500: 101 recipients pass.
    expect(checkOutboxCaps('gmail', testSendRequest({ to: many })).verdict).toBe('ok')
    expect(checkOutboxCaps('graph', testSendRequest({ to: many })).verdict).toBe('ok')
    const over = Array.from({ length: 501 }, (_, i) => `r${i}@example.com`)
    expect(checkOutboxCaps('gmail', testSendRequest({ to: over })).verdict).toBe(
      'too-many-recipients',
    )
    expect(checkOutboxCaps('graph', testSendRequest({ to: over })).verdict).toBe(
      'too-many-recipients',
    )
  })

  it('rejects oversize bodies per driver (IMAP 50 MB, Gmail 35 MB, Graph 4 MB)', () => {
    const bytes36mb = 36 * 1024 * 1024
    // Estimated bytes bypass the text alloc: pass the estimate directly.
    const req = testSendRequest()
    // IMAP warns above 25 MB but still queues; Gmail/Graph reject.
    expect(checkOutboxCaps('imap', req, bytes36mb).verdict).toBe('warn')
    expect(checkOutboxCaps('imap', req, 1024).verdict).toBe('ok')
    expect(checkOutboxCaps('gmail', req, bytes36mb).verdict).toBe('too-large')
    expect(checkOutboxCaps('graph', req, bytes36mb).verdict).toBe('too-large')
    const bytes51mb = 51 * 1024 * 1024
    expect(checkOutboxCaps('imap', req, bytes51mb).verdict).toBe('too-large')
  })

  it('rejects oversize attachments per driver', () => {
    const big = {
      filename: 'video.bin',
      mime: 'application/octet-stream',
      size: 4 * 1024 * 1024,
      path: '/staging/video.bin',
    }
    const req = testSendRequest({ attachments: [big] })
    expect(checkOutboxCaps('graph', req).verdict).toBe('attachment-too-large')
    expect(checkOutboxCaps('imap', req).verdict).toBe('ok')
    expect(checkOutboxCaps('gmail', req).verdict).toBe('ok')
  })

  it('warns above the warn threshold without rejecting', () => {
    const req = testSendRequest()
    const warnBytes = 26 * 1024 * 1024
    expect(checkOutboxCaps('imap', req, warnBytes).verdict).toBe('warn')
  })

  it('estimates encoded bytes deterministically', () => {
    const req = testSendRequest({ text: 'hello', html: '<p>hello</p>' })
    expect(estimateOutboxBytes(req)).toBe(estimateOutboxBytes(req))
    expect(countOutboxRecipients(testSendRequest({ to: ['a@x.com'], cc: ['b@x.com'] }))).toBe(2)
  })
})
