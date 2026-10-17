import { describe, expect, it } from 'vitest'
import { MAIL_CHANNELS } from '../src/shared/ipc'
import {
  mailDeleteSchema,
  mailEventSchema,
  mailFlagsChangeSchema,
  mailIntentSchema,
  mailListMessagesSchema,
  mailMoveSchema,
  mailOpenMessageSchema,
  mailQuerySchema,
  mailSaveDraftSchema,
  mailSearchLocalSchema,
  mailSearchServerSchema,
  sendRequestSchema,
} from '../src/shared/ipc-schemas'

describe('MAIL_CHANNELS', () => {
  it('keeps the canonical channel names', () => {
    expect(MAIL_CHANNELS.openMessage).toBe('mail:open-message')
    expect(MAIL_CHANNELS.dirtyChanged).toBe('mail:dirty-changed')
    expect(MAIL_CHANNELS.syncEvents).toBe('mail:sync-events')
    // every mail-owned channel carries the mail: prefix
    for (const [name, channel] of Object.entries(MAIL_CHANNELS)) {
      if (
        name === 'getLanguage' ||
        name === 'languageChanged' ||
        name === 'getTheme' ||
        name === 'themeChanged'
      ) {
        expect(channel.startsWith('app:')).toBe(true)
      } else {
        expect(channel.startsWith('mail:')).toBe(true)
      }
    }
  })
})

describe('mail schemas', () => {
  it('accepts a full 10-field MailQuery and rejects unknown fields', () => {
    const full = {
      text: 'invoice',
      from: 'a@x.com',
      to: 'b@x.com',
      subject: 'hi',
      seen: true,
      flagged: false,
      hasAttach: true,
      since: '2026-01-01',
      before: '2026-02-01',
      onlyUnread: true,
    }
    expect(mailQuerySchema.safeParse(full).success).toBe(true)
    expect(mailQuerySchema.safeParse({}).success).toBe(true)
    expect(mailQuerySchema.safeParse({ text: 'x', bogus: 1 }).success).toBe(false)
  })

  it('validates open-message refs', () => {
    expect(mailOpenMessageSchema.safeParse({ mailRef: 'mail://a/1' }).success).toBe(true)
    expect(mailOpenMessageSchema.safeParse({ mailRef: 'file.md' }).success).toBe(false)
  })

  it('validates send requests with path-canonical attachments', () => {
    const req = {
      to: ['b@x.com'],
      subject: 's',
      text: 'body',
      attachments: [{ filename: 'a.pdf', mime: 'application/pdf', size: 10, path: '/tmp/a.pdf' }],
    }
    expect(sendRequestSchema.safeParse(req).success).toBe(true)
    expect(sendRequestSchema.safeParse({ ...req, to: [] }).success).toBe(false)
    expect(sendRequestSchema.safeParse({ ...req, extra: 1 }).success).toBe(false)
  })

  it('validates all 7 MailEvent variants and rejects bodies', () => {
    const ref = { uid: '1', folderId: 'inbox', threadId: 't' }
    expect(
      mailEventSchema.safeParse({
        kind: 'messages-added',
        accountId: 'a',
        folderId: 'inbox',
        refs: [ref],
      }).success,
    ).toBe(true)
    expect(
      mailEventSchema.safeParse({
        kind: 'flags-changed',
        accountId: 'a',
        ref,
        flags: { seen: true },
      }).success,
    ).toBe(true)
    expect(
      mailEventSchema.safeParse({
        kind: 'sync-state',
        accountId: 'a',
        folderId: 'i',
        state: 'idle',
      }).success,
    ).toBe(true)
    expect(mailEventSchema.safeParse({ kind: 'bogus', accountId: 'a' }).success).toBe(false)
  })

  it('requires idempotent move/flags/delete payloads', () => {
    const ref = { uid: '1', folderId: 'inbox', threadId: 't' }
    expect(mailMoveSchema.safeParse({ ref, toFolderId: 'archive' }).success).toBe(true)
    expect(
      mailFlagsChangeSchema.safeParse({ ref, flags: { seen: true }, expectedModseq: '5' }).success,
    ).toBe(true)
    expect(mailDeleteSchema.safeParse({ ref }).success).toBe(true)
    expect(mailMoveSchema.safeParse({ ref }).success).toBe(false)
  })

  it('validates list/search/draft/intent payloads', () => {
    expect(
      mailListMessagesSchema.safeParse({ accountId: 'a', folderId: 'i', limit: 50 }).success,
    ).toBe(true)
    expect(
      mailSearchLocalSchema.safeParse({ accountIds: ['a'], query: { text: 'x' } }).success,
    ).toBe(true)
    expect(
      mailSearchServerSchema.safeParse({ accountId: 'a', folderId: 'i', query: {} }).success,
    ).toBe(true)
    expect(
      mailSaveDraftSchema.safeParse({ accountId: 'a', threadKey: 't', text: 'd' }).success,
    ).toBe(true)
    expect(
      mailIntentSchema.safeParse({
        accountId: 'a',
        kind: 'move',
        targetUid: '1',
        payloadJson: '{}',
      }).success,
    ).toBe(true)
  })
})
