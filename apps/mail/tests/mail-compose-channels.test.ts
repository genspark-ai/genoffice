import { describe, expect, it } from 'vitest'
import { MAIL_CHANNELS } from '../src/shared/ipc'
import { COMPOSE_CHANNELS } from '../src/shared/compose-channels'
import {
  composeOpenExternalSchema,
  composeOutboxListSchema,
  composeStageAttachmentSchema,
} from '../src/shared/compose-schemas'

/**
 * Lane C compose-session channels: names stay namespaced and disjoint from
 * the canonical table, and every schema is strict (unknown fields reject).
 * The open-external schema validates shape only — the javascript:/file:
 * gate is `safeExternalUrl` in the handler (covered by handler tests).
 */

describe('COMPOSE_CHANNELS', () => {
  it('keeps compose-only namespaced names disjoint from the canonical table', () => {
    expect(COMPOSE_CHANNELS.stageAttachment).toBe('mail:compose-stage-attachment')
    expect(COMPOSE_CHANNELS.outboxList).toBe('mail:compose-outbox-list')
    expect(COMPOSE_CHANNELS.openExternal).toBe('mail:compose-open-external')
    const canonical: Set<string> = new Set(Object.values(MAIL_CHANNELS))
    for (const channel of Object.values(COMPOSE_CHANNELS)) {
      expect(channel.startsWith('mail:')).toBe(true)
      expect(canonical.has(channel)).toBe(false)
    }
  })
})

describe('compose schemas', () => {
  it('stages attachments strictly', () => {
    const valid = { filename: 'a.txt', mime: 'text/plain', data: new Uint8Array([1, 2, 3]) }
    expect(composeStageAttachmentSchema.safeParse(valid).success).toBe(true)
    expect(
      composeStageAttachmentSchema.safeParse({ filename: '', mime: 'text/plain', data: valid.data })
        .success,
    ).toBe(false)
    expect(composeStageAttachmentSchema.safeParse({ ...valid, bogus: 1 }).success).toBe(false)
  })

  it('lists outbox intents per account strictly', () => {
    expect(composeOutboxListSchema.safeParse({ accountId: 'acc-1' }).success).toBe(true)
    expect(composeOutboxListSchema.safeParse({ accountId: '' }).success).toBe(false)
    expect(composeOutboxListSchema.safeParse({ accountId: 'acc-1', bogus: 1 }).success).toBe(false)
  })

  it('validates external-open shape (gating happens in the handler)', () => {
    expect(composeOpenExternalSchema.safeParse({ url: 'mailto:a@example.com' }).success).toBe(true)
    expect(composeOpenExternalSchema.safeParse({ url: 'https://example.com/x' }).success).toBe(true)
    expect(composeOpenExternalSchema.safeParse({ url: '' }).success).toBe(false)
    expect(composeOpenExternalSchema.safeParse({ url: 'x'.repeat(2049) }).success).toBe(false)
    expect(
      composeOpenExternalSchema.safeParse({ url: 'https://example.com', bogus: 1 }).success,
    ).toBe(false)
  })
})
