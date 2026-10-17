import { existsSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { stageAttachment } from '@genoffice/mail-protocols'
import type { SendRequest } from '@genoffice/mail-protocols'
import { MAIL_CHANNELS } from '../src/shared/ipc'
import { COMPOSE_CHANNELS } from '../src/shared/compose-channels'
import { FakeComposeProvider, registerComposeHandlers } from '../src/main/compose-handlers'
import type { ComposeHandlerDeps, ComposeIpcBus } from '../src/main/compose-handlers'

/**
 * Lane C handler tests over a fake bus: zod validation end-to-end, fake
 * provider sends, path-canonical staging with data-fallback, outbox
 * intent-recording, and the safeExternalUrl mailto gate. Fake addresses
 * use example.com only — never real mailboxes.
 */

function makeBus(): {
  bus: ComposeIpcBus
  invoke: (channel: string, payload: unknown) => Promise<unknown>
} {
  const handlers = new Map<string, (payload: unknown) => Promise<unknown>>()
  return {
    bus: {
      handle: (channel, listener) => {
        handlers.set(channel, listener)
      },
    },
    invoke: async (channel, payload) => {
      const listener = handlers.get(channel)
      if (!listener) throw new Error(`no handler: ${channel}`)
      return listener(payload)
    },
  }
}

function makeDeps(over?: Partial<ComposeHandlerDeps>): {
  deps: ComposeHandlerDeps
  provider: FakeComposeProvider
  opened: string[]
  drafts: Map<string, string>
  stageDir: string
} {
  const provider = new FakeComposeProvider()
  const opened: string[] = []
  const drafts = new Map<string, string>()
  const stageDir = mkdtempSync(join(tmpdir(), 'mail-compose-test-'))
  const deps: ComposeHandlerDeps = {
    provider,
    stage: stageAttachment,
    stageDir,
    openExternal: (url: string) => {
      opened.push(url)
    },
    saveDraft: (accountId, threadKey, text) => {
      drafts.set(`${accountId}/${threadKey}`, text)
    },
    clock: () => '2026-10-15T00:00:00.000Z',
    ...over,
  }
  return { deps, provider, opened, drafts, stageDir }
}

function sendReq(over?: Partial<SendRequest>): SendRequest {
  return { to: ['a@example.com'], subject: 'hi', text: 'hello', ...over }
}

describe('compose-send handler', () => {
  it('sends through the fake provider and records the outbox intent', async () => {
    const { bus, invoke } = makeBus()
    const { deps, provider, stageDir } = makeDeps()
    registerComposeHandlers(bus, deps)
    const staged = stageAttachment({
      filename: 'note.txt',
      mime: 'text/plain',
      bytes: new Uint8Array([104, 105]),
      stageDir,
    })
    const { filename, mime, size, path } = staged
    const result = (await invoke(MAIL_CHANNELS.composeSend, {
      ...sendReq(),
      attachments: [{ filename, mime, size, path }],
    })) as { outboxId: string }
    expect(result.outboxId).toBe('compose-outbox-1')
    expect(provider.sent).toHaveLength(1)
    const status = (await invoke(MAIL_CHANNELS.outboxStatus, { accountId: 'acc-1' })) as {
      queued: number
      failed: number
    }
    expect(status).toEqual({ queued: 1, failed: 0 })
    const list = (await invoke(COMPOSE_CHANNELS.outboxList, { accountId: 'acc-1' })) as Array<{
      id: string
      state: string
      subject: string
    }>
    expect(list).toHaveLength(1)
    expect(list[0]).toMatchObject({ id: 'compose-outbox-1', state: 'queued', subject: 'hi' })
  })

  it('streams staged files from disk and strips memory payloads', async () => {
    const { bus, invoke } = makeBus()
    const { deps, provider, stageDir } = makeDeps()
    registerComposeHandlers(bus, deps)
    const staged = stageAttachment({
      filename: 'disk.txt',
      mime: 'text/plain',
      bytes: new Uint8Array([1, 2, 3, 4]),
      stageDir,
    })
    await invoke(MAIL_CHANNELS.composeSend, {
      ...sendReq(),
      attachments: [{ ...staged, data: new Uint8Array([9, 9, 9, 9]) }],
    })
    expect(provider.sent).toHaveLength(1)
    const got = provider.sent[0].attachments?.[0]
    expect(got?.path).toBe(staged.path)
    expect(got).not.toHaveProperty('data')
    expect(existsSync(staged.path)).toBe(true)
  })

  it('re-stages data-fallback attachments missing a disk path', async () => {
    const { bus, invoke } = makeBus()
    const { deps, provider, stageDir } = makeDeps()
    registerComposeHandlers(bus, deps)
    await invoke(MAIL_CHANNELS.composeSend, {
      ...sendReq(),
      attachments: [
        {
          filename: 'fallback.bin',
          mime: 'application/octet-stream',
          size: 3,
          path: 'memory:fallback.bin',
          data: new Uint8Array([7, 8, 9]),
        },
      ],
    })
    expect(provider.sent).toHaveLength(1)
    const got = provider.sent[0].attachments?.[0]
    expect(got?.path.startsWith(stageDir)).toBe(true)
    expect(existsSync(got?.path ?? '')).toBe(true)
    expect(got).not.toHaveProperty('data')
  })

  it('rejects renderer-supplied staged paths outside the stage dir (M1)', async () => {
    const { bus, invoke } = makeBus()
    const { deps, provider } = makeDeps()
    registerComposeHandlers(bus, deps)
    const outside = mkdtempSync(join(tmpdir(), 'mail-compose-outside-'))
    const victim = join(outside, 'secret.txt')
    writeFileSync(victim, 'topsecret')
    await expect(
      invoke(MAIL_CHANNELS.composeSend, {
        ...sendReq(),
        attachments: [{ filename: 'secret.txt', mime: 'text/plain', size: 9, path: victim }],
      }),
    ).rejects.toThrow('mail: staged attachment outside stage dir')
    expect(provider.sent).toHaveLength(0)
  })

  it('rejects invalid sends without calling the provider', async () => {
    const { bus, invoke } = makeBus()
    const { deps, provider } = makeDeps()
    registerComposeHandlers(bus, deps)
    await expect(
      invoke(MAIL_CHANNELS.composeSend, { to: [], subject: 'x', text: 'y' }),
    ).rejects.toThrow('mail: bad send request')
    await expect(invoke(MAIL_CHANNELS.composeSend, { ...sendReq(), bogus: 1 })).rejects.toThrow(
      'mail: bad send request',
    )
    expect(provider.sent).toHaveLength(0)
    const status = (await invoke(MAIL_CHANNELS.outboxStatus, { accountId: 'acc-1' })) as {
      queued: number
      failed: number
    }
    expect(status).toEqual({ queued: 0, failed: 0 })
  })

  it('enforces total and per-driver size caps', async () => {
    const { bus, invoke } = makeBus()
    const { deps, provider } = makeDeps({
      sendCaps: { warnTotalBytes: 8, maxTotalBytes: 16, maxAttachmentBytes: 8, maxRecipients: 100 },
    })
    registerComposeHandlers(bus, deps)
    await expect(
      invoke(MAIL_CHANNELS.composeSend, sendReq({ text: 'x'.repeat(32) })),
    ).rejects.toThrow('mail: attachments exceed the size cap')
    await expect(
      invoke(MAIL_CHANNELS.composeSend, {
        ...sendReq({ text: 'hi' }),
        attachments: [
          {
            filename: 'big.bin',
            mime: 'application/octet-stream',
            size: 12,
            path: 'memory:big.bin',
            data: new Uint8Array(12),
          },
        ],
      }),
    ).rejects.toThrow('mail: attachment exceeds the per-account cap')
    expect(provider.sent).toHaveLength(0)
  })

  it('records failed sends in the outbox ledger', async () => {
    const { bus, invoke } = makeBus()
    const { deps, provider } = makeDeps()
    provider.failSendsWith = 'fake smtp down'
    registerComposeHandlers(bus, deps)
    await expect(invoke(MAIL_CHANNELS.composeSend, sendReq())).rejects.toThrow(
      'mail: queued send failed',
    )
    const status = (await invoke(MAIL_CHANNELS.outboxStatus, { accountId: 'acc-1' })) as {
      queued: number
      failed: number
    }
    expect(status).toEqual({ queued: 0, failed: 1 })
    const list = (await invoke(COMPOSE_CHANNELS.outboxList, { accountId: 'acc-1' })) as Array<{
      state: string
      error?: string
    }>
    expect(list[0]).toMatchObject({ state: 'failed', error: 'fake smtp down' })
  })
})

describe('draft + staging + external handlers', () => {
  it('saves drafts through the injected store', async () => {
    const { bus, invoke } = makeBus()
    const { deps, drafts } = makeDeps()
    registerComposeHandlers(bus, deps)
    await invoke(MAIL_CHANNELS.saveDraft, {
      accountId: 'acc-1',
      threadKey: 't-1',
      text: '{"hello":1}',
    })
    expect(drafts.get('acc-1/t-1')).toBe('{"hello":1}')
    await expect(
      invoke(MAIL_CHANNELS.saveDraft, { accountId: 'acc-1', threadKey: 't-1' }),
    ).rejects.toThrow('mail: bad draft')
  })

  it('stages attachments to disk and rejects oversize input', async () => {
    const { bus, invoke } = makeBus()
    const { deps } = makeDeps()
    registerComposeHandlers(bus, deps)
    const staged = (await invoke(COMPOSE_CHANNELS.stageAttachment, {
      filename: 'a.txt',
      mime: 'text/plain',
      data: new Uint8Array([1, 2, 3]),
    })) as { filename: string; size: number; path: string }
    expect(staged.size).toBe(3)
    expect(existsSync(staged.path)).toBe(true)
    await expect(
      invoke(COMPOSE_CHANNELS.stageAttachment, {
        filename: 'huge.bin',
        mime: 'application/octet-stream',
        data: new Uint8Array(51 * 1024 * 1024),
      }),
    ).rejects.toThrow('mail: attachments exceed the size cap')
  })

  it('opens only allowlisted external URLs', async () => {
    const { bus, invoke } = makeBus()
    const { deps, opened } = makeDeps()
    registerComposeHandlers(bus, deps)
    const mailto = (await invoke(COMPOSE_CHANNELS.openExternal, {
      url: 'mailto:a@example.com?subject=hi',
    })) as { opened: boolean }
    expect(mailto).toEqual({ opened: true })
    await invoke(COMPOSE_CHANNELS.openExternal, { url: '  https://example.com/x  ' })
    expect(opened).toEqual(['mailto:a@example.com?subject=hi', 'https://example.com/x'])
    await expect(
      invoke(COMPOSE_CHANNELS.openExternal, { url: 'javascript:alert(1)' }),
    ).rejects.toThrow('mail: blocked external url')
    await expect(
      invoke(COMPOSE_CHANNELS.openExternal, { url: 'file:///etc/passwd' }),
    ).rejects.toThrow('mail: blocked external url')
    expect(opened).toHaveLength(2)
  })
})
