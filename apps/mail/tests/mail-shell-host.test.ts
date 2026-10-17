import { describe, expect, it } from 'vitest'
import { MAIL_CHANNELS } from '../src/shared/ipc'
import type { ComposeIpcBus } from '../src/main/compose-handlers'
import { registerMailShellHandlers } from '../src/main/mail-shell-host'
import { resetMailDemoSeed } from '../src/main/mail-read-handlers'

/**
 * Merge-lane shell-host test: one `registerMailShellHandlers` call wires
 * BOTH the read path (demo mailbox) and the compose path (fake send,
 * draft save, outbox status) — the shell-hosted equivalent of QA flows
 * 1-4. Fakes only, never real accounts.
 */

function makeBus(): {
  bus: ComposeIpcBus
  invoke: (channel: string, payload?: unknown) => Promise<unknown>
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

describe('mail shell host', () => {
  it('serves reads and compose from one registration', async () => {
    resetMailDemoSeed()
    const { bus, invoke } = makeBus()
    const opened: string[] = []
    registerMailShellHandlers(bus, {
      openExternal: (url) => {
        opened.push(url)
      },
    })

    // Flow 1-2: fresh tab sees the demo mailbox and reads a message.
    const accounts = (await invoke(MAIL_CHANNELS.listAccounts)) as Array<{
      accountId: string
    }>
    expect(accounts.map((account) => account.accountId)).toEqual(['demo'])
    const page = (await invoke(MAIL_CHANNELS.listMessages, {
      accountId: 'demo',
      folderId: 'inbox',
      limit: 50,
    })) as { envelopes: Array<{ ref: { uid: string } }> }
    expect(page.envelopes.length).toBeGreaterThan(0)
    const body = (await invoke(MAIL_CHANNELS.fetchBody, {
      ref: { uid: 'm1', folderId: 'inbox', threadId: 'thread-m1' },
    })) as { text: string }
    expect(body.text).toContain('INV-1042')

    // Flow 3: reply sends through the fake provider into the outbox stub.
    const sent = (await invoke(MAIL_CHANNELS.composeSend, {
      to: ['billing@example.com'],
      subject: 'Re: Your invoice is ready',
      text: 'Paid — receipt appreciated.',
    })) as { outboxId: string }
    expect(sent.outboxId).toBeDefined()
    const status = (await invoke(MAIL_CHANNELS.outboxStatus, {
      accountId: 'demo',
    })) as { queued: number; failed: number }
    expect(status).toEqual({ queued: 1, failed: 0 })

    // Flow 4: draft save persists without error.
    await invoke(MAIL_CHANNELS.saveDraft, {
      accountId: 'demo',
      threadKey: 'thread-m1',
      text: '{"to":"","cc":"","bcc":"","subject":"Re: invoice","bodyHtml":"draft"}',
    })
    expect(opened).toEqual([])
  })
})
