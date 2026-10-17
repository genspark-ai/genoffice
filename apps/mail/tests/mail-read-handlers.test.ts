import { beforeEach, describe, expect, it } from 'vitest'
import { MAIL_CHANNELS } from '../src/shared/ipc'
import type { ComposeIpcBus } from '../src/main/compose-handlers'
import { registerMailReadHandlers, resetMailDemoSeed } from '../src/main/mail-read-handlers'

/**
 * Merge-lane read-handler tests over a fake bus: zod validation
 * end-to-end, demo-mailbox seed (folders + messages on a fresh tab),
 * pagination cursors, body fetch with the fake-only html extension, local
 * vs server search, and move/flags/delete mutation. Fake addresses use
 * example.com only — never real mailboxes.
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

beforeEach(() => {
  resetMailDemoSeed()
})

describe('mail read handlers', () => {
  it('seeds the demo mailbox: accounts, folders, and inbox messages', async () => {
    const { bus, invoke } = makeBus()
    registerMailReadHandlers(bus)
    const accounts = (await invoke(MAIL_CHANNELS.listAccounts)) as Array<{
      accountId: string
    }>
    expect(accounts).toHaveLength(1)
    expect(accounts[0].accountId).toBe('demo')
    const folders = (await invoke(MAIL_CHANNELS.listFolders, {
      accountId: 'demo',
    })) as Array<{ id: string }>
    expect(folders.map((folder) => folder.id)).toEqual(['inbox', 'sent', 'drafts', 'archive'])
    const page = (await invoke(MAIL_CHANNELS.listMessages, {
      accountId: 'demo',
      folderId: 'inbox',
      limit: 50,
    })) as { envelopes: unknown[]; cursor?: string }
    expect(page.envelopes).toHaveLength(5)
    expect(page.cursor).toBeUndefined()
  })

  it('paginates date-desc with opaque cursors', async () => {
    const { bus, invoke } = makeBus()
    registerMailReadHandlers(bus)
    const first = (await invoke(MAIL_CHANNELS.listMessages, {
      accountId: 'demo',
      folderId: 'inbox',
      limit: 2,
    })) as { envelopes: Array<{ subject: string }>; cursor?: string }
    expect(first.envelopes).toHaveLength(2)
    expect(first.cursor).toBeDefined()
    const second = (await invoke(MAIL_CHANNELS.listMessages, {
      accountId: 'demo',
      folderId: 'inbox',
      limit: 2,
      cursor: first.cursor,
    })) as { envelopes: Array<{ subject: string }>; cursor?: string }
    expect(second.envelopes).toHaveLength(2)
    const subjects = [...first.envelopes, ...second.envelopes].map((env) => env.subject)
    expect(new Set(subjects).size).toBe(4)
  })

  it('fetches bodies with the fake-only html extension for the reader', async () => {
    const { bus, invoke } = makeBus()
    registerMailReadHandlers(bus)
    const body = (await invoke(MAIL_CHANNELS.fetchBody, {
      ref: { uid: 'm1', folderId: 'inbox', threadId: 'thread-m1' },
    })) as { text: string; html?: string }
    expect(body.text).toContain('INV-1042')
    expect(body.html).toContain('<b>INV-1042</b>')
    const missing = (await invoke(MAIL_CHANNELS.fetchBody, {
      ref: { uid: 'nope', folderId: 'inbox', threadId: 'thread-nope' },
    })) as { text: string }
    expect(missing.text).toBe('')
  })

  it('searches local hits and server-extra hits separately', async () => {
    const { bus, invoke } = makeBus()
    registerMailReadHandlers(bus)
    const local = (await invoke(MAIL_CHANNELS.searchLocal, {
      accountIds: ['demo'],
      query: { text: 'invoice' },
    })) as Array<{ ref: { uid: string }; subject: string }>
    const localUids = local.map((env) => env.ref.uid)
    expect(localUids).toContain('m1')
    expect(localUids).not.toContain('srv1')
    const server = (await invoke(MAIL_CHANNELS.searchServer, {
      accountId: 'demo',
      folderId: 'inbox',
      query: { text: 'invoice' },
    })) as Array<{ ref: { uid: string } }>
    const uids = server.map((env) => env.ref.uid)
    expect(uids).toContain('srv1')
    expect(uids.length).toBeGreaterThan(local.length)
  })

  it('mutates through move, flags, and delete', async () => {
    const { bus, invoke } = makeBus()
    registerMailReadHandlers(bus)
    await invoke(MAIL_CHANNELS.flags, {
      ref: { uid: 'm1', folderId: 'inbox', threadId: 'thread-m1' },
      flags: { seen: true },
    })
    await invoke(MAIL_CHANNELS.move, {
      ref: { uid: 'm1', folderId: 'inbox', threadId: 'thread-m1' },
      toFolderId: 'archive',
    })
    const archived = (await invoke(MAIL_CHANNELS.listMessages, {
      accountId: 'demo',
      folderId: 'archive',
      limit: 50,
    })) as {
      envelopes: Array<{ ref: { uid: string }; flags: { seen: boolean } }>
    }
    expect(archived.envelopes.map((env) => env.ref.uid)).toContain('m1')
    expect(archived.envelopes.find((env) => env.ref.uid === 'm1')?.flags.seen).toBe(true)
    await invoke(MAIL_CHANNELS.delete, {
      ref: { uid: 'm5', folderId: 'inbox', threadId: 'thread-m5' },
    })
    const inbox = (await invoke(MAIL_CHANNELS.listMessages, {
      accountId: 'demo',
      folderId: 'inbox',
      limit: 50,
    })) as { envelopes: Array<{ ref: { uid: string } }> }
    expect(inbox.envelopes.map((env) => env.ref.uid)).not.toContain('m5')
  })

  it('adds and removes accounts', async () => {
    const { bus, invoke } = makeBus()
    registerMailReadHandlers(bus)
    const created = (await invoke(MAIL_CHANNELS.addAccount, {
      label: 'Second',
    })) as { accountId: string; label: string }
    expect(created.label).toBe('Second')
    const folders = (await invoke(MAIL_CHANNELS.listFolders, {
      accountId: created.accountId,
    })) as unknown[]
    expect(folders).toEqual([])
    await invoke(MAIL_CHANNELS.removeAccount, {
      accountId: created.accountId,
    })
    const accounts = (await invoke(MAIL_CHANNELS.listAccounts)) as Array<{
      accountId: string
    }>
    expect(accounts.map((account) => account.accountId)).toEqual(['demo'])
  })

  it('rejects malformed payloads with mail errors', async () => {
    const { bus, invoke } = makeBus()
    registerMailReadHandlers(bus)
    await expect(
      invoke(MAIL_CHANNELS.listMessages, {
        accountId: 'demo',
        folderId: 'inbox',
        limit: 5000,
      }),
    ).rejects.toThrow('mail:')
    await expect(
      invoke(MAIL_CHANNELS.fetchBody, { ref: { uid: '', folderId: '' } }),
    ).rejects.toThrow('mail:')
    await expect(invoke(MAIL_CHANNELS.addAccount, { label: '' })).rejects.toThrow('mail:')
  })

  it('registers once: a second registration keeps the same seed', async () => {
    const { bus, invoke } = makeBus()
    registerMailReadHandlers(bus)
    await invoke(MAIL_CHANNELS.delete, {
      ref: { uid: 'm5', folderId: 'inbox', threadId: 'thread-m5' },
    })
    registerMailReadHandlers(bus)
    const inbox = (await invoke(MAIL_CHANNELS.listMessages, {
      accountId: 'demo',
      folderId: 'inbox',
      limit: 50,
    })) as { envelopes: Array<{ ref: { uid: string } }> }
    expect(inbox.envelopes.map((env) => env.ref.uid)).not.toContain('m5')
  })
})
