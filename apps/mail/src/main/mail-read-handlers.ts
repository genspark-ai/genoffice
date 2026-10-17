import { matchesQuery } from '@genoffice/mail-protocols'
import type { MailEnvelope, MailFolder } from '@genoffice/mail-protocols'
import { MAIL_CHANNELS } from '../shared/ipc'
import type { MailAccountSummary, MessagePage } from '../shared/ipc'
import {
  mailDeleteSchema,
  mailFlagsChangeSchema,
  mailListMessagesSchema,
  mailMoveSchema,
  mailSearchLocalSchema,
  mailSearchServerSchema,
} from '../shared/ipc-schemas'
import {
  mailAddAccountSchema,
  mailFetchBodySchema,
  mailListFoldersSchema,
  mailRemoveAccountSchema,
} from '../shared/mail-read-schemas'
import type { ComposeIpcBus } from './compose-handlers'
import { buildMailFakeData } from './mail-fake-data'
import type { MailFakeBody } from './mail-fake-data'
import { getMailStore } from './mail-store'

// Shell-hosted read-path handlers (merge lane owns this file — Lane S/R/C
// files are consumed, never edited). Every handler validates with zod
// end-to-end (strict) before touching state; UI dev serves the fake demo
// mailbox ONLY (no real accounts, no creds, no sockets). Electron-free and
// vitest-safe: the shell main process passes an ipcMain-backed bus, tests
// inject a fake bus.
//
// Store bridge: seeded envelopes are upserted into the shared in-memory
// mail-store (`./mail-store`, Lane S owns) so later lanes can build on one
// stub; the handlers serve from the registry below (folders, bodies and
// server-extra have no store home yet).

interface MailRegistry {
  accounts: MailAccountSummary[]
  folders: Map<string, MailFolder[]>
  envelopes: MailEnvelope[]
  bodies: Map<string, MailFakeBody>
  serverExtra: MailEnvelope[]
}

let registry: MailRegistry | null = null
let handlersRegistered = false

function badRequest(message: string): Error {
  return new Error(`mail: ${message}`)
}

/** Seed once; later calls are no-ops (registration guards double-seed). */
export function ensureMailDemoSeed(): MailRegistry {
  if (registry) return registry
  const seed = buildMailFakeData()
  const next: MailRegistry = {
    accounts: [...seed.accounts],
    folders: new Map([[seed.accounts[0].accountId, [...seed.folders]]]),
    envelopes: [...seed.envelopes],
    bodies: new Map(Object.entries(seed.bodies)),
    serverExtra: [...seed.serverExtra],
  }
  getMailStore().upsertEnvelopes(next.envelopes)
  registry = next
  return next
}

/** Test-only reset (fresh seed on next register/ensure). */
export function resetMailDemoSeed(): void {
  registry = null
  handlersRegistered = false
}

function keyOf(folderId: string, uid: string): string {
  return `${folderId}/${uid}`
}

function findEnvelope(
  state: MailRegistry,
  folderId: string,
  uid: string,
): MailEnvelope | undefined {
  return state.envelopes.find((item) => item.ref.folderId === folderId && item.ref.uid === uid)
}

export function registerMailReadHandlers(bus: ComposeIpcBus): void {
  if (handlersRegistered) return
  handlersRegistered = true
  const state = ensureMailDemoSeed()

  bus.handle(MAIL_CHANNELS.listAccounts, async () => [...state.accounts])

  bus.handle(MAIL_CHANNELS.addAccount, async (payload) => {
    const parsed = mailAddAccountSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad add-account request')
    const account: MailAccountSummary = {
      accountId: `fake-${state.accounts.length + 1}`,
      label: parsed.data.label,
      sessionOnly: true,
    }
    state.accounts.push(account)
    state.folders.set(account.accountId, [])
    return account
  })

  bus.handle(MAIL_CHANNELS.removeAccount, async (payload) => {
    const parsed = mailRemoveAccountSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad remove-account request')
    state.accounts = state.accounts.filter((account) => account.accountId !== parsed.data.accountId)
    state.folders.delete(parsed.data.accountId)
    return undefined
  })

  bus.handle(MAIL_CHANNELS.listFolders, async (payload) => {
    const parsed = mailListFoldersSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad list-folders request')
    return [...(state.folders.get(parsed.data.accountId) ?? [])]
  })

  bus.handle(MAIL_CHANNELS.listMessages, async (payload) => {
    const parsed = mailListMessagesSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad list-messages request')
    const { folderId, limit, cursor } = parsed.data
    const start = cursor === undefined ? 0 : Number.parseInt(cursor, 10)
    if (!Number.isInteger(start) || start < 0) throw badRequest('bad list cursor')
    const inFolder = state.envelopes
      .filter((env) => env.ref.folderId === folderId)
      .sort((a, b) => (a.date < b.date ? 1 : -1))
    const page = inFolder.slice(start, start + limit)
    const next = start + limit < inFolder.length ? String(start + limit) : undefined
    const result: MessagePage = { envelopes: page, cursor: next }
    return result
  })

  bus.handle(MAIL_CHANNELS.fetchBody, async (payload) => {
    const parsed = mailFetchBodySchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad fetch-body request')
    const { ref } = parsed.data
    return state.bodies.get(keyOf(ref.folderId, ref.uid)) ?? { text: '' }
  })

  bus.handle(MAIL_CHANNELS.searchLocal, async (payload) => {
    const parsed = mailSearchLocalSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad local search request')
    void parsed.data.accountIds
    return state.envelopes.filter((env) => matchesQuery(env, parsed.data.query))
  })

  bus.handle(MAIL_CHANNELS.searchServer, async (payload) => {
    const parsed = mailSearchServerSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad server search request')
    const local = state.envelopes.filter((env) => matchesQuery(env, parsed.data.query))
    const remote = state.serverExtra.filter((env) => matchesQuery(env, parsed.data.query))
    return [...local, ...remote]
  })

  bus.handle(MAIL_CHANNELS.move, async (payload) => {
    const parsed = mailMoveSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad move request')
    const env = findEnvelope(state, parsed.data.ref.folderId, parsed.data.ref.uid)
    if (env) env.ref.folderId = parsed.data.toFolderId
    return undefined
  })

  bus.handle(MAIL_CHANNELS.flags, async (payload) => {
    const parsed = mailFlagsChangeSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad flags request')
    const env = findEnvelope(state, parsed.data.ref.folderId, parsed.data.ref.uid)
    if (env) env.flags = { ...env.flags, ...parsed.data.flags }
    return undefined
  })

  bus.handle(MAIL_CHANNELS.delete, async (payload) => {
    const parsed = mailDeleteSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad delete request')
    const { ref } = parsed.data
    state.envelopes = state.envelopes.filter(
      (item) => !(item.ref.folderId === ref.folderId && item.ref.uid === ref.uid),
    )
    return undefined
  })
}
