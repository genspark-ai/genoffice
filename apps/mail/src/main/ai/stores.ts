import type { MailEnvelope } from '@genoffice/mail-protocols'
import { ensureMailDemoSeed } from '../mail-read-handlers'
import { getMailStore } from '../mail-store'
import type { AiRetrieveDeps, AiRetrieveStorePort } from './retrieve'

/**
 * Production store wiring for mail:ai-retrieve (P5 spec section 4).
 * Known accounts come from the demo seed registry; each scope member
 * resolves to its real per-account SQLite store when one is configured,
 * and otherwise to the demo registry slice for that account. Existing
 * fetch-body / search-local handlers are untouched; this adapter only
 * reads through the same public accessors.
 */

type DemoRegistry = ReturnType<typeof ensureMailDemoSeed>

function includes(haystack: string, needle: string): boolean {
  return haystack.toLowerCase().includes(needle.toLowerCase())
}

function registryPort(accountId: string, state: DemoRegistry): AiRetrieveStorePort {
  return {
    search(_self, query, opts) {
      const needle = query.trim().toLowerCase()
      if (!needle) return []
      return state.envelopes
        .filter((env) => opts.folderId === undefined || env.ref.folderId === opts.folderId)
        .filter((env) =>
          includes([...env.from, ...env.to, env.subject, env.snippet].join(' '), needle),
        )
        .slice(0, opts.limit)
    },
    body(_self, folderId, uid) {
      return state.bodies.get(`${folderId}/${uid}`)?.text
    },
    envelope(_self, folderId, uid) {
      return state.envelopes.find((env) => env.ref.folderId === folderId && env.ref.uid === uid)
    },
  }
}

interface SqliteLike {
  searchLocal(query: string, opts?: { limit?: number }): MailEnvelope[]
  getMessageBody(accountId: string, folderId: string, uid: string): { text: string } | undefined
  getEnvelope(accountId: string, folderId: string, uid: string): MailEnvelope | undefined
}

function isSqliteLike(store: unknown): store is SqliteLike {
  if (!store || typeof store !== 'object') return false
  const candidate = store as Record<string, unknown>
  return (
    typeof candidate['searchLocal'] === 'function' &&
    typeof candidate['getMessageBody'] === 'function' &&
    typeof candidate['getEnvelope'] === 'function'
  )
}

function sqlitePort(accountId: string, store: SqliteLike): AiRetrieveStorePort {
  return {
    search(_self, query, opts) {
      const hits = store.searchLocal(query, { limit: opts.limit })
      return opts.folderId === undefined
        ? hits
        : hits.filter((env) => env.ref.folderId === opts.folderId)
    },
    body(_self, folderId, uid) {
      return store.getMessageBody(accountId, folderId, uid)?.text
    },
    envelope(_self, folderId, uid) {
      return store.getEnvelope(accountId, folderId, uid)
    },
  }
}

/** Production deps: demo-seed accounts with per-account SQLite when configured. */
export function defaultAiRetrieveDeps(): AiRetrieveDeps {
  return {
    knownAccountIds: () => ensureMailDemoSeed().accounts.map((account) => account.accountId),
    storeFor: (accountId: string) => {
      const state = ensureMailDemoSeed()
      const configured = getMailStore(accountId)
      if (isSqliteLike(configured)) return sqlitePort(accountId, configured)
      return registryPort(accountId, state)
    },
  }
}
