/**
 * Graph transport: narrow HTTP surface over Microsoft Graph plus the
 * injected-auth contract. The main process (apps/mail/src/main, locked
 * decision D6) owns the OAuth PKCE system-browser flow, the token cache,
 * refresh, and the credential vault (credential-vault.ts, the single
 * safeStorage caller per R2). This module receives only a callback that
 * resolves the current delegated access token and NEVER sees refresh
 * tokens, client secrets, or safeStorage. Least-privilege scopes
 * (delegated Mail.ReadWrite + offline_access + User.Read identity only,
 * NEVER app-only *.ReadWrite.All) are requested by the main-process OAuth
 * caller; GRAPH_MAIL_SCOPES pins that contract for review.
 */

import { Client } from '@microsoft/microsoft-graph-client'

/** Delegated scopes the main-process OAuth flow must request. Pinned for review. */
export const GRAPH_MAIL_SCOPES: readonly string[] = [
  'Mail.ReadWrite',
  'offline_access',
  'User.Read',
]

/**
 * Fresh delegated access token, resolved by the main-process vault caller
 * from its memory-only cache (refresh happens in main only). The driver
 * calls this per request batch; it must never persist or log the value.
 */
export interface GraphAuthProvider {
  getAccessToken(): Promise<string>
}

export interface GraphRequestInit {
  headers?: Record<string, string>
}

/** Minimal Graph surface the driver needs; fakes implement this in tests. */
export interface GraphTransport {
  get<T>(path: string, init?: GraphRequestInit): Promise<T>
  post<T>(path: string, body: unknown, init?: GraphRequestInit): Promise<T>
  patch<T>(path: string, body: unknown, init?: GraphRequestInit): Promise<T>
  del(path: string, init?: GraphRequestInit): Promise<void>
}

/** Normalized service failure. Status/code mirror the Graph error payload. */
export class GraphServiceError extends Error {
  readonly status: number
  readonly code: string
  constructor(status: number, code: string, message: string) {
    super(message)
    this.name = 'GraphServiceError'
    this.status = status
    this.code = code
  }
}

export function toGraphError(status: number, code: string, message: string): GraphServiceError {
  return new GraphServiceError(status, code, message)
}

/**
 * True when the server rejected a delta token: the caller must run a full
 * resync (decideSync with deltaExpired) instead of retrying the token.
 * Graph signals this with 410 Gone (resyncRequired) or an invalid-token
 * 400 (InvalidDeltaToken / syncStateNotFound).
 */
export function isDeltaExpiredError(err: unknown): boolean {
  if (!(err instanceof GraphServiceError)) return false
  if (err.status === 410) return true
  return (
    err.code === 'resyncRequired' ||
    err.code === 'InvalidDeltaToken' ||
    err.code === 'syncStateNotFound'
  )
}

/** True on 401/403 authentication failures (token rejected or expired). */
export function isAuthFailure(err: unknown): boolean {
  return err instanceof GraphServiceError && (err.status === 401 || err.status === 403)
}

/**
 * True when the identity layer reports invalid_grant: the refresh token is
 * dead and the main process must wipe the stored blob and require re-auth
 * (amended spec section 7.1). Surfaced so main can act; never handled here.
 */
export function isInvalidGrant(err: unknown): boolean {
  if (!(err instanceof GraphServiceError)) return false
  if (err.code.toLowerCase() === 'invalid_grant') return true
  return err.status === 401 && err.message.toLowerCase().includes('invalid_grant')
}

function normalizeClientError(err: unknown): GraphServiceError {
  if (err instanceof GraphServiceError) return err
  const rec = (err ?? {}) as {
    statusCode?: number
    status?: number
    code?: string
    message?: string
  }
  const status = rec.statusCode ?? rec.status ?? 0
  const code = rec.code ?? 'network'
  const message = rec.message ?? 'Graph request failed'
  return new GraphServiceError(status, code, message)
}

/**
 * Production transport over @microsoft/microsoft-graph-client (MIT).
 * Paths are relative ("/me/..."); paging cursors are absolute nextLink
 * URLs passed back by list calls. No token, body, or secret is logged.
 */
export function createGraphTransport(auth: GraphAuthProvider): GraphTransport {
  const client = Client.initWithMiddleware({ authProvider: auth })
  const request = async <T>(
    method: 'get' | 'post' | 'patch' | 'delete',
    path: string,
    body?: unknown,
    init?: GraphRequestInit,
  ): Promise<T> => {
    try {
      let req = client.api(path)
      if (init?.headers !== undefined) req = req.headers(init.headers)
      switch (method) {
        case 'get':
          return (await req.get()) as T
        case 'post':
          return (await req.post(body)) as T
        case 'patch':
          return (await req.patch(body)) as T
        case 'delete':
          await req.delete()
          return undefined as T
      }
    } catch (err) {
      throw normalizeClientError(err)
    }
  }
  return {
    get: <T>(path: string, init?: GraphRequestInit) => request<T>('get', path, undefined, init),
    post: <T>(path: string, body: unknown, init?: GraphRequestInit) =>
      request<T>('post', path, body, init),
    patch: <T>(path: string, body: unknown, init?: GraphRequestInit) =>
      request<T>('patch', path, body, init),
    del: async (path: string, init?: GraphRequestInit) => {
      await request<void>('delete', path, undefined, init)
    },
  }
}
