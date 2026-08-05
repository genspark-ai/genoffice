/**
 * GenOffice's own Genspark identity: device-code login (office_addin_auth,
 * app_type=genoffice) minting a gsk API key named "genoffice" — the key_name
 * lands in billing as billing_tag, attributing all traffic (incl. gsk CLI
 * subprocesses) to GenOffice. Stored in ~/.genoffice/auth.json, deliberately
 * NOT the shared config.json that Claw Desktop overwrites on every launch.
 *
 * Flow: POST /device_code → browser approve → poll /token for a 30-day Bearer
 * → POST /session for a cookie → POST /api/api_tokens/create. The Bearer is
 * kept to rebuild a session later (key revoke on logout).
 */

import { existsSync, mkdirSync, readFileSync, unlinkSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { homedir } from 'node:os'
import { dirname, join } from 'node:path'
import { asRecord } from './shared'

/** Progress event for the browser login flow. */
export interface GskLoginProgress {
  phase: 'url' | 'success' | 'error'
  url?: string
  expiresInSec?: number
  /** 'network' | 'expired' | raw error text */
  error?: string
}

const APP_TYPE = 'genoffice'
const KEY_NAME = 'genoffice'
const HTTP_TIMEOUT_MS = 30_000

function baseUrl(): string {
  return (process.env.GSK_BASE_URL || 'https://www.genspark.ai').replace(/\/$/, '')
}

let cachedFetch: typeof fetch | undefined

/**
 * Cloudflare bot-challenges /api/office_addin_auth for non-browser clients;
 * Chromium's net.fetch passes (and exposes getSetCookie). Falls back to
 * global fetch outside the Electron main process (tests). Note net.fetch is
 * not spec fetch: an explicit Cookie header is forwarded, not stripped
 * (verified against a header-echo server).
 */
function resolveFetch(): typeof fetch {
  if (cachedFetch) return cachedFetch
  let impl: typeof fetch = (...args) => fetch(...args)
  if (process.versions.electron) {
    try {
      const require = createRequire(import.meta.url)
      const { net } = require('electron') as { net?: { fetch?: typeof fetch } }
      if (net?.fetch) impl = net.fetch.bind(net)
    } catch {
      /* non-main context: keep global fetch */
    }
  }
  cachedFetch = impl
  return impl
}

/** Override dir via GENOFFICE_AUTH_DIR (test isolation). */
export function genofficeAuthPath(): string {
  return join(process.env.GENOFFICE_AUTH_DIR || join(homedir(), '.genoffice'), 'auth.json')
}

export interface GenofficeAuth {
  apiKey: string
  keyId?: string
  accessToken?: string
}

let cachedAuth: GenofficeAuth | null | undefined

function readAuthFile(): GenofficeAuth | null {
  try {
    const raw = asRecord(JSON.parse(readFileSync(genofficeAuthPath(), 'utf-8')))
    if (typeof raw.api_key !== 'string' || !raw.api_key) return null
    return {
      apiKey: raw.api_key,
      ...(typeof raw.key_id === 'string' && raw.key_id ? { keyId: raw.key_id } : {}),
      ...(typeof raw.access_token === 'string' && raw.access_token
        ? { accessToken: raw.access_token }
        : {}),
    }
  } catch {
    return null
  }
}

export function loadGenofficeAuth(): GenofficeAuth | null {
  if (cachedAuth === undefined) cachedAuth = readAuthFile()
  return cachedAuth
}

/** The GenOffice-named api key; '' when not signed in. Cached (invalidated by login/logout). */
export function genofficeApiKey(): string {
  return loadGenofficeAuth()?.apiKey ?? ''
}

function saveAuth(auth: GenofficeAuth): void {
  const path = genofficeAuthPath()
  mkdirSync(dirname(path), { recursive: true })
  writeFileSync(
    path,
    JSON.stringify(
      { api_key: auth.apiKey, key_id: auth.keyId, access_token: auth.accessToken },
      null,
      2,
    ) + '\n',
    { mode: 0o600 },
  )
  cachedAuth = auth
}

function clearAuth(): void {
  try {
    if (existsSync(genofficeAuthPath())) unlinkSync(genofficeAuthPath())
  } catch {
    /* local sign-out must not throw */
  }
  cachedAuth = null
}

// ── HTTP helpers ─────────────────────────────────────────────────────

class LoginFlowError extends Error {
  constructor(
    /** matches GskLoginProgress.error conventions */
    public reason: 'network' | 'expired' | string,
  ) {
    super(reason)
  }
}

async function httpJson(
  url: string,
  init: RequestInit & { signal: AbortSignal },
): Promise<{ resp: Response; json: Record<string, unknown> }> {
  let resp: Response
  try {
    resp = await resolveFetch()(url, init)
  } catch (e) {
    if (init.signal.aborted) throw e
    throw new LoginFlowError('network')
  }
  let json: Record<string, unknown> = {}
  try {
    json = asRecord(await resp.json())
  } catch {
    /* non-JSON body; status checks below carry the error */
  }
  return { resp, json }
}

function sessionCookie(resp: Response): string {
  const headers = resp.headers as Headers & { getSetCookie?: () => string[] }
  const cookies = headers.getSetCookie?.() ?? []
  if (cookies.length === 0) {
    const single = resp.headers.get('set-cookie')
    if (single) cookies.push(single)
  }
  return cookies.map((c) => c.split(';')[0]!).join('; ')
}

function sleep(ms: number, signal: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    const onAbort = () => {
      clearTimeout(timer)
      reject(new DOMException('login cancelled', 'AbortError'))
    }
    // remove on resolve — the poll loop reuses one signal and would accumulate listeners
    const timer = setTimeout(() => {
      signal.removeEventListener('abort', onAbort)
      resolve()
    }, ms)
    signal.addEventListener('abort', onAbort, { once: true })
  })
}

// ── Device-code login ────────────────────────────────────────────────

async function revokeKey(cookie: string, keyId: string, signal: AbortSignal): Promise<void> {
  await resolveFetch()(`${baseUrl()}/api/api_tokens/revoke`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Cookie: cookie },
    body: JSON.stringify({ key_id: keyId }),
    signal,
  })
}

/** Bearer token → session cookie. Shared by login (key create) and logout (key revoke). */
async function establishSession(accessToken: string, signal: AbortSignal): Promise<string> {
  const { resp } = await httpJson(`${baseUrl()}/api/office_addin_auth/session`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${accessToken}` },
    signal,
  })
  const cookie = sessionCookie(resp)
  if (!resp.ok || !cookie) throw new LoginFlowError(`session failed (HTTP ${resp.status})`)
  return cookie
}

async function runDeviceLogin(
  signal: AbortSignal,
  emit: (progress: GskLoginProgress) => void,
): Promise<void> {
  const { resp, json } = await httpJson(
    `${baseUrl()}/api/office_addin_auth/device_code?app_type=${APP_TYPE}`,
    { method: 'POST', signal },
  )
  const code = String(json.device_code ?? '')
  const authUrl = String(json.auth_url ?? '')
  if (!resp.ok || !code || !authUrl) throw new LoginFlowError('network')
  const expiresInSec = Number(json.expires_in) > 0 ? Number(json.expires_in) : 600
  const pollMs = Number(json.poll_interval) > 0 ? Number(json.poll_interval) * 1000 : 2000
  emit({ phase: 'url', url: authUrl, expiresInSec })

  const deadline = Date.now() + expiresInSec * 1000
  let accessToken: string
  for (;;) {
    if (Date.now() >= deadline) throw new LoginFlowError('expired')
    await sleep(pollMs, signal)
    const { json: poll } = await httpJson(`${baseUrl()}/api/office_addin_auth/token?code=${code}`, {
      method: 'GET',
      signal,
    })
    const status = String(poll.status ?? '')
    if (status === 'pending') continue
    if (status === 'approved') {
      accessToken = String(poll.access_token ?? '')
      if (!accessToken) throw new LoginFlowError('expired')
      break
    }
    throw new LoginFlowError('expired')
  }

  const cookie = await establishSession(accessToken, signal)
  const { json: created } = await httpJson(`${baseUrl()}/api/api_tokens/create`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Cookie: cookie },
    body: JSON.stringify({ key_name: KEY_NAME }),
    signal,
  })
  const data = asRecord(created.data)
  const apiKey = typeof data.token === 'string' ? data.token : ''
  if (!apiKey) {
    throw new LoginFlowError(String(created.message ?? 'API key creation failed'))
  }
  const previousKeyId = loadGenofficeAuth()?.keyId
  saveAuth({
    apiKey,
    ...(typeof data.key_id === 'string' ? { keyId: data.key_id } : {}),
    accessToken,
  })
  emit({ phase: 'success' })
  // re-login replaces the key: revoke the superseded one so it can't pile up
  // server-side (best-effort — the new login must not fail on this)
  if (previousKeyId && previousKeyId !== data.key_id) {
    try {
      await revokeKey(cookie, previousKeyId, signal)
    } catch {
      /* ignore */
    }
  }
}

let activeLogin: { cancel: () => void } | null = null

/**
 * Starts the GenOffice device-code login, cancelling a previous in-flight one
 * (its device code would otherwise be approved into a dead flow). The caller
 * opens `url` in the system browser. Returns whether the flow was started.
 */
export function startGenofficeLogin(onEvent?: (progress: GskLoginProgress) => void): boolean {
  activeLogin?.cancel()
  const emit = onEvent ?? (() => {})
  const controller = new AbortController()
  let done = false
  const self = {
    cancel: () => {
      done = true
      controller.abort()
    },
  }
  activeLogin = self
  const finish = (progress: GskLoginProgress) => {
    if (done) return
    done = true
    if (activeLogin === self) activeLogin = null
    emit(progress)
  }
  void runDeviceLogin(controller.signal, (progress) => {
    if (done) return
    if (progress.phase === 'url') emit(progress)
    else finish(progress)
  }).catch((e: unknown) => {
    if (controller.signal.aborted) return
    finish({
      phase: 'error',
      error: e instanceof LoginFlowError ? e.reason : String((e as Error)?.message ?? e),
    })
  })
  return true
}

/** True while a login started via startGenofficeLogin is in flight. */
export function genofficeLoginInFlight(): boolean {
  return activeLogin !== null
}

/**
 * Fire-and-forget login for entry points without progress UI. Reuses an
 * in-flight flow (restarting would strand it on a dead device code); openUrl
 * is the caller's browser opener (this module is Electron-free).
 */
export function ensureGenofficeLogin(openUrl: (url: string) => void): void {
  if (genofficeLoginInFlight()) return
  startGenofficeLogin((progress) => {
    if (progress.url) openUrl(progress.url)
  })
}

/**
 * Signs out of GenOffice only: best-effort server-side revoke of the
 * genoffice key, then local removal. The shared gsk CLI login
 * (~/.genspark-tool-cli) is untouched — terminal gsk and Claw keep working.
 */
export async function genofficeLogout(): Promise<void> {
  const auth = loadGenofficeAuth()
  if (auth?.accessToken && auth.keyId) {
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), HTTP_TIMEOUT_MS)
    try {
      const cookie = await establishSession(auth.accessToken, controller.signal)
      await revokeKey(cookie, auth.keyId, controller.signal)
    } catch {
      /* revoke is best-effort; the local key removal below is the contract */
    } finally {
      clearTimeout(timer)
    }
  }
  clearAuth()
}

/** Test hook: drop the auth cache so the next read hits the file system. */
export function resetGenofficeAuthCache(): void {
  cachedAuth = undefined
}
