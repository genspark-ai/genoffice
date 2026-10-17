/**
 * Encrypted credential store for GenOffice Mail (H2 hardening lane).
 *
 * The ONLY module in the mail app that may touch Electron safeStorage.
 * OAuth refresh tokens (Gmail, Graph) and IMAP/SMTP app passwords are
 * stored as ONE opaque encrypted blob per account; access tokens stay
 * memory-only and are never persisted.
 *
 * Electron-free by construction: safeStorage is an injected dependency
 * (typed only), so this module stays vitest-safe. Only async safeStorage
 * calls are used (encryptStringAsync / decryptStringAsync /
 * isAsyncEncryptionAvailable / getSelectedStorageBackend); every public
 * operation requires the app-ready gate first.
 *
 * Linux-no-keyring posture (spec C-5): a `basic_text` backend, an
 * unavailable encryption API, or a forced session-only launch enters
 * SESSION-ONLY mode for that launch. Secrets stay in memory and are
 * never written; the caller surfaces NO_KEYRING_BANNER. Persisting a
 * basic_text blob as if protected is refused unless the caller passes
 * the explicit opt-in, which must be labeled UNPROTECTED_LABEL.
 */

import { mkdirSync, rmSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { Buffer } from 'node:buffer'

/** Account providers the store accepts secrets for. */
export type CredentialProvider = 'gmail' | 'graph' | 'imap'

/** Least-privilege Gmail OAuth scopes: modify plus identity only, never full mail. */
export const GMAIL_OAUTH_SCOPES: readonly string[] = Object.freeze([
  'openid',
  'email',
  'profile',
  'https://www.googleapis.com/auth/gmail.modify',
])

/** Least-privilege Graph OAuth scopes: delegated only, never app-only. */
export const GRAPH_OAUTH_SCOPES: readonly string[] = Object.freeze([
  'Mail.ReadWrite',
  'offline_access',
  'User.Read',
])

/** Scopes an IMAP/SMTP account may record: none, passwords carry no OAuth scope. */
export const IMAP_OAUTH_SCOPES: readonly string[] = Object.freeze([])

/** Requested scopes per provider. Requesting anything broader is refused. */
export function scopesForProvider(provider: CredentialProvider): readonly string[] {
  if (provider === 'gmail') return GMAIL_OAUTH_SCOPES
  if (provider === 'graph') return GRAPH_OAUTH_SCOPES
  return IMAP_OAUTH_SCOPES
}

/** Persistent non-dismissible banner text for the no-keyring launch mode. */
export const NO_KEYRING_BANNER = 'No OS keyring detected — accounts will need sign-in each launch.'

/** Label the explicit unprotected-persistence opt-in must carry in UI. */
export const UNPROTECTED_LABEL = 'unprotected (obfuscated only)'

/** Storage backend posture resolved from the Electron backend probe. */
export type BackendPosture = 'protected-keyring' | 'session-only'

/** Minimal async-only safeStorage surface this module is allowed to use.
 * Widened to Uint8Array so fakes stay simple; the real Electron
 * safeStorage (Buffer-based) remains structurally assignable. */
export interface SafeStorageLike {
  encryptStringAsync(plainText: string): Promise<Uint8Array>
  decryptStringAsync(encrypted: Uint8Array): Promise<string>
  isAsyncEncryptionAvailable(): Promise<boolean>
  getSelectedStorageBackend(): Promise<string>
}

/** Compile-time pin: the real Electron safeStorage stays assignable here. */
export function connectSafeStorage(
  vault: Pick<
    import('electron').SafeStorage,
    | 'encryptStringAsync'
    | 'decryptStringAsync'
    | 'isAsyncEncryptionAvailable'
    | 'getSelectedStorageBackend'
  >,
): SafeStorageLike {
  return vault as unknown as SafeStorageLike
}

/** One opaque encrypted blob per account: plaintext metadata plus blob. */
export interface StoredCredentialRecord {
  accountId: string
  provider: CredentialProvider
  address: string
  scopesGranted: string[]
  blob: string
  backend: string
  protection: 'os-keyring' | 'obfuscated-only'
  sessionOnly: boolean
  updatedAt: string
}

/** Safe account summary for logs and IPC: never carries secrets or blobs. */
export interface CredentialSummary {
  accountId: string
  provider: CredentialProvider
  address: string
  scopesGranted: string[]
  backend: string
  protection: 'os-keyring' | 'obfuscated-only'
  sessionOnly: boolean
  hasSecret: boolean
  updatedAt: string
}

export interface PutCredentialInput {
  accountId: string
  provider: CredentialProvider
  address: string
  scopesGranted: string[]
  /** OAuth refresh token or app password. Synthetic fixtures only in tests. */
  secret: string
  /** Per-account session-only toggle (C-10): memory only when true. */
  sessionOnly?: boolean
}

export interface CredentialStoreOptions {
  rootDir: string
  safeStorage: SafeStorageLike
  /** App-ready gate (C-2): operations run only after the app is ready. */
  isAppReady: () => boolean
  clock?: () => string
  /**
   * Headless/CI launches are always session-only with no fallback files.
   * The main process sets this when no OS keyring can exist.
   */
  forceSessionOnly?: boolean
  /**
   * Explicit opt-in persistence without a keyring. Default OFF; the UI
   * must label it UNPROTECTED_LABEL when this is true.
   */
  allowUnprotectedPersistence?: boolean
}

function fail(message: string): never {
  throw new Error(message)
}

/** Thrown when a stored blob cannot be decrypted: wipe and re-auth. */
export class CredentialReauthRequiredError extends Error {
  readonly accountId: string
  constructor(accountId: string) {
    super('mail: stored credential unreadable, sign-in required')
    this.name = 'CredentialReauthRequiredError'
    this.accountId = accountId
  }
}

/** Thrown when stored data fails integrity checks before any crypto use. */
export class CredentialCorruptError extends Error {
  readonly accountId: string
  constructor(accountId: string) {
    super('mail: stored credential invalid, sign-in required')
    this.name = 'CredentialCorruptError'
    this.accountId = accountId
  }
}

/** Returned secret payload: plaintext only in memory, never logged. */
export interface ResolvedSecret {
  secret: string
  sessionOnly: boolean
}

const ACCOUNT_ID_MAX = 256
const ADDRESS_MAX = 320
const SECRET_MAX = 16_384

function sanitizeFileStem(accountId: string): string {
  const safe = accountId.replace(/[^A-Za-z0-9._-]/g, '_').slice(0, 128)
  return safe.length > 0 ? safe : 'account'
}

/**
 * A `basic_text` backend (or an unavailable encryption API) means no OS
 * keyring: SESSION-ONLY for that launch. Every other named backend is a
 * platform keyring (Electron only reports `basic_text` for the
 * unprotected case), so anything else counts as protected.
 */
export function postureForBackend(backend: string, encryptionAvailable: boolean): BackendPosture {
  if (!encryptionAvailable) return 'session-only'
  const name = backend.trim().toLowerCase()
  if (name === '' || name === 'basic_text' || name === 'basic') return 'session-only'
  return 'protected-keyring'
}

function isValidRecord(value: unknown): value is StoredCredentialRecord {
  if (typeof value !== 'object' || value === null) return false
  const record = value as Record<string, unknown>
  return (
    typeof record['accountId'] === 'string' &&
    (record['provider'] === 'gmail' ||
      record['provider'] === 'graph' ||
      record['provider'] === 'imap') &&
    typeof record['address'] === 'string' &&
    Array.isArray(record['scopesGranted']) &&
    typeof record['blob'] === 'string' &&
    typeof record['backend'] === 'string' &&
    (record['protection'] === 'os-keyring' || record['protection'] === 'obfuscated-only') &&
    typeof record['sessionOnly'] === 'boolean' &&
    typeof record['updatedAt'] === 'string'
  )
}

export class CredentialStore {
  private readonly dir: string
  private readonly vault: SafeStorageLike
  private readonly isAppReady: () => boolean
  private readonly clock: () => string
  private readonly forceSessionOnly: boolean
  private readonly allowUnprotectedPersistence: boolean
  private readonly secrets = new Map<string, { secret: string; sessionOnly: boolean }>()
  private readonly metas = new Map<
    string,
    {
      provider: CredentialProvider
      address: string
      scopesGranted: string[]
      backend: string
      protection: 'os-keyring' | 'obfuscated-only'
      updatedAt: string
    }
  >()
  private readonly accessTokens = new Map<string, { token: string; expiresAt: number }>()
  private readonly sessionFlags = new Map<string, boolean>()
  private backendCache: { backend: string; available: boolean } | null = null

  constructor(opts: CredentialStoreOptions) {
    if (opts.rootDir === '') fail('mail: credential store needs a root dir')
    this.dir = join(opts.rootDir, 'credentials')
    this.vault = opts.safeStorage
    this.isAppReady = opts.isAppReady
    this.clock = opts.clock ?? (() => new Date().toISOString())
    this.forceSessionOnly = opts.forceSessionOnly ?? false
    this.allowUnprotectedPersistence = opts.allowUnprotectedPersistence ?? false
  }

  private ensureReady(): void {
    if (!this.isAppReady()) fail('mail: credential store used before app ready')
  }

  private recordPath(accountId: string): string {
    return join(this.dir, `${sanitizeFileStem(accountId)}.cred.json`)
  }

  private async probeBackend(): Promise<{ backend: string; available: boolean }> {
    if (!this.backendCache) {
      const [backend, available] = await Promise.all([
        this.vault.getSelectedStorageBackend(),
        this.vault.isAsyncEncryptionAvailable(),
      ])
      this.backendCache = { backend, available }
    }
    return this.backendCache
  }

  /** Re-probe on next call; used when the keyring state may have changed. */
  refreshBackend(): void {
    this.backendCache = null
  }

  /** Current backend posture; session-only forces the no-keyring mode. */
  async backendPosture(): Promise<BackendPosture> {
    this.ensureReady()
    if (this.forceSessionOnly) return 'session-only'
    const { backend, available } = await this.probeBackend()
    return postureForBackend(backend, available)
  }

  /** True when this launch must show the no-keyring sign-in banner. */
  async requiresSignInEachLaunch(): Promise<boolean> {
    return (await this.backendPosture()) === 'session-only'
  }

  private validateInput(input: PutCredentialInput): void {
    if (input.accountId === '' || input.accountId.length > ACCOUNT_ID_MAX) {
      fail('mail: bad credential account id')
    }
    if (
      input.address === '' ||
      input.address.length > ADDRESS_MAX ||
      !input.address.includes('@')
    ) {
      fail('mail: bad credential address')
    }
    if (input.secret === '' || input.secret.length > SECRET_MAX) fail('mail: bad credential secret')
    const allowed = scopesForProvider(input.provider)
    for (const scope of input.scopesGranted) {
      if (!allowed.includes(scope)) fail('mail: refusing broader OAuth scope')
    }
    if (input.provider === 'imap' && input.scopesGranted.length > 0) {
      fail('mail: refusing broader OAuth scope')
    }
  }

  /**
   * Store one secret. Persistent unless the account is session-only, the
   * launch is forced session-only, or the backend has no keyring (then
   * memory only, unless the explicit unprotected opt-in is on).
   */
  async putCredential(input: PutCredentialInput): Promise<CredentialSummary> {
    this.ensureReady()
    this.validateInput(input)
    const posture = await this.backendPosture()
    const sessionOnly =
      input.sessionOnly === true || this.forceSessionOnly || posture === 'session-only'
    const persist =
      !sessionOnly ||
      (posture === 'session-only' &&
        !this.forceSessionOnly &&
        this.allowUnprotectedPersistence &&
        input.sessionOnly !== true)
    const encrypted = await this.vault.encryptStringAsync(input.secret)
    const blob = Buffer.from(encrypted).toString('base64')
    const { backend } = await this.probeBackend()
    const protection = persist && posture === 'protected-keyring' ? 'os-keyring' : 'obfuscated-only'
    const record: StoredCredentialRecord = {
      accountId: input.accountId,
      provider: input.provider,
      address: input.address,
      scopesGranted: [...input.scopesGranted],
      blob,
      backend,
      protection,
      sessionOnly: persist ? (input.sessionOnly ?? false) : true,
      updatedAt: this.clock(),
    }
    this.secrets.set(input.accountId, { secret: input.secret, sessionOnly: record.sessionOnly })
    this.metas.set(input.accountId, {
      provider: record.provider,
      address: record.address,
      scopesGranted: [...record.scopesGranted],
      backend: record.backend,
      protection: record.protection,
      updatedAt: record.updatedAt,
    })
    this.sessionFlags.set(input.accountId, record.sessionOnly)
    if (persist) {
      mkdirSync(this.dir, { recursive: true })
      writeFileSync(this.recordPath(input.accountId), JSON.stringify(record), { mode: 0o600 })
    }
    return this.describe(input.accountId)
  }

  /**
   * Resolve one secret: memory first, else decrypt the stored blob.
   * Decrypt failure wipes the stale blob and raises re-auth (C-7): never
   * retries with weaker storage. A backend change with a good decrypt
   * re-encrypts and persists under the current backend.
   */
  async getCredential(accountId: string): Promise<ResolvedSecret> {
    this.ensureReady()
    if (accountId === '') fail('mail: bad credential account id')
    const cached = this.secrets.get(accountId)
    if (cached) return { secret: cached.secret, sessionOnly: cached.sessionOnly }
    let raw: string
    try {
      raw = readFileSync(this.recordPath(accountId), 'utf8')
    } catch {
      throw new CredentialReauthRequiredError(accountId)
    }
    let record: unknown
    try {
      record = JSON.parse(raw) as unknown
    } catch {
      this.wipeAccount(accountId)
      throw new CredentialCorruptError(accountId)
    }
    if (!isValidRecord(record)) {
      this.wipeAccount(accountId)
      throw new CredentialCorruptError(accountId)
    }
    let secret: string
    try {
      secret = await this.vault.decryptStringAsync(Buffer.from(record.blob, 'base64'))
    } catch {
      this.wipeAccount(accountId)
      throw new CredentialReauthRequiredError(accountId)
    }
    this.secrets.set(accountId, { secret, sessionOnly: record.sessionOnly })
    this.metas.set(accountId, {
      provider: record.provider,
      address: record.address,
      scopesGranted: [...record.scopesGranted],
      backend: record.backend,
      protection: record.protection,
      updatedAt: record.updatedAt,
    })
    this.sessionFlags.set(accountId, record.sessionOnly)
    const { backend } = await this.probeBackend()
    if (!record.sessionOnly && record.backend !== backend) {
      const blob = Buffer.from(await this.vault.encryptStringAsync(secret)).toString('base64')
      const next: StoredCredentialRecord = { ...record, blob, backend, updatedAt: this.clock() }
      writeFileSync(this.recordPath(accountId), JSON.stringify(next), { mode: 0o600 })
    }
    return { secret, sessionOnly: record.sessionOnly }
  }

  /** True when a secret is cached in memory or a stored blob exists. */
  hasCredential(accountId: string): boolean {
    this.ensureReady()
    if (accountId === '') return false
    if (this.secrets.has(accountId)) return true
    try {
      readFileSync(this.recordPath(accountId), 'utf8')
      return true
    } catch {
      return false
    }
  }

  /** Per-account session-only toggle (C-10); consulted before any persist. */
  isSessionOnlyAccount(accountId: string): boolean {
    this.ensureReady()
    return this.sessionFlags.get(accountId) ?? this.forceSessionOnly
  }

  /**
   * Server rejected the token (invalid_grant): wipe the blob and cached
   * secrets, then require re-auth. Never loops refresh on invalid_grant.
   */
  reportInvalidGrant(accountId: string): void {
    this.ensureReady()
    this.wipeAccount(accountId)
  }

  /** Remove an account: drop memory secrets, tokens, flags, and the blob. */
  removeCredential(accountId: string): void {
    this.ensureReady()
    this.wipeAccount(accountId)
  }

  private wipeAccount(accountId: string): void {
    this.secrets.delete(accountId)
    this.metas.delete(accountId)
    this.accessTokens.delete(accountId)
    this.sessionFlags.delete(accountId)
    try {
      rmSync(this.recordPath(accountId))
    } catch {
      // Missing blob is already wiped; nothing to do.
    }
  }

  /**
   * Access tokens are memory-only (C-4): cached for driver calls, never
   * written to disk, logs, or IPC. Entries expire and drop on read.
   */
  cacheAccessToken(accountId: string, token: string, ttlMs = 45 * 60 * 1000): void {
    this.ensureReady()
    if (accountId === '' || token === '') fail('mail: bad access token cache write')
    this.accessTokens.set(accountId, { token, expiresAt: Date.now() + ttlMs })
  }

  takeAccessToken(accountId: string): string | undefined {
    this.ensureReady()
    const entry = this.accessTokens.get(accountId)
    if (!entry) return undefined
    if (entry.expiresAt <= Date.now()) {
      this.accessTokens.delete(accountId)
      return undefined
    }
    return entry.token
  }

  clearAccessToken(accountId: string): void {
    this.ensureReady()
    this.accessTokens.delete(accountId)
  }

  clearAllAccessTokens(): void {
    this.ensureReady()
    this.accessTokens.clear()
  }

  /**
   * Log-safe account summary: metadata only, no secret and no blob bytes.
   * The blob stays opaque to the rest of the app.
   */
  describe(accountId: string): CredentialSummary {
    this.ensureReady()
    const cached = this.secrets.get(accountId)
    const meta = this.metas.get(accountId)
    let record: StoredCredentialRecord | null = null
    if (!cached) {
      try {
        const parsed: unknown = JSON.parse(readFileSync(this.recordPath(accountId), 'utf8'))
        if (isValidRecord(parsed)) record = parsed
      } catch {
        record = null
      }
    }
    const sessionOnly =
      this.sessionFlags.get(accountId) ?? record?.sessionOnly ?? this.forceSessionOnly
    const provider = meta?.provider ?? record?.provider ?? 'imap'
    const address = meta?.address ?? record?.address ?? ''
    const scopesGranted = meta ? [...meta.scopesGranted] : record ? [...record.scopesGranted] : []
    const backend = meta?.backend ?? record?.backend ?? ''
    const protection = meta?.protection ?? record?.protection ?? 'obfuscated-only'
    const updatedAt = meta?.updatedAt ?? record?.updatedAt ?? ''
    return {
      accountId,
      provider,
      address,
      scopesGranted,
      backend,
      protection,
      sessionOnly,
      hasSecret: cached !== undefined || record !== null,
      updatedAt,
    }
  }
}
