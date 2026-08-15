import { CodexError, type CodexAuthContext, type CodexErrorCode } from './types'

const CODEX_CLIENT_ID = 'app_EMoamEEZ73f0CkXaXp7hrann'
const CODEX_ISSUER = 'https://auth.openai.com'
const CODEX_REDIRECT_URI = 'http://localhost:1455/auth/callback'
const CODEX_SCOPE = 'openid profile email offline_access api.connectors.read api.connectors.invoke'
const REFRESH_SKEW_MS = 60_000
export interface CodexCredentials extends CodexAuthContext {
  refreshToken: string
  email?: string
}

/** Implement this with main-process credential storage. Never use renderer state or AiSettings. */
export interface CodexCredentialStore {
  get(): Promise<CodexCredentials | undefined>
  set(credentials: CodexCredentials): Promise<void>
  delete(): Promise<void>
}

export interface CodexAccountStatus {
  loggedIn: boolean
  email?: string
  /** Safe, localizable sign-in failure. */
  errorCode?: CodexErrorCode
}

export interface CodexLoginCallback {
  state: string
  code: string
}

export interface CodexCallbackHandle {
  /** Resolves only after the callback transport is listening. */
  ready: Promise<void>
  wait: Promise<CodexLoginCallback>
  cancel(): void
  /** Complete browser response after exchange/storage; false renders safe failure. */
  complete?(success: boolean): void
}

export interface CodexAuthDependencies {
  store: CodexCredentialStore
  clock(): number
  fetch(input: RequestInfo | URL, init?: RequestInit): Promise<Response>
  openBrowser(url: string): Promise<void> | void
  /** Bind localhost callback server before opening authorization URL. */
  beginCallback(state: string): CodexCallbackHandle
}

interface TokenResponse {
  access_token?: unknown
  refresh_token?: unknown
  id_token?: unknown
  expires_in?: unknown
}

class CodexTokenHttpError extends CodexError {
  constructor(status: number) {
    super([400, 401].includes(status) ? 'auth-expired' : 'auth-temporary', { status })
    this.name = 'CodexTokenHttpError'
  }
}

function base64Url(bytes: Uint8Array): string {
  return btoa(String.fromCharCode(...bytes))
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '')
}

function randomValue(): string {
  return base64Url(crypto.getRandomValues(new Uint8Array(32)))
}

async function pkceChallenge(verifier: string): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier))
  return base64Url(new Uint8Array(digest))
}

function decodeBase64Url(value: string): string {
  const normalized = value.replace(/-/g, '+').replace(/_/g, '/')
  const padded = normalized + '='.repeat((4 - (normalized.length % 4)) % 4)
  const binary = atob(padded)
  return new TextDecoder().decode(Uint8Array.from(binary, (char) => char.charCodeAt(0)))
}

function accountFromIdToken(idToken: string): { accountId: string; email?: string } | undefined {
  const payload = idToken.split('.')[1]
  if (!payload) return undefined
  try {
    const claims = JSON.parse(decodeBase64Url(payload)) as Record<string, unknown>
    const auth = claims['https://api.openai.com/auth']
    const accountId =
      auth && typeof auth === 'object'
        ? (auth as Record<string, unknown>).chatgpt_account_id
        : undefined
    if (typeof accountId !== 'string' || !accountId) return undefined
    return { accountId, ...(typeof claims.email === 'string' ? { email: claims.email } : {}) }
  } catch {
    return undefined
  }
}

function tokenField(value: unknown): string {
  if (typeof value !== 'string' || !value) throw new CodexError('provider-failure')
  return value
}

function expiresAt(clock: number, seconds: unknown): number {
  return clock + (typeof seconds === 'number' && Number.isFinite(seconds) ? seconds : 3600) * 1000
}

export class CodexAuthService {
  private refreshPromise: Promise<CodexAuthContext> | undefined
  private loginPromise: Promise<CodexAccountStatus> | undefined
  private activeLogin: CodexCallbackHandle | undefined

  constructor(private readonly deps: CodexAuthDependencies) {}

  async status(): Promise<CodexAccountStatus> {
    const credentials = await this.deps.store.get()
    if (!credentials) return { loggedIn: false }
    if (credentials.expiresAt > this.deps.clock() + REFRESH_SKEW_MS) {
      return { loggedIn: true, ...(credentials.email ? { email: credentials.email } : {}) }
    }
    try {
      await this.getContext()
      const refreshed = await this.deps.store.get()
      return refreshed
        ? { loggedIn: true, ...(refreshed.email ? { email: refreshed.email } : {}) }
        : { loggedIn: false }
    } catch (error) {
      const retained = await this.deps.store.get()
      const errorCode =
        error instanceof CodexError && error.code === 'auth-expired'
          ? 'auth-expired'
          : 'auth-temporary'
      return retained
        ? {
            loggedIn: true,
            ...(retained.email ? { email: retained.email } : {}),
            errorCode,
          }
        : { loggedIn: false, errorCode }
    }
  }

  login(): Promise<CodexAccountStatus> {
    if (this.loginPromise) return this.loginPromise
    const attempt = this.performLogin()
    const shared = attempt.finally(() => {
      if (this.loginPromise === shared) this.loginPromise = undefined
    })
    this.loginPromise = shared
    return shared
  }

  private async performLogin(): Promise<CodexAccountStatus> {
    const verifier = randomValue()
    const state = randomValue()
    const callback = this.deps.beginCallback(state)
    void callback.ready.catch(() => {})
    void callback.wait.catch(() => {})
    this.activeLogin = callback
    const authorize = new URL(`${CODEX_ISSUER}/oauth/authorize`)
    authorize.search = new URLSearchParams({
      response_type: 'code',
      client_id: CODEX_CLIENT_ID,
      redirect_uri: CODEX_REDIRECT_URI,
      scope: CODEX_SCOPE,
      code_challenge: await pkceChallenge(verifier),
      code_challenge_method: 'S256',
      id_token_add_organizations: 'true',
      codex_cli_simplified_flow: 'true',
      state,
      originator: 'codex_cli_rs',
    }).toString()
    try {
      await callback.ready
      await this.deps.openBrowser(authorize.toString())
      const result = await callback.wait
      if (result.state !== state) throw new CodexError('provider-failure')
      const credentials = await this.exchange({
        grant_type: 'authorization_code',
        code: result.code,
        code_verifier: verifier,
        redirect_uri: CODEX_REDIRECT_URI,
      })
      await this.deps.store.set(credentials)
      callback.complete?.(true)
      return { loggedIn: true, ...(credentials.email ? { email: credentials.email } : {}) }
    } catch (error) {
      callback.complete?.(false)
      throw error instanceof CodexError ? error : new CodexError('auth-temporary')
    } finally {
      if (this.activeLogin === callback) this.activeLogin = undefined
    }
  }

  cancelLogin(): void {
    this.activeLogin?.cancel()
  }

  async logout(): Promise<void> {
    this.cancelLogin()
    await this.deps.store.delete()
  }

  async getContext(): Promise<CodexAuthContext> {
    const credentials = await this.deps.store.get()
    if (!credentials) throw new CodexError('auth-required')
    if (credentials.expiresAt > this.deps.clock() + REFRESH_SKEW_MS)
      return this.context(credentials)
    if (!this.refreshPromise) {
      this.refreshPromise = this.refresh(credentials).finally(() => {
        this.refreshPromise = undefined
      })
    }
    return this.refreshPromise
  }

  private context(credentials: CodexCredentials): CodexAuthContext {
    return {
      accessToken: credentials.accessToken,
      accountId: credentials.accountId,
      expiresAt: credentials.expiresAt,
    }
  }

  private async refresh(current: CodexCredentials): Promise<CodexAuthContext> {
    const result = await this.exchange(
      { grant_type: 'refresh_token', refresh_token: current.refreshToken },
      current,
    ).then(
      (credentials) => ({ credentials }),
      (error: unknown) => ({ error }),
    )
    if ('error' in result) {
      if (result.error instanceof CodexError && result.error.code === 'auth-expired') {
        await this.deps.store.delete()
        throw result.error
      }
      throw new CodexError('auth-temporary')
    }
    try {
      await this.deps.store.set(result.credentials)
    } catch {
      throw new CodexError('auth-temporary')
    }
    return this.context(result.credentials)
  }

  private async exchange(
    fields: Record<string, string>,
    current?: CodexCredentials,
  ): Promise<CodexCredentials> {
    const response = await this.deps.fetch(`${CODEX_ISSUER}/oauth/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ client_id: CODEX_CLIENT_ID, ...fields }).toString(),
    })
    if (!response.ok) throw new CodexTokenHttpError(response.status)
    let tokens: TokenResponse
    try {
      tokens = (await response.json()) as TokenResponse
    } catch {
      throw new CodexError('provider-failure')
    }
    const accessToken = tokenField(tokens.access_token)
    const account =
      typeof tokens.id_token === 'string' ? accountFromIdToken(tokens.id_token) : undefined
    if (!current && !account) throw new CodexError('provider-failure')
    const email = account?.email ?? current?.email
    return {
      accessToken,
      refreshToken:
        typeof tokens.refresh_token === 'string' && tokens.refresh_token
          ? tokens.refresh_token
          : (current?.refreshToken ?? tokenField(tokens.refresh_token)),
      accountId: account?.accountId ?? current!.accountId,
      ...(email ? { email } : {}),
      expiresAt: expiresAt(this.deps.clock(), tokens.expires_in),
    }
  }
}
