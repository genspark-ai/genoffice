/**
 * Google OAuth 2.0 (Desktop app, loopback + PKCE) for Drive access.
 *
 * Credential resolution order (first hit wins), never logged:
 *   1. env GENOFFICE_GOOGLE_CLIENT_ID / GENOFFICE_GOOGLE_CLIENT_SECRET
 *   2. macOS Keychain generic password (account "icarus", services
 *      GENOFFICE_GOOGLE_CLIENT_ID / GENOFFICE_GOOGLE_CLIENT_SECRET)
 *   3. JSON file at userData/google-oauth.json: { "clientId": "...", "clientSecret": "..." }
 *
 * Tokens (access + refresh) are stored at userData/google-tokens.json, mode 0600.
 * Never logged, never sent anywhere but Google's token endpoint.
 */
import { execFile } from 'node:child_process'
import { createHash, randomBytes } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { chmod, writeFile } from 'node:fs/promises'
import { createServer } from 'node:http'
import { platform } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { app, shell } from 'electron'

const execFileAsync = promisify(execFile)

const AUTH_ENDPOINT = 'https://accounts.google.com/o/oauth2/v2/auth'
const TOKEN_ENDPOINT = 'https://oauth2.googleapis.com/token'

/** drive.file: files this app creates/opens. drive.readonly: listing + exporting
 *  the user's existing Google Docs for the import picker. */
export const GOOGLE_SCOPES = [
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/drive.readonly',
].join(' ')

export interface GoogleOAuthCredentials {
  clientId: string
  clientSecret: string
}

export interface GoogleTokens {
  accessToken: string
  refreshToken: string
  /** epoch ms */
  expiresAt: number
  scope: string
}

function userDataPath(...parts: string[]): string {
  return join(app.getPath('userData'), ...parts)
}

const TOKENS_PATH = () => userDataPath('google-tokens.json')
const OAUTH_CONFIG_PATH = () => userDataPath('google-oauth.json')

async function readKeychainSecret(service: string): Promise<string | null> {
  if (platform() !== 'darwin') return null
  try {
    const { stdout } = await execFileAsync('security', [
      'find-generic-password',
      '-a',
      'icarus',
      '-s',
      service,
      '-w',
    ])
    const value = stdout.trim()
    return value || null
  } catch {
    return null
  }
}

let cachedCredentials: GoogleOAuthCredentials | null | undefined

/** Resolve client id/secret; result cached for the process lifetime (Keychain
 *  reads can be slow and this is called on every token refresh). */
export async function resolveGoogleCredentials(): Promise<GoogleOAuthCredentials | null> {
  if (cachedCredentials !== undefined) return cachedCredentials

  const envId = process.env.GENOFFICE_GOOGLE_CLIENT_ID
  const envSecret = process.env.GENOFFICE_GOOGLE_CLIENT_SECRET
  if (envId && envSecret) {
    cachedCredentials = { clientId: envId, clientSecret: envSecret }
    return cachedCredentials
  }

  const [kcId, kcSecret] = await Promise.all([
    readKeychainSecret('GENOFFICE_GOOGLE_CLIENT_ID'),
    readKeychainSecret('GENOFFICE_GOOGLE_CLIENT_SECRET'),
  ])
  if (kcId && kcSecret) {
    cachedCredentials = { clientId: kcId, clientSecret: kcSecret }
    return cachedCredentials
  }

  try {
    if (existsSync(OAUTH_CONFIG_PATH())) {
      const parsed = JSON.parse(readFileSync(OAUTH_CONFIG_PATH(), 'utf-8')) as Partial<
        Record<'clientId' | 'clientSecret', string>
      >
      if (parsed.clientId && parsed.clientSecret) {
        cachedCredentials = { clientId: parsed.clientId, clientSecret: parsed.clientSecret }
        return cachedCredentials
      }
    }
  } catch {
    /* malformed config file: treat as unconfigured */
  }

  cachedCredentials = null
  return null
}

/** test-only: reset the credential cache */
export function _resetGoogleCredentialsCacheForTests(): void {
  cachedCredentials = undefined
}

function base64url(input: Buffer): string {
  return input.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
}

function pkcePair(): { verifier: string; challenge: string } {
  const verifier = base64url(randomBytes(32))
  const challenge = base64url(createHash('sha256').update(verifier).digest())
  return { verifier, challenge }
}

async function readTokens(): Promise<GoogleTokens | null> {
  try {
    if (!existsSync(TOKENS_PATH())) return null
    return JSON.parse(readFileSync(TOKENS_PATH(), 'utf-8')) as GoogleTokens
  } catch {
    return null
  }
}

async function writeTokens(tokens: GoogleTokens): Promise<void> {
  await writeFile(TOKENS_PATH(), JSON.stringify(tokens), { mode: 0o600 })
  await chmod(TOKENS_PATH(), 0o600).catch(() => {})
}

/** true once a refresh token has been stored (does not verify it still works) */
export async function isGoogleSignedIn(): Promise<boolean> {
  const tokens = await readTokens()
  return !!tokens?.refreshToken
}

export async function signOutGoogle(): Promise<void> {
  try {
    await writeFile(TOKENS_PATH(), JSON.stringify(null), { mode: 0o600 })
  } catch {
    /* best-effort */
  }
}

/** Run the full loopback + PKCE authorization flow: opens the system browser,
 *  waits for the redirect on a local ephemeral port, exchanges the code. */
export async function runGoogleAuthFlow(): Promise<
  { ok: true } | { ok: false; error: string; unconfigured?: boolean }
> {
  const creds = await resolveGoogleCredentials()
  if (!creds) return { ok: false, error: 'Google account not configured', unconfigured: true }

  const { verifier, challenge } = pkcePair()
  const state = base64url(randomBytes(16))

  return new Promise((resolve) => {
    let redirectUri = ''
    const server = createServer((req, res) => {
      const url = new URL(req.url ?? '/', 'http://127.0.0.1')
      if (url.pathname !== '/callback') {
        res.writeHead(404).end()
        return
      }
      const code = url.searchParams.get('code')
      const returnedState = url.searchParams.get('state')
      const error = url.searchParams.get('error')
      res.writeHead(200, { 'Content-Type': 'text/html' })
      if (error || !code || returnedState !== state) {
        res.end(
          '<html><body>Sign-in failed. You can close this tab and return to the app.</body></html>',
        )
        server.close()
        resolve({ ok: false, error: error ?? 'invalid response from Google' })
        return
      }
      res.end('<html><body>Signed in. You can close this tab and return to the app.</body></html>')
      server.close()
      void (async () => {
        try {
          const tokenRes = await fetch(TOKEN_ENDPOINT, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
              grant_type: 'authorization_code',
              code,
              redirect_uri: redirectUri,
              client_id: creds.clientId,
              client_secret: creds.clientSecret,
              code_verifier: verifier,
            }),
          })
          if (!tokenRes.ok) {
            resolve({ ok: false, error: `token exchange failed (${tokenRes.status})` })
            return
          }
          const body = (await tokenRes.json()) as {
            access_token: string
            refresh_token?: string
            expires_in: number
            scope: string
          }
          if (!body.refresh_token) {
            // Google omits refresh_token on re-consent for an already-authorized
            // app; keep the previous one rather than losing offline access.
            const prev = await readTokens()
            if (!prev?.refreshToken) {
              resolve({ ok: false, error: 'Google did not return a refresh token' })
              return
            }
            await writeTokens({
              accessToken: body.access_token,
              refreshToken: prev.refreshToken,
              expiresAt: Date.now() + body.expires_in * 1000,
              scope: body.scope,
            })
          } else {
            await writeTokens({
              accessToken: body.access_token,
              refreshToken: body.refresh_token,
              expiresAt: Date.now() + body.expires_in * 1000,
              scope: body.scope,
            })
          }
          resolve({ ok: true })
        } catch (err) {
          resolve({ ok: false, error: String(err) })
        }
      })()
    })

    server.listen(0, '127.0.0.1', () => {
      const port = (server.address() as { port: number }).port
      redirectUri = `http://127.0.0.1:${port}/callback`
      const authUrl = new URL(AUTH_ENDPOINT)
      authUrl.searchParams.set('client_id', creds.clientId)
      authUrl.searchParams.set('redirect_uri', redirectUri)
      authUrl.searchParams.set('response_type', 'code')
      authUrl.searchParams.set('scope', GOOGLE_SCOPES)
      authUrl.searchParams.set('access_type', 'offline')
      authUrl.searchParams.set('prompt', 'consent')
      authUrl.searchParams.set('state', state)
      authUrl.searchParams.set('code_challenge', challenge)
      authUrl.searchParams.set('code_challenge_method', 'S256')
      void shell.openExternal(authUrl.toString())
    })

    // loopback flows are user-paced (switching to the browser, picking an
    // account); 5 minutes is generous without hanging forever on abandonment
    setTimeout(
      () => {
        server.close()
        resolve({ ok: false, error: 'timed out waiting for Google sign-in' })
      },
      5 * 60 * 1000,
    ).unref()
  })
}

/** Exchange the stored refresh token for a fresh access token. Exported
 *  separately (not folded into getValidAccessToken) so it is easy to unit-test
 *  the refresh math without spinning up the loopback server. */
export async function refreshAccessToken(
  creds: GoogleOAuthCredentials,
  refreshToken: string,
  fetchImpl: typeof fetch = fetch,
): Promise<{ ok: true; tokens: GoogleTokens } | { ok: false; error: string }> {
  const res = await fetchImpl(TOKEN_ENDPOINT, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_id: creds.clientId,
      client_secret: creds.clientSecret,
    }),
  })
  if (!res.ok) return { ok: false, error: `refresh failed (${res.status})` }
  const body = (await res.json()) as { access_token: string; expires_in: number; scope: string }
  return {
    ok: true,
    tokens: {
      accessToken: body.access_token,
      refreshToken,
      expiresAt: Date.now() + body.expires_in * 1000,
      scope: body.scope,
    },
  }
}

/** True when the access token is missing or expires within this window —
 *  refreshing a little early avoids a race against an in-flight Drive call. */
export function tokenNeedsRefresh(tokens: GoogleTokens, skewMs = 60_000): boolean {
  return Date.now() + skewMs >= tokens.expiresAt
}

/** Valid (non-expired) access token, refreshing and persisting as needed.
 *  Returns null when never signed in or unconfigured. */
export async function getValidAccessToken(): Promise<
  { ok: true; accessToken: string } | { ok: false; error: string; unconfigured?: boolean }
> {
  const creds = await resolveGoogleCredentials()
  if (!creds) return { ok: false, error: 'Google account not configured', unconfigured: true }
  const tokens = await readTokens()
  if (!tokens?.refreshToken) return { ok: false, error: 'not signed in to Google' }
  if (!tokenNeedsRefresh(tokens)) return { ok: true, accessToken: tokens.accessToken }
  const refreshed = await refreshAccessToken(creds, tokens.refreshToken)
  if (!refreshed.ok) return { ok: false, error: refreshed.error }
  await writeTokens(refreshed.tokens)
  return { ok: true, accessToken: refreshed.tokens.accessToken }
}
