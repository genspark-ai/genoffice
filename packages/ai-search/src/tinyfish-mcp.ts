/**
 * TinyFish web access over MCP (main process). Connects to TinyFish's hosted MCP
 * server with browser OAuth (authorization code + PKCE + Dynamic Client
 * Registration) and calls its `fetch_content` tool — so GenOffice never stores a
 * TinyFish API key locally; the user signs in once and a token is kept on disk.
 *
 * The official SDK's OAuthClientProvider drives the whole spec flow; we supply the
 * same three integration points as the OpenWorker connector this mirrors:
 *   - token/registration storage → a plain ~/.genoffice/tinyfish-mcp.json file
 *     (matches how genoffice-auth.ts stores the Genspark key; not an OS keychain)
 *   - redirect  → open the system browser (an opener is injected by the caller,
 *     so this module stays free of an Electron dependency and is testable)
 *   - callback  → a loopback http server catches the ?code= redirect
 *
 * Interactive sign-in (browser + loopback wait) is an explicit-action privilege:
 * the fetch tool runs non-interactive and, when there is no token yet, returns a
 * clear "connect first" error instead of hijacking the user's browser mid-answer.
 */

import { createServer, type Server } from 'node:http'
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, join } from 'node:path'

import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { UnauthorizedError, type OAuthClientProvider } from '@modelcontextprotocol/sdk/client/auth.js'
import type {
  OAuthClientInformation,
  OAuthClientInformationFull,
  OAuthClientMetadata,
  OAuthTokens,
} from '@modelcontextprotocol/sdk/shared/auth.js'

/** TinyFish's hosted MCP server (overridable for staging/tests). */
const MCP_URL = () => process.env.TINYFISH_MCP_URL || 'https://agent.tinyfish.ai/mcp'
const CLIENT_NAME = 'GenOffice'
const FLOW_TIMEOUT_MS = 300_000
/**
 * Loopback redirect port. Fixed (not OS-assigned) so the redirect address stays
 * stable across sign-ins: dynamic client registration records this exact address,
 * and a later re-auth from a different port would be rejected. Overridable in case
 * the default is taken.
 */
const CALLBACK_PORT = () => Number(process.env.TINYFISH_OAUTH_PORT) || 33418

/** Where tokens + the dynamic-registration record live (override for tests). */
export function tinyfishAuthPath(): string {
  return join(process.env.GENOFFICE_AUTH_DIR || join(homedir(), '.genoffice'), 'tinyfish-mcp.json')
}

// ── on-disk storage ─────────────────────────────────────────────────
interface StoredAuth {
  tokens?: OAuthTokens
  clientInformation?: OAuthClientInformationFull
}

function readStore(): StoredAuth {
  try {
    return JSON.parse(readFileSync(tinyfishAuthPath(), 'utf-8')) as StoredAuth
  } catch {
    return {}
  }
}

function writeStore(patch: StoredAuth): void {
  const path = tinyfishAuthPath()
  mkdirSync(dirname(path), { recursive: true })
  writeFileSync(path, JSON.stringify({ ...readStore(), ...patch }, null, 2), { mode: 0o600 })
}

/** True when a TinyFish token is already stored (drives the "connected" status). */
export function hasTinyFishAuth(): boolean {
  return !!readStore().tokens
}

/** Forget the token and the registration; the next connect runs a fresh sign-in. */
export function signOutTinyFish(): void {
  try {
    rmSync(tinyfishAuthPath())
  } catch {
    /* already gone */
  }
}

/** Raised when a background context needs sign-in but must not open a browser. */
export class TinyFishAuthRequired extends Error {
  constructor() {
    super('TinyFish sign-in required — connect TinyFish first')
    this.name = 'TinyFishAuthRequired'
  }
}

// ── OAuth client provider (the SDK's three hooks) ───────────────────
class TinyFishOAuthProvider implements OAuthClientProvider {
  private verifier = ''

  constructor(
    private readonly redirectUri: string,
    private readonly openBrowser: (url: string) => void | Promise<void>,
  ) {}

  get redirectUrl(): string {
    return this.redirectUri
  }

  get clientMetadata(): OAuthClientMetadata {
    return {
      client_name: CLIENT_NAME,
      redirect_uris: [this.redirectUri],
      grant_types: ['authorization_code', 'refresh_token'],
      response_types: ['code'],
      // public client: dynamic registration issues no secret a desktop app could keep
      token_endpoint_auth_method: 'none',
    }
  }

  clientInformation(): OAuthClientInformation | undefined {
    return readStore().clientInformation
  }

  saveClientInformation(info: OAuthClientInformationFull): void {
    writeStore({ clientInformation: info })
  }

  tokens(): OAuthTokens | undefined {
    return readStore().tokens
  }

  saveTokens(tokens: OAuthTokens): void {
    writeStore({ tokens })
  }

  saveCodeVerifier(verifier: string): void {
    this.verifier = verifier
  }

  codeVerifier(): string {
    if (!this.verifier) throw new Error('no PKCE code verifier for this flow')
    return this.verifier
  }

  redirectToAuthorization(authorizationUrl: URL): void | Promise<void> {
    return this.openBrowser(authorizationUrl.toString())
  }
}

// ── loopback callback catcher ───────────────────────────────────────
interface Loopback {
  redirectUri: string
  waitForCode(): Promise<string>
  close(): void
}

// Only one interactive sign-in can hold the fixed loopback port at a time. If the
// user clicks Connect again after abandoning the first browser tab, the stale
// listener is still holding the port — track it so a new sign-in cancels it first.
let activeLoopback: Loopback | null = null

/** Start a one-shot http server on 127.0.0.1 that resolves with the ?code=. */
function startLoopback(): Promise<Loopback> {
  // Cancel any abandoned prior sign-in so its port is free to rebind.
  activeLoopback?.close()
  activeLoopback = null

  return new Promise((resolve, reject) => {
    let onCode: ((code: string) => void) | null = null
    let onError: ((err: Error) => void) | null = null

    const server: Server = createServer((req, res) => {
      const url = new URL(req.url ?? '/', 'http://127.0.0.1')
      if (url.pathname !== '/callback') {
        res.writeHead(404).end()
        return
      }
      const code = url.searchParams.get('code')
      const err = url.searchParams.get('error')
      res.writeHead(200, { 'content-type': 'text/html' }).end(
        '<html><body style="font:16px system-ui;padding:2rem">You can close this tab and return to GenOffice.</body></html>',
      )
      if (code) onCode?.(code)
      else onError?.(new Error(`authorization failed: ${err ?? 'no code returned'}`))
    })

    const port = CALLBACK_PORT()
    // The prior listener may not have released the port yet — retry briefly on EADDRINUSE.
    const attempt = (retriesLeft: number): void => {
      server.removeAllListeners('error')
      server.once('error', (err: NodeJS.ErrnoException) => {
        if (err.code === 'EADDRINUSE' && retriesLeft > 0) {
          setTimeout(() => attempt(retriesLeft - 1), 300)
        } else {
          reject(err)
        }
      })
      server.listen(port, '127.0.0.1', () => {
        const handle: Loopback = {
          redirectUri: `http://127.0.0.1:${port}/callback`,
          waitForCode: () =>
            new Promise<string>((res, rej) => {
              onCode = res
              onError = rej
              setTimeout(() => rej(new Error('sign-in timed out')), FLOW_TIMEOUT_MS)
            }),
          close: () => {
            onError?.(new Error('sign-in canceled'))
            if (activeLoopback === handle) activeLoopback = null
            server.close()
          },
        }
        activeLoopback = handle
        resolve(handle)
      })
    }
    attempt(5)
  })
}

// ── connect (with optional interactive sign-in) ─────────────────────
/** The fixed loopback redirect the provider advertises (must match DCR registration). */
function fixedRedirectUri(): string {
  return `http://127.0.0.1:${CALLBACK_PORT()}/callback`
}

async function connect(
  openBrowser: (url: string) => void | Promise<void>,
  interactive: boolean,
): Promise<{ client: Client; close: () => Promise<void> }> {
  // Only bind the loopback catcher for interactive sign-in. A non-interactive fetch
  // either already has a token (refresh needs no redirect) or short-circuits before
  // here, so binding a port on every fetch would be wasteful and clash when two
  // fetches overlap. Either way the advertised redirect URI is the same fixed one,
  // so the dynamic-registration record stays valid.
  const loopback = interactive ? await startLoopback() : null
  const redirectUri = loopback?.redirectUri ?? fixedRedirectUri()
  const provider = new TinyFishOAuthProvider(redirectUri, openBrowser)
  const transport = new StreamableHTTPClientTransport(new URL(MCP_URL()), { authProvider: provider })
  const client = new Client({ name: CLIENT_NAME, version: '0.1.0' }, { capabilities: {} })

  const close = async () => {
    try {
      await client.close()
    } finally {
      loopback?.close()
    }
  }

  try {
    await client.connect(transport)
  } catch (err) {
    if (!(err instanceof UnauthorizedError)) {
      loopback?.close()
      throw err
    }
    // No usable token: connect() already opened the browser via redirectToAuthorization.
    if (!interactive || !loopback) {
      loopback?.close()
      throw new TinyFishAuthRequired()
    }
    const code = await loopback.waitForCode()
    await transport.finishAuth(code)
    await client.connect(transport) // reconnect, now authorized
  }
  return { client, close }
}

// ── public surface ──────────────────────────────────────────────────

/** Runs the interactive browser sign-in and stores the token. Throws on failure. */
export async function signInToTinyFish(
  openBrowser: (url: string) => void | Promise<void>,
): Promise<void> {
  const { close } = await connect(openBrowser, true)
  await close()
}

/**
 * Fetch a page's readable text via TinyFish's MCP `fetch_content` tool. Runs
 * non-interactive: if there is no stored token, throws TinyFishAuthRequired so the
 * caller can tell the user to connect rather than popping a browser mid-answer.
 */
export async function fetchViaTinyFish(url: string): Promise<string> {
  // No stored token → don't even open a connection; let the caller prompt sign-in.
  if (!hasTinyFishAuth()) throw new TinyFishAuthRequired()
  const { client, close } = await connect(() => {}, false)
  try {
    const result = await client.callTool({
      name: 'fetch_content',
      // fetch_content requires urls/format/links/image_links; we only need page text.
      arguments: { urls: [url], format: 'markdown', links: false, image_links: false },
    })
    return extractText(result)
  } finally {
    await close()
  }
}

/**
 * Pull the clean page text out of a fetch_content result. TinyFish returns a
 * { results: [{ url, title, text }] } payload — either as structuredContent or as
 * JSON inside the text block — so we unwrap it to the readable `text` rather than
 * handing the raw JSON envelope to the model. Falls back to the raw text if the
 * shape is ever different.
 */
function extractText(result: unknown): string {
  const r = (result ?? {}) as {
    content?: Array<{ type?: string; text?: string }>
    structuredContent?: unknown
  }
  const rawText = (r.content ?? [])
    .filter((c) => c.type === 'text' && typeof c.text === 'string')
    .map((c) => c.text as string)
    .join('\n')
    .trim()

  const payload = r.structuredContent ?? safeJsonParse(rawText)
  const pageText = pagesToText(payload)
  return pageText || rawText
}

function safeJsonParse(s: string): unknown {
  try {
    return JSON.parse(s)
  } catch {
    return undefined
  }
}

/** Flatten a { results: [{ title, url, text }] } payload into readable text. */
function pagesToText(payload: unknown): string {
  const results = (payload as { results?: unknown } | undefined)?.results
  if (!Array.isArray(results)) return ''
  const parts: string[] = []
  for (const item of results) {
    const rec = (item ?? {}) as { title?: unknown; url?: unknown; text?: unknown }
    const text = typeof rec.text === 'string' ? rec.text.trim() : ''
    if (!text) continue
    const header = [rec.title, rec.url].filter((v) => typeof v === 'string' && v).join(' — ')
    parts.push(header ? `${header}\n\n${text}` : text)
  }
  return parts.join('\n\n---\n\n').trim()
}
