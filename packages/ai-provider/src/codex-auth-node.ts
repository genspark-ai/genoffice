import { randomUUID } from 'node:crypto'
import { chmod, mkdir, open, readFile, rename, unlink } from 'node:fs/promises'
import { createServer, type ServerResponse } from 'node:http'
import { homedir } from 'node:os'
import { dirname, join } from 'node:path'
import type { CodexCallbackHandle, CodexCredentialStore, CodexCredentials } from './auth'

const CODEX_AUTH_DIR = () => process.env.GENOFFICE_AUTH_DIR || join(homedir(), '.genoffice')
const CODEX_CREDENTIALS_PATH = () => join(CODEX_AUTH_DIR(), 'codex-auth.json')
const CODEX_CALLBACK_PORT = 1455
const CODEX_LOGIN_TIMEOUT_MS = 5 * 60_000

interface StoredCodexCredentials {
  type: 'oauth'
  access: string
  refresh: string
  expires: number
  accountId: string
  email?: string
}

function fileNotFound(error: unknown): boolean {
  return typeof error === 'object' && error !== null && 'code' in error && error.code === 'ENOENT'
}

function credentialFileError(): Error {
  return new Error('ChatGPT sign-in credentials unavailable')
}

function isStoredCodexCredentials(value: unknown): value is StoredCodexCredentials {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const record = value as Record<string, unknown>
  return (
    record.type === 'oauth' &&
    typeof record.access === 'string' &&
    record.access.length > 0 &&
    typeof record.refresh === 'string' &&
    record.refresh.length > 0 &&
    typeof record.accountId === 'string' &&
    record.accountId.length > 0 &&
    typeof record.expires === 'number' &&
    Number.isFinite(record.expires) &&
    (record.email === undefined || typeof record.email === 'string')
  )
}

function toStoredCodexCredentials(credentials: CodexCredentials): StoredCodexCredentials {
  return {
    type: 'oauth',
    access: credentials.accessToken,
    refresh: credentials.refreshToken,
    expires: credentials.expiresAt,
    accountId: credentials.accountId,
    ...(credentials.email ? { email: credentials.email } : {}),
  }
}

function fromStoredCodexCredentials(value: unknown): CodexCredentials | undefined {
  if (!isStoredCodexCredentials(value)) return undefined
  return {
    accessToken: value.access,
    refreshToken: value.refresh,
    accountId: value.accountId,
    expiresAt: value.expires,
    ...(value.email ? { email: value.email } : {}),
  }
}

async function removeCredentialFile(path: string): Promise<void> {
  try {
    await unlink(path)
  } catch (error) {
    if (!fileNotFound(error)) throw credentialFileError()
  }
}

async function writeCredentialFile(path: string, credentials: CodexCredentials): Promise<void> {
  const directory = dirname(path)
  const temporaryPath = join(directory, `.codex-auth-${randomUUID()}.tmp`)
  try {
    await mkdir(directory, { recursive: true, mode: 0o700 })
    const handle = await open(temporaryPath, 'wx', 0o600)
    try {
      await handle.writeFile(`${JSON.stringify(toStoredCodexCredentials(credentials), null, 2)}\n`)
      await handle.sync()
    } finally {
      await handle.close()
    }
    await rename(temporaryPath, path)
    await chmod(path, 0o600)
  } catch {
    throw credentialFileError()
  } finally {
    await unlink(temporaryPath).catch(() => {})
  }
}

async function readCredentialFile(path: string): Promise<CodexCredentials | undefined> {
  let raw: string
  try {
    raw = await readFile(path, 'utf8')
  } catch (error) {
    if (fileNotFound(error)) return undefined
    throw credentialFileError()
  }
  try {
    const parsed: unknown = JSON.parse(raw)
    const credentials = fromStoredCodexCredentials(parsed)
    if (credentials) return credentials
  } catch {
    // Treat malformed credentials as signed out, then remove the untrusted file.
  }
  await removeCredentialFile(path)
  return undefined
}

export function codexCredentialStore(): CodexCredentialStore {
  return {
    async get() {
      return readCredentialFile(CODEX_CREDENTIALS_PATH())
    },
    async set(credentials) {
      await writeCredentialFile(CODEX_CREDENTIALS_PATH(), credentials)
    },
    async delete() {
      await removeCredentialFile(CODEX_CREDENTIALS_PATH())
    },
  }
}

export function beginCodexCallback(
  expectedState: string,
  port = CODEX_CALLBACK_PORT,
): CodexCallbackHandle {
  let server: ReturnType<typeof createServer> | undefined
  let pendingResponse: ServerResponse | undefined
  let finish: ((result: { state: string; code: string }) => void) | undefined
  let fail: ((error: Error) => void) | undefined
  let settled = false
  let completed = false
  let resolveReady!: () => void
  let rejectReady!: (error: Error) => void
  let readySettled = false
  const ready = new Promise<void>((resolve, reject) => {
    resolveReady = resolve
    rejectReady = reject
  })
  const close = () => {
    const current = server
    server = undefined
    try {
      current?.close()
    } catch {
      // The server may already have failed before it started listening.
    }
  }
  const markReady = () => {
    if (readySettled) return
    readySettled = true
    resolveReady()
  }
  const settle = (result: { state: string; code: string }) => {
    if (settled) return
    settled = true
    if (timeout) clearTimeout(timeout)
    finish?.(result)
  }
  const reject = (error: Error) => {
    if (settled) return
    settled = true
    if (!readySettled) {
      readySettled = true
      rejectReady(error)
    }
    if (timeout) clearTimeout(timeout)
    if (pendingResponse) {
      pendingResponse
        .writeHead(500, { 'Content-Type': 'text/html; charset=utf-8' })
        .end('Sign-in could not be completed. You may close this window.')
      pendingResponse = undefined
    }
    close()
    fail?.(error)
  }
  const complete = (success: boolean) => {
    if (completed) return
    completed = true
    if (pendingResponse) {
      pendingResponse
        .writeHead(success ? 200 : 500, { 'Content-Type': 'text/html; charset=utf-8' })
        .end(
          success
            ? 'Sign-in complete. You may close this window.'
            : 'Sign-in could not be completed. You may close this window.',
        )
      pendingResponse = undefined
    }
    close()
  }
  const wait = new Promise<{ state: string; code: string }>((resolve, rejectWait) => {
    finish = resolve
    fail = rejectWait
    server = createServer((request, response) => {
      const url = new URL(request.url ?? '/', `http://127.0.0.1:${port}`)
      if (url.pathname !== '/auth/callback') {
        response.writeHead(404).end('Not found')
        return
      }
      const state = url.searchParams.get('state')
      const code = url.searchParams.get('code')
      if (!state || !code || state !== expectedState) {
        response.writeHead(400).end('Sign-in could not be completed.')
        return
      }
      if (settled) {
        response
          .writeHead(409, { 'Content-Type': 'text/plain; charset=utf-8' })
          .end('Sign-in callback already received.')
        return
      }
      pendingResponse = response
      settle({ state, code })
    })
    server.once('listening', markReady)
    server.once('error', () => reject(new Error('ChatGPT sign-in callback unavailable')))
    server.listen(port, '127.0.0.1')
  })
  const timeout = setTimeout(
    () => reject(new Error('ChatGPT sign-in timed out')),
    CODEX_LOGIN_TIMEOUT_MS,
  )
  timeout.unref()
  return {
    ready,
    wait,
    cancel: () => reject(new Error('ChatGPT sign-in cancelled')),
    complete,
  }
}
