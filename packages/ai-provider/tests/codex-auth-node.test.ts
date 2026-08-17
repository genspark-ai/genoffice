import {
  chmodSync,
  existsSync,
  mkdtempSync,
  readFileSync,
  renameSync,
  rmSync,
  statSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs'
import { createServer, type AddressInfo } from 'node:net'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { beginCodexCallback, codexCredentialStore } from '../src/codex-auth-node'

const credentials = {
  accessToken: 'access',
  refreshToken: 'refresh',
  accountId: 'account',
  expiresAt: 1,
  email: 'person@example.test',
}

let authDir: string
const originalAuthDir = process.env.GENOFFICE_AUTH_DIR

async function openPort(port = 0) {
  const server = createServer()
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject)
    server.listen(port, '127.0.0.1', resolve)
  })
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('test server has no address')
  return { server, port: (address as AddressInfo).port }
}

async function closeServer(server: ReturnType<typeof createServer>): Promise<void> {
  await new Promise<void>((resolve, reject) =>
    server.close((error) => (error ? reject(error) : resolve())),
  )
}

async function startCallback(state: string) {
  const handle = beginCodexCallback(state, 0)
  await handle.ready
  if (!handle.port) throw new Error('callback server did not expose a port')
  return { handle, port: handle.port }
}

beforeEach(() => {
  authDir = mkdtempSync(join(tmpdir(), 'codex-auth-json-'))
  process.env.GENOFFICE_AUTH_DIR = authDir
})

afterEach(() => {
  if (originalAuthDir === undefined) delete process.env.GENOFFICE_AUTH_DIR
  else process.env.GENOFFICE_AUTH_DIR = originalAuthDir
  rmSync(authDir, { recursive: true, force: true })
})

describe('codexCredentialStore', () => {
  it('writes JSON credentials with restrictive permissions', async () => {
    const store = codexCredentialStore()

    await store.set(credentials)

    const path = join(authDir, 'codex-auth.json')
    expect(JSON.parse(readFileSync(path, 'utf8'))).toEqual({
      type: 'oauth',
      access: 'access',
      refresh: 'refresh',
      expires: 1,
      accountId: 'account',
      email: 'person@example.test',
    })
    expect(statSync(path).mode & 0o777).toBe(0o600)
    await expect(store.get()).resolves.toEqual(credentials)
  })

  it.skipIf(process.platform === 'win32')(
    'rejects group/world-readable credential files without deleting them',
    async () => {
      const store = codexCredentialStore()
      const path = join(authDir, 'codex-auth.json')
      await store.set(credentials)
      chmodSync(path, 0o644)

      await expect(store.get()).rejects.toThrow('ChatGPT sign-in credentials unavailable')
      expect(existsSync(path)).toBe(true)
    },
  )

  it.skipIf(process.platform === 'win32')(
    'rejects symlinked credential files without deleting either path',
    async () => {
      const store = codexCredentialStore()
      const path = join(authDir, 'codex-auth.json')
      const fixturePath = join(authDir, 'codex-auth-fixture.json')
      await store.set(credentials)
      renameSync(path, fixturePath)
      symlinkSync(fixturePath, path)

      await expect(store.get()).rejects.toThrow('ChatGPT sign-in credentials unavailable')
      expect(existsSync(path)).toBe(true)
      expect(existsSync(fixturePath)).toBe(true)
    },
  )

  it('treats malformed JSON as signed out', async () => {
    const path = join(authDir, 'codex-auth.json')
    writeFileSync(path, '{broken', { mode: 0o600 })

    await expect(codexCredentialStore().get()).resolves.toBeUndefined()
    expect(existsSync(path)).toBe(false)
  })

  it('deletes the JSON credential file on logout', async () => {
    const store = codexCredentialStore()
    await store.set(credentials)
    await store.delete()

    expect(existsSync(join(authDir, 'codex-auth.json'))).toBe(false)
  })
})

describe('beginCodexCallback', () => {
  it('keeps listening after an invalid callback and completes a later valid callback', async () => {
    const { handle, port } = await startCallback('valid-state')

    const invalid = await fetch(`http://127.0.0.1:${port}/auth/callback?state=wrong&code=code`)
    expect(invalid.status).toBe(400)
    let settled = false
    void handle.wait.then(
      () => {
        settled = true
      },
      () => {
        settled = true
      },
    )
    await Promise.resolve()
    expect(settled).toBe(false)

    const validResponse = fetch(
      `http://127.0.0.1:${port}/auth/callback?state=valid-state&code=code`,
    )
    await expect(handle.wait).resolves.toEqual({ state: 'valid-state', code: 'code' })
    handle.complete?.(true)
    expect((await validResponse).status).toBe(200)
  })

  it('does not replace the first valid callback response', async () => {
    const { handle, port } = await startCallback('valid-state')

    const firstResponse = fetch(
      `http://127.0.0.1:${port}/auth/callback?state=valid-state&code=first-code`,
    )
    await expect(handle.wait).resolves.toEqual({ state: 'valid-state', code: 'first-code' })

    const duplicateResponse = await fetch(
      `http://127.0.0.1:${port}/auth/callback?state=valid-state&code=second-code`,
    )
    expect(duplicateResponse.status).toBe(409)
    expect(await duplicateResponse.text()).toBe('Sign-in callback already received.')

    handle.complete?.(true)
    expect((await firstResponse).status).toBe(200)
  })

  it('rejects readiness and wait on a port collision', async () => {
    const occupied = await openPort()
    const handle = beginCodexCallback('state', occupied.port)

    await expect(handle.ready).rejects.toThrow('callback unavailable')
    await expect(handle.wait).rejects.toThrow('callback unavailable')
    await closeServer(occupied.server)
  })

  it('cancels and closes the callback server once', async () => {
    const { handle, port } = await startCallback('state')

    const wait = handle.wait.catch((error: unknown) => error)
    handle.cancel()
    handle.cancel()
    await expect(wait).resolves.toMatchObject({ message: 'ChatGPT sign-in cancelled' })
    await expect(fetch(`http://127.0.0.1:${port}/auth/callback`)).rejects.toThrow()
  })

  it('times out and closes the callback server', async () => {
    vi.useFakeTimers()
    try {
      const { handle, port } = await startCallback('state')

      const result = expect(handle.wait).rejects.toThrow('ChatGPT sign-in timed out')
      await vi.advanceTimersByTimeAsync(5 * 60_000)
      await result
      await expect(fetch(`http://127.0.0.1:${port}/auth/callback`)).rejects.toThrow()
    } finally {
      vi.useRealTimers()
    }
  })
})
