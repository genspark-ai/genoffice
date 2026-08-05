import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  genofficeApiKey,
  genofficeAuthPath,
  genofficeLoginInFlight,
  genofficeLogout,
  loadGenofficeAuth,
  resetGenofficeAuthCache,
  startGenofficeLogin,
  type GskLoginProgress,
} from '../src/genoffice-auth'
import { gskApiKey } from '../src/gsk'

const CODE = 'a'.repeat(64)
const AUTH_URL = `https://www.genspark.ai/api/office_addin_auth/verify?code=${CODE}`

let dir: string

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'genoffice-auth-'))
  process.env.GENOFFICE_AUTH_DIR = dir
  delete process.env.GSK_API_KEY
  resetGenofficeAuthCache()
})

afterEach(() => {
  vi.unstubAllGlobals()
  rmSync(dir, { recursive: true, force: true })
  delete process.env.GENOFFICE_AUTH_DIR
  delete process.env.GSK_API_KEY
  resetGenofficeAuthCache()
})

function jsonResponse(json: unknown, opts: { status?: number; setCookie?: string } = {}): Response {
  const headers = new Headers()
  if (opts.setCookie) headers.append('set-cookie', opts.setCookie)
  return {
    ok: (opts.status ?? 200) < 400,
    status: opts.status ?? 200,
    json: async () => json,
    headers,
  } as unknown as Response
}

/** fetch stub for the full flow; token polls: pending × (pendingPolls) then approved */
function stubFlow(opts: { pendingPolls?: number; createResponse?: unknown } = {}) {
  let tokenPolls = 0
  const fetchMock = vi.fn(async (input: string | URL, _init?: RequestInit) => {
    const url = String(input)
    if (url.includes('/office_addin_auth/device_code')) {
      return jsonResponse({
        device_code: CODE,
        auth_url: AUTH_URL,
        expires_in: 600,
        poll_interval: 0.001,
      })
    }
    if (url.includes('/office_addin_auth/token')) {
      tokenPolls++
      if (tokenPolls <= (opts.pendingPolls ?? 1)) return jsonResponse({ status: 'pending' })
      return jsonResponse({ status: 'approved', access_token: 'bearer-token' })
    }
    if (url.includes('/office_addin_auth/session')) {
      return jsonResponse(
        { status: 'ok', cogen_id: 'user-1' },
        { setCookie: 'session_id=sess-abc; Path=/; HttpOnly' },
      )
    }
    if (url.includes('/api_tokens/create')) {
      return jsonResponse(
        opts.createResponse ?? {
          status: 0,
          data: { key_id: 'kid-1', key_name: 'genoffice', token: 'gsk-genoffice-key' },
        },
      )
    }
    if (url.includes('/api_tokens/revoke')) return jsonResponse({ status: 0 })
    throw new Error(`unexpected fetch: ${url}`)
  })
  vi.stubGlobal('fetch', fetchMock)
  return fetchMock
}

async function loginAndCollect(): Promise<GskLoginProgress[]> {
  const events: GskLoginProgress[] = []
  startGenofficeLogin((progress) => events.push(progress))
  await vi.waitFor(() => {
    expect(['success', 'error']).toContain(events.at(-1)?.phase)
  })
  return events
}

describe('startGenofficeLogin', () => {
  it('runs device_code → poll → session → key create and stores the key', async () => {
    const fetchMock = stubFlow({ pendingPolls: 2 })
    const events = await loginAndCollect()

    expect(events[0]).toEqual({ phase: 'url', url: AUTH_URL, expiresInSec: 600 })
    expect(events.at(-1)).toEqual({ phase: 'success' })

    const saved = JSON.parse(readFileSync(genofficeAuthPath(), 'utf-8'))
    expect(saved).toEqual({
      api_key: 'gsk-genoffice-key',
      key_id: 'kid-1',
      access_token: 'bearer-token',
    })
    expect(statSync(genofficeAuthPath()).mode & 0o777).toBe(0o600)
    expect(genofficeApiKey()).toBe('gsk-genoffice-key')
    expect(genofficeLoginInFlight()).toBe(false)

    // the key create call must ride on the session cookie
    const createCall = fetchMock.mock.calls.find(([u]) => String(u).includes('/api_tokens/create'))!
    const init = createCall[1]!
    expect((init.headers as Record<string, string>).Cookie).toBe('session_id=sess-abc')
    expect(JSON.parse(String(init.body))).toEqual({ key_name: 'genoffice' })
  })

  it('feeds gskApiKey(), losing only to an explicit GSK_API_KEY env override', async () => {
    stubFlow()
    await loginAndCollect()
    expect(gskApiKey()).toBe('gsk-genoffice-key')
    process.env.GSK_API_KEY = 'gsk-env-override'
    expect(gskApiKey()).toBe('gsk-env-override')
  })

  it('reports a network failure at device_code as error "network"', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('ECONNREFUSED')))
    const events = await loginAndCollect()
    expect(events).toEqual([{ phase: 'error', error: 'network' }])
    expect(loadGenofficeAuth()).toBeNull()
  })

  it('reports an expired device code as error "expired"', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: string | URL) => {
        const url = String(input)
        if (url.includes('/device_code')) {
          return jsonResponse({
            device_code: CODE,
            auth_url: AUTH_URL,
            expires_in: 600,
            poll_interval: 0.001,
          })
        }
        return jsonResponse({ status: 'expired' })
      }),
    )
    const events = await loginAndCollect()
    expect(events.at(-1)).toEqual({ phase: 'error', error: 'expired' })
  })

  it('surfaces the server message when key creation fails', async () => {
    stubFlow({ createResponse: { status: -1, message: 'Invalid API key name.' } })
    const events = await loginAndCollect()
    expect(events.at(-1)).toEqual({ phase: 'error', error: 'Invalid API key name.' })
    expect(loadGenofficeAuth()).toBeNull()
  })

  it('re-login revokes the superseded key (best-effort)', async () => {
    stubFlow()
    await loginAndCollect()

    const fetchMock = stubFlow({
      createResponse: { status: 0, data: { key_id: 'kid-2', token: 'gsk-genoffice-key-2' } },
    })
    await loginAndCollect()
    expect(loadGenofficeAuth()).toMatchObject({ apiKey: 'gsk-genoffice-key-2', keyId: 'kid-2' })
    await vi.waitFor(() => {
      const revoke = fetchMock.mock.calls.find(([u]) => String(u).includes('/api_tokens/revoke'))
      expect(revoke).toBeDefined()
      expect(JSON.parse(String(revoke![1]!.body))).toEqual({ key_id: 'kid-1' })
    })
  })

  it('a new login cancels the in-flight one without emitting on it', async () => {
    // first flow polls pending forever
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: string | URL) => {
        const url = String(input)
        if (url.includes('/device_code')) {
          return jsonResponse({
            device_code: CODE,
            auth_url: AUTH_URL,
            expires_in: 600,
            poll_interval: 0.001,
          })
        }
        return jsonResponse({ status: 'pending' })
      }),
    )
    const first: GskLoginProgress[] = []
    startGenofficeLogin((progress) => first.push(progress))
    await vi.waitFor(() => expect(first.length).toBeGreaterThan(0))
    expect(genofficeLoginInFlight()).toBe(true)

    stubFlow()
    const second = await loginAndCollect()
    expect(second.at(-1)).toEqual({ phase: 'success' })
    expect(first.every((e) => e.phase === 'url')).toBe(true)
  })
})

describe('genofficeLogout', () => {
  it('revokes the key server-side and removes the local file', async () => {
    const fetchMock = stubFlow()
    await loginAndCollect()

    await genofficeLogout()
    expect(existsSync(genofficeAuthPath())).toBe(false)
    expect(genofficeApiKey()).toBe('')
    const revokeCall = fetchMock.mock.calls.find(([u]) => String(u).includes('/api_tokens/revoke'))!
    const init = revokeCall[1]!
    expect((init.headers as Record<string, string>).Cookie).toBe('session_id=sess-abc')
    expect(JSON.parse(String(init.body))).toEqual({ key_id: 'kid-1' })
  })

  it('still clears locally when the server revoke fails', async () => {
    stubFlow()
    await loginAndCollect()
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('offline')))
    await genofficeLogout()
    expect(existsSync(genofficeAuthPath())).toBe(false)
    expect(loadGenofficeAuth()).toBeNull()
  })

  it('is a no-op network-wise when not signed in', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)
    await genofficeLogout()
    expect(fetchMock).not.toHaveBeenCalled()
  })
})
