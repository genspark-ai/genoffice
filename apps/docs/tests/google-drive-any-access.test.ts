/**
 * setAnyoneAccess: regression test for AOF-437 P2 (a second "make public"
 * call was POSTing a duplicate 'anyone' permission instead of changing the
 * existing one's role). Electron + google-auth are mocked so this stays a
 * pure fetch-call-shape test.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({
  app: { getPath: () => '/tmp/genoffice-test-userdata' },
  shell: { openExternal: vi.fn() },
}))

vi.mock('../src/main/google-auth', () => ({
  getValidAccessToken: vi.fn(async () => ({ ok: true, accessToken: 'token-abc' })),
}))

const { setAnyoneAccess } = await import('../src/main/google-drive')

describe('setAnyoneAccess', () => {
  beforeEach(() => vi.restoreAllMocks())
  afterEach(() => vi.restoreAllMocks())

  it('PATCHes the existing anyone permission instead of POSTing a duplicate', async () => {
    const calls: Array<{ url: string; method: string; body?: string }> = []
    vi.spyOn(globalThis, 'fetch').mockImplementation(async (url, init) => {
      const method = (init as RequestInit)?.method ?? 'GET'
      calls.push({ url: url as string, method, body: (init as RequestInit)?.body as string })
      if (method === 'GET') {
        return new Response(
          JSON.stringify({
            permissions: [
              { id: 'perm-anyone-1', type: 'anyone', role: 'reader' },
              { id: 'perm-user-1', type: 'user', role: 'writer', emailAddress: 'a@b.com' },
            ],
          }),
          { status: 200 },
        )
      }
      return new Response(JSON.stringify({ id: 'perm-anyone-1' }), { status: 200 })
    })

    const result = await setAnyoneAccess('file-1', 'writer')
    expect(result.ok).toBe(true)

    const listCall = calls.find((c) => c.method === 'GET')
    expect(listCall).toBeDefined()

    const writeCall = calls.find((c) => c.method !== 'GET')
    expect(writeCall?.method).toBe('PATCH')
    expect(writeCall?.url).toContain('/permissions/perm-anyone-1')
    expect(JSON.parse(writeCall!.body!)).toEqual({ role: 'writer' })

    // never a second POST creating a duplicate 'anyone' permission
    expect(calls.some((c) => c.method === 'POST')).toBe(false)
  })

  it('POSTs a new anyone permission only when none exists yet', async () => {
    const calls: Array<{ url: string; method: string; body?: string }> = []
    vi.spyOn(globalThis, 'fetch').mockImplementation(async (url, init) => {
      const method = (init as RequestInit)?.method ?? 'GET'
      calls.push({ url: url as string, method, body: (init as RequestInit)?.body as string })
      if (method === 'GET') {
        return new Response(JSON.stringify({ permissions: [] }), { status: 200 })
      }
      return new Response(JSON.stringify({ id: 'perm-new' }), { status: 200 })
    })

    const result = await setAnyoneAccess('file-2', 'reader')
    expect(result.ok).toBe(true)

    const writeCall = calls.find((c) => c.method !== 'GET')
    expect(writeCall?.method).toBe('POST')
    expect(JSON.parse(writeCall!.body!)).toEqual({ type: 'anyone', role: 'reader' })
  })
})
