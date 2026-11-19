import { afterEach, describe, expect, it, vi } from 'vitest'
import { AI_DEFAULT_USER_AGENT, aiFetch, setAiUserAgent, setRescueFetch } from '../src/fetch'

afterEach(() => {
  vi.unstubAllGlobals()
  setRescueFetch(null)
  setAiUserAgent(AI_DEFAULT_USER_AGENT)
})

function sentHeaders(fetchMock: ReturnType<typeof vi.fn>): Headers {
  const init = fetchMock.mock.calls[0]![1] as RequestInit
  return new Headers(init.headers)
}

describe('aiFetch', () => {
  it('identifies the client to gateways that flag anonymous traffic', async () => {
    const fetchMock = vi.fn().mockResolvedValue(new Response('ok'))
    vi.stubGlobal('fetch', fetchMock)
    await aiFetch('https://x/', { headers: { Authorization: 'Bearer k' } })
    const headers = sentHeaders(fetchMock)
    expect(headers.get('user-agent')).toBe('GenOffice')
    expect(headers.get('authorization')).toBe('Bearer k')
  })

  it('lets the host refine the user agent and never overrides an explicit one', async () => {
    const fetchMock = vi.fn().mockResolvedValue(new Response('ok'))
    vi.stubGlobal('fetch', fetchMock)
    setAiUserAgent('GenOffice/1.2.3')
    await aiFetch('https://x/', {})
    expect(sentHeaders(fetchMock).get('user-agent')).toBe('GenOffice/1.2.3')

    fetchMock.mockClear()
    await aiFetch('https://x/', { headers: { 'User-Agent': 'custom/9' } })
    expect(sentHeaders(fetchMock).get('user-agent')).toBe('custom/9')
  })

  it('passes the same identified request to the rescue fetch', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('fetch failed')))
    const rescue = vi.fn().mockResolvedValue(new Response('rescued'))
    setRescueFetch(rescue)
    await aiFetch('https://x/', {})
    expect(sentHeaders(rescue).get('user-agent')).toBe('GenOffice')
  })

  it('returns the primary response without touching the rescue path', async () => {
    const ok = new Response('ok')
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(ok))
    const rescue = vi.fn()
    setRescueFetch(rescue)
    expect(await aiFetch('https://x/', {})).toBe(ok)
    expect(rescue).not.toHaveBeenCalled()
  })

  it('retries over the rescue fetch when the primary fails', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('fetch failed')))
    const ok = new Response('rescued')
    const rescue = vi.fn().mockResolvedValue(ok)
    setRescueFetch(rescue)
    expect(await aiFetch('https://x/', {})).toBe(ok)
    expect(rescue).toHaveBeenCalledOnce()
  })

  it('throws the primary error when no rescue fetch is set', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('ECONNRESET')))
    await expect(aiFetch('https://x/', {})).rejects.toThrow('ECONNRESET')
  })

  it('throws the primary error when the rescue fetch also fails', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('primary down')))
    setRescueFetch(vi.fn().mockRejectedValue(new Error('rescue down')))
    await expect(aiFetch('https://x/', {})).rejects.toThrow('primary down')
  })

  it('retries a 403 HTML block page over the rescue fetch', async () => {
    const blocked = () =>
      new Response('<!doctype html><title>Just a moment...</title>', {
        status: 403,
        headers: { 'content-type': 'text/html; charset=utf-8' },
      })
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(blocked()))
    const ok = new Response('{}', { headers: { 'content-type': 'application/json' } })
    const rescue = vi.fn().mockResolvedValue(ok)
    setRescueFetch(rescue)
    expect(await aiFetch('https://www.genspark.ai/api/x', { method: 'POST', body: '{}' })).toBe(ok)
    expect(rescue).toHaveBeenCalledOnce()

    // still blocked on the rescue path: the primary answer stands
    const primary = blocked()
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(primary))
    setRescueFetch(vi.fn().mockResolvedValue(blocked()))
    expect(await aiFetch('https://www.genspark.ai/api/x', { body: '{}' })).toBe(primary)
  })

  it('does not treat an API 403 or a stream body as a block page', async () => {
    const denied = new Response('{"error":"forbidden"}', {
      status: 403,
      headers: { 'content-type': 'application/json' },
    })
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(denied))
    const rescue = vi.fn()
    setRescueFetch(rescue)
    expect(await aiFetch('https://x/', { body: '{}' })).toBe(denied)
    expect(rescue).not.toHaveBeenCalled()

    const html = new Response('<html>', { status: 403, headers: { 'content-type': 'text/html' } })
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(html))
    expect(await aiFetch('https://x/', { body: new ReadableStream() })).toBe(html)
    expect(rescue).not.toHaveBeenCalled()
  })

  it('does not retry an aborted request', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('aborted')))
    const rescue = vi.fn()
    setRescueFetch(rescue)
    const controller = new AbortController()
    controller.abort()
    await expect(aiFetch('https://x/', { signal: controller.signal })).rejects.toThrow('aborted')
    expect(rescue).not.toHaveBeenCalled()
  })
})
