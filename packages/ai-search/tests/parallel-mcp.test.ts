import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { defaultAiSettings } from '@genoffice/ai-provider'
import { webSearch } from '../src/index'
import { searchOptionsFromSettings, testSearchProvider } from '../src/search-tools'

const endpoint = 'https://search.parallel.ai/mcp'
const source = { title: 'GenOffice', url: 'https://example.com', excerpts: ['First.', 'Second.'] }
const payload = { results: [source, { ...source, title: 'Extra' }] }
const fallback = () =>
  new Response('<a class="result__a" href="https://fallback.example.com">Fallback</a>')

beforeEach(() => {
  vi.stubEnv('AI_SEARCH_DISABLE_GSK', '1')
  for (const key of ['SERPER_API_KEY', 'TAVILY_API_KEY', 'PARALLEL_API_KEY']) vi.stubEnv(key, '')
})

afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
  vi.useRealTimers()
})

// Exercise the real MCP client against a simulated Streamable HTTP server.
function mockServer(result: Record<string, unknown>, sse = false) {
  const requests: { url: string; init: RequestInit; body: Record<string, unknown> }[] = []
  const fetch = vi.fn<typeof globalThis.fetch>(async (input, init = {}) => {
    const url = String(input)
    const body = init.body ? JSON.parse(String(init.body)) : {}
    requests.push({ url, init, body })
    if (url !== endpoint) return fallback()
    if (init.method === 'GET') return new Response('', { status: 405 })
    if (init.method === 'DELETE') return new Response(null, { status: 204 })
    if (body.method === 'initialize') {
      return Response.json(
        {
          jsonrpc: '2.0',
          id: body.id,
          result: {
            protocolVersion: body.params.protocolVersion,
            capabilities: { tools: {} },
            serverInfo: { name: 'parallel-test', version: '1' },
          },
        },
        { headers: { 'mcp-session-id': 'test-mcp-session' } },
      )
    }
    if (body.method === 'tools/call') {
      const rpc = { jsonrpc: '2.0', id: body.id, result }
      return sse
        ? new Response(`event: message\ndata: ${JSON.stringify(rpc)}\n\n`, {
            headers: { 'Content-Type': 'text/event-stream' },
          })
        : Response.json(rpc)
    }
    return new Response(null, { status: 202 })
  })
  vi.stubGlobal('fetch', fetch)
  return { requests, fetch }
}

describe('Parallel free Search MCP', () => {
  it.each([false, true])('searches anonymously and maps results (SSE: %s)', async (sse) => {
    const { requests } = mockServer({ content: [], structuredContent: payload }, sse)
    const r = await webSearch('office tools', 1, {
      useGsk: false,
      prefer: 'parallel',
      parallelKey: '',
    })
    expect(r).toEqual({
      method: 'parallel',
      results: [{ title: 'GenOffice', url: 'https://example.com', snippet: 'First.\nSecond.' }],
    })
    expect(requests.find((r) => r.body.method === 'tools/call')?.body.params).toEqual({
      name: 'web_search',
      arguments: { objective: 'office tools', search_queries: ['office tools'] },
    })
    expect(requests.every((r) => r.url === endpoint)).toBe(true)
    const cleanup = requests.find((r) => r.init.method === 'DELETE')
    expect(cleanup).toBeDefined()
    expect(new Headers(cleanup?.init.headers).get('mcp-session-id')).toBe('test-mcp-session')
    for (const { init } of requests) {
      const headers = new Headers(init.headers)
      expect(headers.has('Authorization')).toBe(false)
      expect(headers.has('x-api-key')).toBe(false)
    }
  })

  it('accepts JSON text content when structured content is absent', async () => {
    mockServer({ content: [{ type: 'text', text: JSON.stringify(payload) }] })
    expect((await webSearch('office', 1, { prefer: 'parallel' })).method).toBe('parallel')
  })

  it('honors a blank saved key even when an environment key is present', async () => {
    vi.stubEnv('PARALLEL_API_KEY', 'environment-key')
    const settings = defaultAiSettings()
    settings.search!.provider = 'parallel'
    const { requests } = mockServer({ content: [], structuredContent: payload })
    expect(searchOptionsFromSettings(settings)).toMatchObject({
      useGsk: false,
      prefer: 'parallel',
      parallelKey: '',
    })
    expect((await webSearch('office', 1, searchOptionsFromSettings(settings))).method).toBe(
      'parallel',
    )
    expect(requests.every((r) => r.url === endpoint)).toBe(true)
  })

  it('answers the unconfigured default chain anonymously before DuckDuckGo', async () => {
    const { requests } = mockServer({ content: [], structuredContent: payload })
    expect((await webSearch('office')).method).toBe('parallel')
    expect(requests.every((r) => r.url === endpoint)).toBe(true)
  })

  it('runs after a failing keyed provider and ahead of DuckDuckGo', async () => {
    vi.stubEnv('TAVILY_API_KEY', 'tavily-key')
    const { requests, fetch } = mockServer({ content: [], structuredContent: payload })
    const respond = fetch.getMockImplementation()!
    fetch.mockImplementation((url, init) =>
      String(url) === 'https://api.tavily.com/search'
        ? Promise.resolve(new Response('', { status: 401 }))
        : respond(url, init),
    )
    expect((await webSearch('office')).method).toBe('parallel')
    expect(String(fetch.mock.calls[0]?.[0])).toBe('https://api.tavily.com/search')
    expect(requests.some((r) => r.url.includes('duckduckgo.com'))).toBe(false)
  })

  it.each([
    { isError: true, content: [{ type: 'text', text: 'Rate limit reached' }] },
    { content: [], structuredContent: { results: [] } },
    { content: [{ type: 'text', text: 'not JSON' }] },
  ])('falls back without reporting success for an unusable MCP result: %j', async (result) => {
    const { requests } = mockServer(result)
    expect((await testSearchProvider('parallel', '')).ok).toBe(false)
    expect(requests.some((r) => r.body.method === 'tools/call')).toBe(true)
    expect(requests.at(-1)?.url).toContain('duckduckgo.com')
  })

  it('tests the free service successfully with no key', async () => {
    mockServer({ content: [], structuredContent: payload })
    expect(await testSearchProvider('parallel', '   ')).toEqual({ ok: true })
  })

  it.each([401, 429, 503])('falls back on MCP HTTP %i without trying OAuth', async (status) => {
    const fetch = vi.fn<typeof globalThis.fetch>(async (url) =>
      String(url) === endpoint ? new Response('', { status }) : fallback(),
    )
    vi.stubGlobal('fetch', fetch)
    expect((await webSearch('office', 1, { prefer: 'parallel' })).method).toBe('duckduckgo')
    expect(String(fetch.mock.calls[0]?.[0])).toBe(endpoint)
    expect(fetch.mock.calls).toHaveLength(2)
  })

  it('bounds a stalled MCP handshake and aborts the network request', async () => {
    vi.useFakeTimers()
    let signal: AbortSignal | null | undefined
    vi.stubGlobal(
      'fetch',
      vi.fn<typeof globalThis.fetch>((url, init) => {
        if (String(url) !== endpoint) return Promise.resolve(fallback())
        signal = init?.signal
        return new Promise<Response>((_resolve, reject) => {
          signal?.addEventListener('abort', () => reject(new Error('aborted')), { once: true })
        })
      }),
    )
    const pending = webSearch('office', 1, { prefer: 'parallel' })
    await vi.advanceTimersByTimeAsync(15000)
    expect((await pending).method).toBe('duckduckgo')
    expect(signal?.aborted).toBe(true)
  })

  it.each([
    ['tools/call', 'duckduckgo'],
    ['DELETE', 'parallel'],
  ])('bounds stalled %s without losing completed results', async (stage, expectedMethod) => {
    vi.useFakeTimers()
    const { fetch } = mockServer({ content: [], structuredContent: payload })
    const respond = fetch.getMockImplementation()!
    let signal: AbortSignal | null | undefined
    fetch.mockImplementation((url, init) => {
      const method = init?.body ? JSON.parse(String(init.body)).method : init?.method
      if (String(url) !== endpoint || method !== stage) return respond(url, init)
      signal = init?.signal
      return new Promise<Response>((_resolve, reject) => {
        signal?.addEventListener('abort', () => reject(new Error('aborted')), { once: true })
      })
    })
    const pending = webSearch('office', 1, { prefer: 'parallel' })
    await vi.advanceTimersByTimeAsync(15000)
    expect((await pending).method).toBe(expectedMethod)
    expect(signal?.aborted).toBe(true)
  })
})
