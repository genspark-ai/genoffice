import { afterEach, describe, expect, it, vi } from 'vitest'
import { endpointUrl } from '../src/protocols/shared'
import { getProviderAdapter } from '../src/registry'
import { streamForProvider } from '../src/stream'
import { listCustomModels } from '../src/custom-models'
import { jsonResponse, okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

const customConfig = (baseUrl: string) => ({
  apiKey: 'k',
  model: 'my-model',
  baseUrl,
})

describe('endpointUrl', () => {
  it('adds the path to the path component and keeps the query', () => {
    expect(endpointUrl('https://host.test/v1?api-version=2024-05-01', 'chat/completions')).toBe(
      'https://host.test/v1/chat/completions?api-version=2024-05-01',
    )
  })

  it('keeps a query that the endpoint path itself carries', () => {
    expect(
      endpointUrl('https://host.test/v1beta', 'models/m:streamGenerateContent', '?alt=sse'),
    ).toBe('https://host.test/v1beta/models/m:streamGenerateContent?alt=sse')
  })

  it('drops the fragment, which is never sent to the server', () => {
    expect(endpointUrl('https://host.test/v1#frag', 'chat/completions')).toBe(
      'https://host.test/v1/chat/completions',
    )
  })

  it('tolerates trailing and leading slashes on both halves', () => {
    expect(endpointUrl('https://host.test/v1/', '/chat/completions')).toBe(
      'https://host.test/v1/chat/completions',
    )
    expect(endpointUrl('https://host.test', 'v1/messages')).toBe('https://host.test/v1/messages')
    expect(endpointUrl('https://host.test/', '')).toBe('https://host.test/')
  })
})

describe('custom base URL normalization', () => {
  it('keeps a query string and drops a fragment', () => {
    const endpoint = getProviderAdapter('custom').resolveEndpoint(
      customConfig('https://host.test/v1?api-version=2024-05-01#docs'),
    )
    expect(endpoint.protocol).toBe('openai-compatible')
    expect(endpoint.baseUrl).toBe('https://host.test/v1?api-version=2024-05-01')
  })

  it('refuses embedded credentials', () => {
    expect(() =>
      getProviderAdapter('custom').resolveEndpoint(customConfig('https://user:pw@host.test/v1')),
    ).toThrow(/credentials/i)
  })

  it('still refuses non-http schemes and empty values', () => {
    expect(() =>
      getProviderAdapter('custom').resolveEndpoint(customConfig('file:///etc/passwd')),
    ).toThrow(/http/i)
    expect(() => getProviderAdapter('custom').resolveEndpoint(customConfig(''))).toThrow(/Base URL/)
  })
})

describe('requests against a base URL with a query', () => {
  it('puts /chat/completions in the path, not inside the query', async () => {
    const fetchMock = vi.fn(async (_input: RequestInfo | URL) =>
      okResponse(sseStream(['data: {"choices":[{"delta":{"content":"hi"}}]}', 'data: [DONE]'])),
    )
    vi.stubGlobal('fetch', fetchMock)
    await streamForProvider(
      'custom',
      customConfig('https://host.test/v1?api-version=2024-05-01'),
      'sys',
      [{ role: 'user', text: 'q' }],
      [],
      100,
      {
        signal: new AbortController().signal,
        onDelta: () => undefined,
        onToolCall: () => undefined,
      },
    )
    const url = String(fetchMock.mock.calls[0]![0])
    expect(url).toBe('https://host.test/v1/chat/completions?api-version=2024-05-01')
    expect(new URL(url).pathname).toBe('/v1/chat/completions')
  })

  it('probes /models on the path too', async () => {
    const fetchMock = vi.fn(async (_input: RequestInfo | URL) =>
      jsonResponse({ data: [{ id: 'my-model' }] }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const catalog = await listCustomModels('https://host.test/v1?api-version=2024-05-01')
    expect(String(fetchMock.mock.calls[0]![0])).toBe(
      'https://host.test/v1/models?api-version=2024-05-01',
    )
    expect(catalog.models).toEqual(['my-model'])
  })

  it('returns an empty catalogue for an unparsable base instead of throwing', async () => {
    const fetchMock = vi.fn(async () => okResponse(sseStream([])))
    vi.stubGlobal('fetch', fetchMock)
    await expect(listCustomModels('not a url')).resolves.toEqual({ models: [], defaultModel: '' })
  })
})
