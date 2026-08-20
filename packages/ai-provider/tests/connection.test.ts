import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { testProviderConnection } from '../src/connection'

describe('testProviderConnection', () => {
  let fetchSpy: ReturnType<typeof vi.fn>

  beforeEach(() => {
    fetchSpy = vi.fn()
    vi.stubGlobal('fetch', fetchSpy)
  })

  afterEach(() => {
    vi.restoreAllMocks()
  })

  it('ollama: connected on 200 from /api/tags', async () => {
    fetchSpy.mockResolvedValueOnce({ ok: true, status: 200 })
    const result = await testProviderConnection('ollama', {
      apiKey: '',
      model: '',
      baseUrl: 'http://localhost:11434/v1',
    })
    expect(result).toEqual({ ok: true, status: 'connected' })
    expect(fetchSpy).toHaveBeenCalledWith(
      'http://localhost:11434/api/tags',
      expect.objectContaining({ method: 'GET' }),
    )
  })

  it('ollama: not-running on ECONNREFUSED (server down)', async () => {
    fetchSpy.mockRejectedValueOnce(Object.assign(new Error('fetch failed'), { cause: { code: 'ECONNREFUSED' } }))
    const result = await testProviderConnection('ollama', { apiKey: '', model: '' })
    expect(result).toEqual({ ok: false, status: 'not-running' })
  })

  it('openai: connected with Bearer auth on GET /models', async () => {
    fetchSpy.mockResolvedValueOnce({ ok: true, status: 200 })
    const result = await testProviderConnection('openai', {
      apiKey: 'sk-test',
      model: 'gpt-4.1-mini',
      baseUrl: 'https://api.openai.com/v1',
    })
    expect(result.ok).toBe(true)
    const [url, init] = fetchSpy.mock.calls[0]
    expect(url).toBe('https://api.openai.com/v1/models')
    expect((init.headers as Record<string, string>).Authorization).toBe('Bearer sk-test')
  })

  it('auth: 401 maps to auth', async () => {
    fetchSpy.mockResolvedValueOnce({ ok: false, status: 401 })
    const result = await testProviderConnection('openai', { apiKey: 'bad', model: '' })
    expect(result).toEqual({ ok: false, status: 'auth' })
  })

  it('invalid endpoint: 404 maps to invalid', async () => {
    fetchSpy.mockResolvedValueOnce({ ok: false, status: 404 })
    const result = await testProviderConnection('custom', { apiKey: '', model: '', baseUrl: 'http://x/v1' })
    expect(result).toEqual({ ok: false, status: 'invalid' })
  })

  it('timeout: TimeoutError maps to timeout', async () => {
    const err = new Error('The operation was aborted due to timeout')
    err.name = 'TimeoutError'
    fetchSpy.mockRejectedValueOnce(err)
    const result = await testProviderConnection('gemini', { apiKey: 'k', model: '' })
    expect(result).toEqual({ ok: false, status: 'timeout' })
  })

  it('cloud: network failure maps to refused (not not-running)', async () => {
    fetchSpy.mockRejectedValueOnce(new Error('fetch failed'))
    const result = await testProviderConnection('openai', { apiKey: 'k', model: '' })
    expect(result).toEqual({ ok: false, status: 'refused' })
  })

  it('anthropic: hits /v1/models with x-api-key', async () => {
    fetchSpy.mockResolvedValueOnce({ ok: true, status: 200 })
    await testProviderConnection('anthropic', { apiKey: 'sk-ant', model: '' })
    const [url, init] = fetchSpy.mock.calls[0]
    expect(url).toBe('https://api.anthropic.com/v1/models')
    expect((init.headers as Record<string, string>)['x-api-key']).toBe('sk-ant')
  })
})
