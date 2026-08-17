import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { listOllamaModels } from '../src/ollama'

describe('listOllamaModels', () => {
  let fetchSpy: ReturnType<typeof vi.fn>

  beforeEach(() => {
    fetchSpy = vi.fn()
    vi.stubGlobal('fetch', fetchSpy)
  })

  afterEach(() => {
    vi.restoreAllMocks()
  })

  it('returns parsed models from /api/tags', async () => {
    fetchSpy.mockResolvedValueOnce({
      ok: true,
      json: () =>
        Promise.resolve({
          models: [
            {
              name: 'llama3.2:3b',
              modified_at: '2024-01-01T00:00:00Z',
              details: { parameter_size: '3B' },
            },
            {
              name: 'codellama:7b',
              modified_at: '2024-01-02T00:00:00Z',
              details: { parameter_size: '7B' },
            },
          ],
        }),
    })

    const result = await listOllamaModels()
    expect(result.models).toHaveLength(2)
    expect(result.models[0]).toEqual({
      name: 'llama3.2:3b',
      parameterSize: '3B',
      modifiedAt: '2024-01-01T00:00:00Z',
    })
    expect(result.models[1]).toEqual({
      name: 'codellama:7b',
      parameterSize: '7B',
      modifiedAt: '2024-01-02T00:00:00Z',
    })
  })

  it('calls correct URL with /v1 stripped', async () => {
    fetchSpy.mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({ models: [] }) })
    await listOllamaModels('http://localhost:11434/v1')
    expect(fetchSpy).toHaveBeenCalledWith(
      'http://localhost:11434/api/tags',
      expect.objectContaining({ method: 'GET' }),
    )
  })

  it('returns empty list and error on non-ok response', async () => {
    fetchSpy.mockResolvedValueOnce({ ok: false, status: 503 })
    const result = await listOllamaModels()
    expect(result.models).toHaveLength(0)
    expect(result.error).toContain('503')
  })

  it('returns empty list and error on fetch failure', async () => {
    fetchSpy.mockRejectedValueOnce(new Error('ECONNREFUSED'))
    const result = await listOllamaModels()
    expect(result.models).toHaveLength(0)
    expect(result.error).toContain('ECONNREFUSED')
  })

  it('handles missing models array gracefully', async () => {
    fetchSpy.mockResolvedValueOnce({ ok: true, json: () => Promise.resolve({}) })
    const result = await listOllamaModels()
    expect(result.models).toEqual([])
  })

  it('handles missing details gracefully', async () => {
    fetchSpy.mockResolvedValueOnce({
      ok: true,
      json: () => Promise.resolve({ models: [{ name: 'qwen:7b' }] }),
    })
    const result = await listOllamaModels()
    expect(result.models[0]).toEqual({
      name: 'qwen:7b',
      parameterSize: undefined,
      modifiedAt: undefined,
    })
  })
})
