import { afterEach, describe, expect, it, vi } from 'vitest'
import { chatForProvider } from '../src/chat'
import { errorResponse, jsonResponse } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('chatForProvider', () => {
  it('anthropic: extracts joined text content blocks', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        jsonResponse({
          content: [
            { type: 'text', text: 'hello ' },
            { type: 'text', text: 'world' },
          ],
        }),
      ),
    )
    const result = await chatForProvider(
      'anthropic',
      { apiKey: 'k', model: 'claude-sonnet-5' },
      'sys',
      'hi',
    )
    expect(result).toEqual({ ok: true, content: 'hello world' })
  })

  it('anthropic: surfaces HTTP errors', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(errorResponse(401, 'bad key')))
    const result = await chatForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', 'hi')
    expect(result.ok).toBe(false)
    expect(result.error).toMatch(/Claude HTTP 401/)
  })

  it('anthropic: replaces an HTML error body with a readable note', async () => {
    const html =
      '<!doctype html>\n<html>\n<head><title>Genspark</title></head><body>app shell</body></html>'
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(errorResponse(403, html)))
    const result = await chatForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', 'hi')
    expect(result.ok).toBe(false)
    expect(result.error).toMatch(/Claude HTTP 403/)
    expect(result.error).toMatch(/web page instead of an API response/)
    expect(result.error).not.toContain('<!doctype')
  })

  it('gemini: extracts joined parts text', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(
          jsonResponse({ candidates: [{ content: { parts: [{ text: 'hi there' }] } }] }),
        ),
    )
    const result = await chatForProvider(
      'gemini',
      { apiKey: 'k', model: 'gemini-2.5-flash' },
      'sys',
      'hi',
    )
    expect(result).toEqual({ ok: true, content: 'hi there' })
  })

  it('deepseek and openai hit their fixed base URLs', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ choices: [{ message: { content: 'ok' } }] }))
    vi.stubGlobal('fetch', fetchMock)
    await chatForProvider('deepseek', { apiKey: 'k', model: 'deepseek-chat' }, 'sys', 'hi')
    expect(fetchMock).toHaveBeenCalledWith(
      'https://api.deepseek.com/v1/chat/completions',
      expect.anything(),
    )
  })

  it('custom: uses the configured base URL', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ choices: [{ message: { content: 'ok' } }] }))
    vi.stubGlobal('fetch', fetchMock)
    await chatForProvider(
      'custom',
      { apiKey: 'k', model: 'm', baseUrl: 'https://my-endpoint.example.com/v1' },
      'sys',
      'hi',
    )
    expect(fetchMock).toHaveBeenCalledWith(
      'https://my-endpoint.example.com/v1/chat/completions',
      expect.anything(),
    )
  })

  it('custom: rejects without a base URL, without calling fetch', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)
    const result = await chatForProvider('custom', { apiKey: 'k', model: 'm' }, 'sys', 'hi')
    expect(result).toEqual({ ok: false, error: 'A custom provider requires a Base URL' })
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it('genspark: routes by model prefix to the proxy endpoints', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ content: [{ type: 'text', text: 'ok' }] }))
    vi.stubGlobal('fetch', fetchMock)
    await chatForProvider('genspark', { apiKey: 'gsk-k', model: 'claude-opus-4-7' }, 'sys', 'hi')
    expect(fetchMock).toHaveBeenCalledWith(
      'https://www.genspark.ai/api/anthropic/v1/messages',
      expect.anything(),
    )
    fetchMock.mockResolvedValue(jsonResponse({ choices: [{ message: { content: 'ok' } }] }))
    await chatForProvider('genspark', { apiKey: 'gsk-k', model: 'gpt-5.2' }, 'sys', 'hi')
    expect(fetchMock).toHaveBeenLastCalledWith(
      'https://www.genspark.ai/api/llm_proxy/v1/chat/completions',
      expect.anything(),
    )
  })

  it('genspark: stamps X-Agent-Type; direct vendors do not get it', async () => {
    const fetchMock = vi
      .fn()
      .mockImplementation(async () => jsonResponse({ content: [{ type: 'text', text: 'ok' }] }))
    vi.stubGlobal('fetch', fetchMock)
    await chatForProvider('genspark', { apiKey: 'gsk-k', model: 'claude-opus-4-7' }, 'sys', 'hi')
    expect((fetchMock.mock.calls[0]![1].headers as Record<string, string>)['X-Agent-Type']).toBe(
      'genoffice',
    )
    fetchMock.mockClear()
    await chatForProvider('anthropic', { apiKey: 'k', model: 'claude-opus-4-7' }, 'sys', 'hi')
    expect(
      (fetchMock.mock.calls[0]![1].headers as Record<string, string>)['X-Agent-Type'],
    ).toBeUndefined()
  })

  it('treats an empty response body as an error', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(jsonResponse({ choices: [{ message: {} }] })))
    const result = await chatForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-4.1-mini' },
      'sys',
      'hi',
    )
    expect(result).toEqual({ ok: false, error: 'AI returned an empty response' })
  })
})
