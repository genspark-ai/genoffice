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
    expect(result.error).toMatch(/web page.*instead of an API response/)
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
    await chatForProvider('deepseek', { apiKey: 'k', model: 'deepseek-v4-pro' }, 'sys', 'hi')
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

  it('custom: omits Authorization when the key is empty for local servers', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ choices: [{ message: { content: 'ok' } }] }))
    vi.stubGlobal('fetch', fetchMock)
    await chatForProvider(
      'custom',
      { apiKey: '', model: 'llama3', baseUrl: 'http://localhost:11434/v1' },
      'sys',
      'hi',
    )
    const headers = fetchMock.mock.calls[0]![1].headers as Record<string, string>
    expect(headers.Authorization).toBeUndefined()
    fetchMock.mockClear()
    fetchMock.mockResolvedValue(jsonResponse({ choices: [{ message: { content: 'ok' } }] }))
    await chatForProvider(
      'custom',
      { apiKey: 'k', model: 'm', baseUrl: 'https://my-endpoint.example.com/v1' },
      'sys',
      'hi',
    )
    expect((fetchMock.mock.calls[0]![1].headers as Record<string, string>).Authorization).toBe(
      'Bearer k',
    )
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

  it('opencode: a one-shot call gets its own x-opencode-session', async () => {
    const fetchMock = vi
      .fn()
      .mockImplementation(async () => jsonResponse({ choices: [{ message: { content: 'ok' } }] }))
    vi.stubGlobal('fetch', fetchMock)
    await chatForProvider('opencode-go', { apiKey: 'k', model: 'kimi-k2.7-code' }, 'sys', 'hi')
    await chatForProvider('opencode-go', { apiKey: 'k', model: 'kimi-k2.7-code' }, 'sys', 'hi')
    const first = (fetchMock.mock.calls[0]![1].headers as Record<string, string>)[
      'x-opencode-session'
    ]
    const second = (fetchMock.mock.calls[1]![1].headers as Record<string, string>)[
      'x-opencode-session'
    ]
    expect(first).toMatch(/^[0-9a-f-]{36}$/)
    expect(second).not.toBe(first)
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

  it('anthropic: a 200 with an HTML body is an error, not a thrown SyntaxError', async () => {
    const html = '<!doctype html><html><body>gateway</body></html>'
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response(html, { status: 200 })))
    const result = await chatForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', 'hi')
    expect(result.ok).toBe(false)
    expect(result.error).toMatch(/non-JSON response/)
    expect(result.error).toMatch(/web page.*instead of an API response/)
  })

  it('openai: a 200 with an empty body is an error, not a thrown SyntaxError', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('', { status: 200 })))
    const result = await chatForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-4.1-mini' },
      'sys',
      'hi',
    )
    expect(result).toEqual({ ok: false, error: 'AI returned a non-JSON response: ' })
  })
})

describe('reasoning content normalization', () => {
  it('strips inline <think>…</think> tags from non-streaming chat replies', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        choices: [
          {
            message: {
              content: '<think>The user wants a short reply.</think>\n\nhi',
              role: 'assistant',
            },
          },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const result = await chatForProvider(
      'minimax',
      { apiKey: 'k', model: 'MiniMax-M3' },
      'sys',
      'say hi',
    )
    expect(result).toEqual({ ok: true, content: 'hi' })
  })

  it('strips multiple consecutive <think> blocks (re-entrant reasoning)', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        choices: [
          {
            message: {
              content:
                '<think>step 1</think>partial<think>step 2</think>final answer',
              role: 'assistant',
            },
          },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const result = await chatForProvider(
      'minimax',
      { apiKey: 'k', model: 'MiniMax-M3' },
      'sys',
      'go',
    )
    expect(result).toEqual({ ok: true, content: 'partialfinal answer' })
  })

  it('returns reasoning_content as a separate field when the server splits it out', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        choices: [
          {
            message: {
              content: 'final',
              reasoning_content: 'I thought about this',
              role: 'assistant',
            },
          },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const result = await chatForProvider(
      'deepseek',
      { apiKey: 'k', model: 'deepseek-v4-pro' },
      'sys',
      'go',
    )
    expect(result).toEqual({ ok: true, content: 'final', reasoning: 'I thought about this' })
  })

  it('reports reasoning-only responses as a failed reply (no empty success)', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        choices: [
          {
            message: {
              content: '',
              reasoning_content: 'the user said stop, so I stopped',
              role: 'assistant',
            },
          },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const result = await chatForProvider(
      'deepseek',
      { apiKey: 'k', model: 'deepseek-v4-pro' },
      'sys',
      'stop',
    )
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.error).toContain('AI responded with reasoning only')
    }
  })
})
