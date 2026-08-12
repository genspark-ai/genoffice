import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentMessage } from '@genoffice/agent-core'
import { spawn } from 'node:child_process'
import {
  OpencodeServerError,
  disposeOpencodeServer,
  ensureOpencodeServer,
  opencodePromptText,
  parseOpencodeModel,
  streamOpencode,
} from '../src/opencode'
import { errorResponse, jsonResponse, okResponse, sseStream } from './test-utils'

vi.mock('node:child_process', () => ({ spawn: vi.fn() }))
const mockSpawn = vi.mocked(spawn)

/** fake child that records the 'error' handler so tests can fire it */
function fakeChild() {
  const handlers = new Map<string, (...args: never[]) => void>()
  return {
    killed: false,
    once(event: string, cb: (...args: never[]) => void) {
      handlers.set(event, cb)
      return this
    },
    kill() {
      this.killed = true
    },
    fire(event: string, ...args: never[]) {
      handlers.get(event)?.(...args)
    },
  }
}

afterEach(() => {
  vi.unstubAllGlobals()
  mockSpawn.mockReset()
  disposeOpencodeServer()
})

describe('parseOpencodeModel', () => {
  it('returns null for blank specs (server default model)', () => {
    expect(parseOpencodeModel('')).toBeNull()
    expect(parseOpencodeModel('   ')).toBeNull()
  })

  it('treats a bare model id as the opencode provider default', () => {
    expect(parseOpencodeModel('gpt-5')).toEqual({ id: 'gpt-5', providerID: 'opencode' })
  })

  it('splits a provider/model spec on the first slash', () => {
    expect(parseOpencodeModel('openrouter/anthropic/claude-3.5-sonnet')).toEqual({
      id: 'anthropic/claude-3.5-sonnet',
      providerID: 'openrouter',
    })
  })
})

describe('opencodePromptText', () => {
  it('returns just the system prompt when there are no messages', () => {
    expect(opencodePromptText('be concise', [])).toBe('be concise')
    expect(opencodePromptText('  ', [])).toBe('')
  })

  it('serializes a user message verbatim', () => {
    const text = opencodePromptText('sys', [{ role: 'user', text: 'hello' }])
    expect(text).toContain('## User\nhello')
    expect(text).toContain('# Conversation history')
    expect(text).toContain('# Task')
  })

  it('lists tool calls an assistant turn made', () => {
    const messages: AgentMessage[] = [
      {
        role: 'assistant',
        text: 'let me check',
        toolCalls: [{ name: 'web_search', input: { q: 'x' }, id: 't1' }],
      },
    ]
    const text = opencodePromptText('sys', messages)
    expect(text).toContain('## Assistant\nlet me check')
    expect(text).toContain('web_search({"q":"x"})')
  })

  it('formats tool results with error markers', () => {
    const messages: AgentMessage[] = [
      {
        role: 'tool',
        results: [
          { name: 'web_search', output: 'ok', isError: false, id: 't1' },
          { name: 'fetch', output: 'boom', isError: true, id: 't2' },
        ],
      },
    ]
    const text = opencodePromptText('sys', messages)
    expect(text).toContain('- web_search: ok')
    expect(text).toContain('- fetch (error): boom')
  })
})

describe('ensureOpencodeServer', () => {
  it('returns immediately when a server already answers', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(new Response('').body!)))
    await expect(ensureOpencodeServer('http://127.0.0.1:3456')).resolves.toBeUndefined()
    expect(mockSpawn).not.toHaveBeenCalled()
  })

  it('spawns "opencode serve" and waits until the server answers', async () => {
    let reachable = false
    vi.stubGlobal(
      'fetch',
      vi.fn().mockImplementation(async () => {
        if (!reachable) throw new Error('connection refused')
        return okResponse(new Response('').body!)
      }),
    )
    const child = fakeChild()
    mockSpawn.mockReturnValue(child as never)

    const run = ensureOpencodeServer('http://127.0.0.1:3456')
    // let the first reachability probe fail, then bring the server up
    await new Promise((r) => setTimeout(r, 0))
    reachable = true
    await expect(run).resolves.toBeUndefined()

    expect(mockSpawn).toHaveBeenCalledWith(
      expect.stringMatching(/^opencode(\.cmd)?$/),
      ['serve', '--port', '3456', '--hostname', '127.0.0.1'],
      expect.anything(),
    )
  })

  it('reports a spawn failure as OpencodeServerError', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockRejectedValue(new Error('connection refused')),
    )
    const child = fakeChild()
    mockSpawn.mockReturnValue(child as never)
    // fire the spawn 'error' event synchronously so the poll loop sees it
    const originalOnce = child.once.bind(child)
    child.once = ((event: string, cb: (...a: never[]) => void) => {
      const self = originalOnce(event, cb)
      if (event === 'error') cb(new Error('ENOENT') as never)
      return self
    }) as typeof child.once

    await expect(ensureOpencodeServer('http://127.0.0.1:3456')).rejects.toBeInstanceOf(
      OpencodeServerError,
    )
    // a failed spawn leaves no shared server behind: the next call must try again
    disposeOpencodeServer()
    await expect(
      ensureOpencodeServer('http://127.0.0.1:3456'),
    ).rejects.toThrow(/Could not start the opencode CLI/)
  })

  it('disposeOpencodeServer kills the spawned child', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(new Response('').body!)))
    await ensureOpencodeServer('http://127.0.0.1:3456')
    expect(mockSpawn).not.toHaveBeenCalled()
  })
})

describe('streamOpencode: opencode 1.18.16 API shapes', () => {
  // opencode 1.18.16 returns the created session as {"data":{"id":"ses_..."}}
  // (not a top-level "sessionID"), and every SSE event carries its payload in
  // a nested "data" object. These tests pin the shapes the integration must
  // tolerate; the real server was probed while writing them.

  function collector() {
    const deltas: string[] = []
    return {
      deltas,
      cb: {
        signal: new AbortController().signal,
        onDelta: (text: string) => deltas.push(text),
        onToolCall: () => {},
        onStopReason: () => {},
      },
    }
  }

  it('creates a session from data.id and streams text from nested event data', async () => {
    const fetchMock = vi.fn()
    // 0) ensureOpencodeServer health probe (GET /api/session?limit=1)
    fetchMock
      .mockResolvedValueOnce(okResponse(new Response('').body!))
      // 1) session create → {"data":{"id":...}}
      .mockResolvedValueOnce(jsonResponse({ data: { id: 'ses_abc123' } }))
      // 2) prompt → accepted
      .mockResolvedValueOnce(jsonResponse({ data: { admittedSeq: 1 } }))
      // 3) event stream → text + finish carried in nested "data"
      .mockResolvedValueOnce(
        okResponse(
          sseStream([
            `data: ${JSON.stringify({ type: 'session.next.text.ended', data: { text: 'Hello from opencode' } })}`,
            `data: ${JSON.stringify({ type: 'session.next.step.ended', data: { finish: 'stop' } })}`,
          ]),
        ),
      )
    vi.stubGlobal('fetch', fetchMock)

    const { deltas, cb } = collector()
    await streamOpencode(
      { apiKey: '', model: 'opencode/big-pickle', baseUrl: 'http://127.0.0.1:3456' },
      'sys',
      [{ role: 'user', text: 'hi' }],
      [],
      100,
      cb,
    )

    expect(deltas.join('')).toBe('Hello from opencode')
    // the session create request carried the model spec
    const createBody = JSON.parse(String(fetchMock.mock.calls[1][1].body))
    expect(createBody.model).toEqual({ id: 'big-pickle', providerID: 'opencode' })
  })

  it('reports a step error carried in nested event data', async () => {
    const fetchMock = vi.fn()
    fetchMock
      .mockResolvedValueOnce(okResponse(new Response('').body!))
      .mockResolvedValueOnce(jsonResponse({ data: { id: 'ses_abc123' } }))
      .mockResolvedValueOnce(jsonResponse({ data: { admittedSeq: 1 } }))
      .mockResolvedValueOnce(
        okResponse(
          sseStream([
            `data: ${JSON.stringify({ type: 'session.next.error', data: { message: 'rate limited' } })}`,
          ]),
        ),
      )
    vi.stubGlobal('fetch', fetchMock)

    const { cb } = collector()
    await expect(
      streamOpencode(
        { apiKey: '', model: 'opencode/big-pickle', baseUrl: 'http://127.0.0.1:3456' },
        'sys',
        [{ role: 'user', text: 'hi' }],
        [],
        100,
        cb,
      ),
    ).rejects.toThrow(/rate limited/)
  })
})

describe('streamOpencode: empty model auto-resolves a usable server model', () => {
  // A blank model field must not create a model-less session: `opencode serve`
  // without a configured default model accepts the prompt and then sits idle
  // forever (the request eventually dies on GenOffice's idle watchdog). When
  // the user leaves the field blank, resolve the model from the server itself.
  // Server probed while writing these tests: /config has "model": null here,
  // and /api/model advertises OpenCode Zen hosted models (apiKey "public").

  function collector() {
    const deltas: string[] = []
    return {
      deltas,
      cb: {
        signal: new AbortController().signal,
        onDelta: (text: string) => deltas.push(text),
        onToolCall: () => {},
        onStopReason: () => {},
      },
    }
  }

  function okEventStream(): Response {
    return okResponse(
      sseStream([
        `data: ${JSON.stringify({ type: 'session.next.text.ended', data: { text: 'hi from auto model' } })}`,
        `data: ${JSON.stringify({ type: 'session.next.step.ended', data: { finish: 'stop' } })}`,
      ]),
    )
  }

  it('picks the first /api/model entry when the server has no default model', async () => {
    const fetchMock = vi.fn()
    // 0) ensureOpencodeServer health probe
    fetchMock
      .mockResolvedValueOnce(okResponse(new Response('').body!))
      // 1) /config → no default model configured
      .mockResolvedValueOnce(jsonResponse({ model: null }))
      // 2) /api/model → server's advertised models, first one wins
      .mockResolvedValueOnce(
        jsonResponse({
          data: [
            { id: 'nemotron-3.5-lightning-free', providerID: 'opencode', name: 'Nemotron 3.5 Lightning Free' },
            { id: 'big-pickle', providerID: 'opencode', name: 'Big Pickle' },
          ],
        }),
      )
      // 3) session create → data.id shape
      .mockResolvedValueOnce(jsonResponse({ data: { id: 'ses_abc123' } }))
      // 4) prompt → accepted
      .mockResolvedValueOnce(jsonResponse({ data: { admittedSeq: 1 } }))
      // 5) event stream → text + finish
      .mockResolvedValueOnce(okEventStream())
    vi.stubGlobal('fetch', fetchMock)

    const { deltas, cb } = collector()
    await streamOpencode(
      { apiKey: '', model: '', baseUrl: 'http://127.0.0.1:3456' },
      'sys',
      [{ role: 'user', text: 'hi' }],
      [],
      100,
      cb,
    )

    expect(deltas.join('')).toBe('hi from auto model')
    // resolution hit /config then /api/model before creating the session
    expect(fetchMock.mock.calls[1][0]).toBe('http://127.0.0.1:3456/config')
    expect(fetchMock.mock.calls[2][0]).toBe('http://127.0.0.1:3456/api/model')
    const createBody = JSON.parse(String(fetchMock.mock.calls[3][1].body))
    expect(createBody.model).toEqual({
      id: 'nemotron-3.5-lightning-free',
      providerID: 'opencode',
    })
  })

  it('omits the model when the server config has a default (server default wins)', async () => {
    const fetchMock = vi.fn()
    fetchMock
      .mockResolvedValueOnce(okResponse(new Response('').body!))
      // /config → a default model IS configured; keep the model-less create
      .mockResolvedValueOnce(jsonResponse({ model: 'anthropic/claude-sonnet-4-5' }))
      .mockResolvedValueOnce(jsonResponse({ data: { id: 'ses_abc123' } }))
      .mockResolvedValueOnce(jsonResponse({ data: { admittedSeq: 1 } }))
      .mockResolvedValueOnce(okEventStream())
    vi.stubGlobal('fetch', fetchMock)

    const { cb } = collector()
    await streamOpencode(
      { apiKey: '', model: '', baseUrl: 'http://127.0.0.1:3456' },
      'sys',
      [{ role: 'user', text: 'hi' }],
      [],
      100,
      cb,
    )

    // resolution stopped at /config — no /api/model call
    expect(fetchMock.mock.calls.length).toBe(5)
    expect(fetchMock.mock.calls[1][0]).toBe('http://127.0.0.1:3456/config')
    const createBody = JSON.parse(String(fetchMock.mock.calls[2][1].body))
    expect(createBody.model).toBeUndefined()
  })

  it('falls back to a model-less session when the config query fails', async () => {
    const fetchMock = vi.fn()
    fetchMock
      .mockResolvedValueOnce(okResponse(new Response('').body!))
      // /config → 404 (older server or one that hides the endpoint)
      .mockResolvedValueOnce(errorResponse(404, 'not found'))
      // /api/model → also unavailable
      .mockResolvedValueOnce(errorResponse(404, 'not found'))
      .mockResolvedValueOnce(jsonResponse({ data: { id: 'ses_abc123' } }))
      .mockResolvedValueOnce(jsonResponse({ data: { admittedSeq: 1 } }))
      .mockResolvedValueOnce(okEventStream())
    vi.stubGlobal('fetch', fetchMock)

    const { cb } = collector()
    await streamOpencode(
      { apiKey: '', model: '', baseUrl: 'http://127.0.0.1:3456' },
      'sys',
      [{ role: 'user', text: 'hi' }],
      [],
      100,
      cb,
    )

    const createBody = JSON.parse(String(fetchMock.mock.calls[3][1].body))
    expect(createBody.model).toBeUndefined()
  })
})
