import { afterEach, describe, expect, it, vi } from 'vitest'
import { listCustomModels, listCustomModelsForIpc } from '../src/custom-models'
import { AI_DEFAULT_USER_AGENT, setRescueFetch } from '../src/fetch'

/**
 * `listCustomModels` is the only thing standing between the settings screen and
 * whatever a user-hosted server answers, so these cases pin three halves: the
 * real wire shapes it must read, the fact that nothing it can be handed makes it
 * throw, and the fact that it is a silent probe — no rescue fetch, no console
 * noise, no guessed default.
 *
 * The three success fixtures are verbatim bodies recorded from real servers —
 * Apple MLX (OMLX), Ollama 0.20.5 and LM Studio — and the 401 body is the real
 * refusal OMLX gives with no key. They differ in ways that matter: LM Studio
 * omits `created` and puts `data` before `object`, OMLX carries an extra
 * `max_model_len` that is `null` on one entry, and every server nests the ids
 * the same way regardless.
 *
 * They also show why no default is guessed: LM Studio's only listed model is an
 * embedding model, and OMLX lists an embedding model and a document converter
 * among its chat models.
 */

// ── recorded fixtures ─────────────────────────────────────

/** OMLX (Apple MLX server) at http://127.0.0.1:8800/v1/models, 16 models, keyed */
const OMLX_BODY =
  '{"object":"list","data":[{"id":"Qwen3.6-35B-A3B-8bit","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":262144},{"id":"Qwen3.8-27B-oQ4e-mtp","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":262144},{"id":"DeepSeek-V4-Flash-0731-MXFP4-MLX","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":262144},{"id":"GLM-4.7-Flash-8bit","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":131072},{"id":"MiniMax-M2.7-6bit","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":196608},{"id":"Qwen3-Embedding-0.6B-4bit-DWQ","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":32768},{"id":"Qwen3-VL-32B-Instruct-8bit","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":262144},{"id":"Qwen3-VL-32B-Instruct-bf16","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":262144},{"id":"Qwen3.5-122B-A10B-oQ4-fp16-mtp","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":262144},{"id":"Qwen3.6-27B-oQ4-mtp","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":160000},{"id":"Qwen3.6-27B-oQ8-mtp","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":160000},{"id":"Qwen3.8-Flash-Next-oQ4e-mtp","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":262144},{"id":"Step-3.7-Flash-4bit","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":262144},{"id":"Step-3.7-Flash-oQ6-MLX","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":262144},{"id":"gemma-4-31B-it-qat-oQ4e-mtp","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":262144},{"id":"MarkItDown","object":"model","created":1789386655,"owned_by":"omlx","max_model_len":null}]}'

/** Ollama 0.20.5 at http://127.0.0.1:11434/v1/models */
const OLLAMA_BODY =
  '{"object":"list","data":[{"id":"gemma3:4b","object":"model","created":1775135625,"owned_by":"library"}]}'

/** LM Studio at http://127.0.0.1:1234/v1/models — note: no `created`, and `data` comes first */
const LM_STUDIO_BODY =
  '{"data":[{"id":"text-embedding-nomic-embed-text-v1.5","object":"model","owned_by":"organization_owner"}],"object":"list"}'

/** OMLX's real refusal when the Authorization header is missing */
const OMLX_401_BODY =
  '{"error":{"message":"API key required","type":"authentication_error","param":null,"code":null}}'

const BASE_URL = 'http://127.0.0.1:8800/v1'

function jsonBody(body: string, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(body, {
    status,
    headers: { 'content-type': 'application/json', ...headers },
  })
}

function stubFetch(response: Response | (() => Promise<Response>)): ReturnType<typeof vi.fn> {
  const mock =
    typeof response === 'function' ? vi.fn(response) : vi.fn().mockResolvedValue(response)
  vi.stubGlobal('fetch', mock)
  return mock
}

function requestedUrl(mock: ReturnType<typeof vi.fn>): string {
  return mock.mock.calls[0]![0] as string
}

function sentHeaders(mock: ReturnType<typeof vi.fn>): Headers {
  return new Headers((mock.mock.calls[0]![1] as RequestInit).headers)
}

afterEach(() => {
  vi.unstubAllGlobals()
  vi.useRealTimers()
  vi.restoreAllMocks()
  setRescueFetch(null)
})

// ── the real servers ──────────────────────────────────────

describe('listCustomModels against recorded real servers', () => {
  it('reads all 16 OMLX models in server order and names no default', async () => {
    stubFetch(jsonBody(OMLX_BODY))
    const catalog = await listCustomModels(BASE_URL, 'a-key')
    expect(catalog.models).toHaveLength(16)
    expect(catalog.models[0]).toBe('Qwen3.6-35B-A3B-8bit')
    expect(catalog.models[15]).toBe('MarkItDown')
    expect(catalog.defaultModel).toBe('')
    // the entries carry `max_model_len` (null on one); extra fields must not matter
    expect(catalog.models).toContain('Qwen3-VL-32B-Instruct-bf16')
  })

  it('reads Ollama, whose ids carry a tag after a colon', async () => {
    stubFetch(jsonBody(OLLAMA_BODY))
    expect(await listCustomModels('http://127.0.0.1:11434/v1')).toEqual({
      models: ['gemma3:4b'],
      defaultModel: '',
    })
  })

  it('reads LM Studio, which omits `created` and orders its keys the other way', async () => {
    stubFetch(jsonBody(LM_STUDIO_BODY))
    expect(await listCustomModels('http://127.0.0.1:1234/v1')).toEqual({
      models: ['text-embedding-nomic-embed-text-v1.5'],
      defaultModel: '',
    })
  })

  it('never guesses a default, because the first entry is not a chat model', async () => {
    // LM Studio's single entry is an embedding model; OMLX's list holds an
    // embedding model and a document converter. `/models` has no `isDefault`
    // marker to read, unlike Codex's model/list, so nothing is picked.
    stubFetch(jsonBody(LM_STUDIO_BODY))
    expect((await listCustomModels('http://127.0.0.1:1234/v1')).defaultModel).toBe('')
    vi.unstubAllGlobals()
    stubFetch(jsonBody(OMLX_BODY))
    expect((await listCustomModels(BASE_URL, 'k')).defaultModel).toBe('')
  })
})

// ── the request it makes ──────────────────────────────────

describe('the request', () => {
  it('appends /models with exactly one slash, trailing slash or not', async () => {
    const bare = stubFetch(jsonBody(OMLX_BODY))
    await listCustomModels(BASE_URL)
    expect(requestedUrl(bare)).toBe('http://127.0.0.1:8800/v1/models')

    vi.unstubAllGlobals()
    const slashed = stubFetch(jsonBody(OMLX_BODY))
    await listCustomModels('http://127.0.0.1:8800/v1/')
    expect(requestedUrl(slashed)).toBe('http://127.0.0.1:8800/v1/models')

    vi.unstubAllGlobals()
    const padded = stubFetch(jsonBody(OMLX_BODY))
    await listCustomModels('  http://127.0.0.1:8800/v1///  ')
    expect(requestedUrl(padded)).toBe('http://127.0.0.1:8800/v1/models')
  })

  it('sends Bearer <key> when a key is given', async () => {
    const mock = stubFetch(jsonBody(OMLX_BODY))
    await listCustomModels(BASE_URL, 'sk-secret')
    expect(sentHeaders(mock).get('authorization')).toBe('Bearer sk-secret')
    expect(sentHeaders(mock).get('accept')).toBe('application/json')
  })

  it('omits Authorization entirely when there is no key — an empty Bearer is malformed', async () => {
    const absent = stubFetch(jsonBody(OLLAMA_BODY))
    await listCustomModels('http://127.0.0.1:11434/v1')
    expect(sentHeaders(absent).has('authorization')).toBe(false)

    vi.unstubAllGlobals()
    const blank = stubFetch(jsonBody(OLLAMA_BODY))
    await listCustomModels('http://127.0.0.1:11434/v1', '   ')
    expect(sentHeaders(blank).has('authorization')).toBe(false)
  })

  it('identifies itself, because a bare `node` user agent gets refused', async () => {
    // the probe skips aiFetch, but not the user agent aiFetch would have added:
    // gateways that police anonymous automation reject Node's default outright,
    // and the picker would then silently never appear
    const mock = stubFetch(jsonBody(OMLX_BODY))
    await listCustomModels(BASE_URL)
    expect(sentHeaders(mock).get('user-agent')).toBe(AI_DEFAULT_USER_AGENT)
  })

  it('does not call out at all for an empty base URL', async () => {
    const mock = stubFetch(jsonBody(OMLX_BODY))
    expect(await listCustomModels('   ')).toEqual({ models: [], defaultModel: '' })
    expect(mock).not.toHaveBeenCalled()
  })
})

// ── a silent probe, not a chat turn ───────────────────────

describe('the probe stays silent', () => {
  it('never retries through the rescue fetch and never logs', async () => {
    // aiFetch would do both on a network failure. A local server that is simply
    // not running is the normal case here, not an incident.
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined)
    const rescue = vi.fn().mockResolvedValue(jsonBody(OMLX_BODY))
    setRescueFetch(rescue)
    stubFetch(() => Promise.reject(new Error('fetch failed')))

    expect(await listCustomModels(BASE_URL)).toEqual({ models: [], defaultModel: '' })
    expect(rescue).not.toHaveBeenCalled()
    expect(warn).not.toHaveBeenCalled()
  })
})

// ── everything that can go wrong ──────────────────────────

describe('a server that cannot answer leaves the text box in place', () => {
  it('returns empty on the real 401 OMLX gives with no key', async () => {
    stubFetch(jsonBody(OMLX_401_BODY, 401))
    expect(await listCustomModels(BASE_URL)).toEqual({ models: [], defaultModel: '' })
  })

  it('returns empty on a non-JSON body (an HTML error page from a proxy)', async () => {
    stubFetch(new Response('<html><body>502 Bad Gateway</body></html>', { status: 200 }))
    expect(await listCustomModels(BASE_URL)).toEqual({ models: [], defaultModel: '' })
  })

  it('returns empty when the server answers with no models at all', async () => {
    stubFetch(jsonBody('{"object":"list","data":[]}'))
    expect(await listCustomModels(BASE_URL)).toEqual({ models: [], defaultModel: '' })
  })

  it('returns empty when `data` is not an array', async () => {
    stubFetch(jsonBody('{"models":["a","b"]}'))
    expect(await listCustomModels(BASE_URL)).toEqual({ models: [], defaultModel: '' })
  })

  it('returns empty when the connection fails outright', async () => {
    stubFetch(() => Promise.reject(new Error('fetch failed')))
    expect(await listCustomModels(BASE_URL)).toEqual({ models: [], defaultModel: '' })
  })

  it('gives up on a server that never answers, and aborts the request', async () => {
    vi.useFakeTimers()
    let aborted = false
    const mock = vi.fn(
      (_url: string, init: RequestInit) =>
        new Promise<Response>((_resolve, reject) => {
          init.signal?.addEventListener('abort', () => {
            aborted = true
            reject(new DOMException('The operation was aborted.', 'AbortError'))
          })
        }),
    )
    vi.stubGlobal('fetch', mock)
    const pending = listCustomModels(BASE_URL)
    await vi.advanceTimersByTimeAsync(5000)
    expect(await pending).toEqual({ models: [], defaultModel: '' })
    expect(aborted).toBe(true)
  })

  it('refuses a body the endpoint declares as oversized, without reading it', async () => {
    const body = jsonBody(OMLX_BODY, 200, { 'content-length': String(3 * 1024 * 1024) })
    const text = vi.spyOn(body, 'text')
    stubFetch(body)
    expect(await listCustomModels(BASE_URL)).toEqual({ models: [], defaultModel: '' })
    expect(text).not.toHaveBeenCalled()
    expect(body.bodyUsed).toBe(true) // the stream was cancelled, not left dangling
  })

  it('stops reading an oversized chunked body instead of buffering it', async () => {
    // A chunked answer declares no length, so the cap has to be enforced while
    // reading: the endpoint below never ends, and buffering it would run until
    // the watchdog fired, with everything it sent held in memory.
    const CHUNKS = 64 // 4 MB in 64 KB pieces, twice the cap
    const chunk = new TextEncoder().encode('x'.repeat(64 * 1024))
    let pulled = 0
    let cancelled = false
    const chunked = new ReadableStream<Uint8Array>({
      pull(controller) {
        // the ceiling is the test's own safety net, not the endpoint's promise:
        // a real one of these would keep going
        if (pulled >= CHUNKS) return controller.close()
        pulled += 1
        controller.enqueue(chunk)
      },
      cancel() {
        cancelled = true
      },
    })
    stubFetch(new Response(chunked, { status: 200 }))

    expect(await listCustomModels(BASE_URL)).toEqual({ models: [], defaultModel: '' })
    expect(cancelled).toBe(true) // stopped mid-stream, not read to the end
    // 2 MB of 64 KB chunks is 32 reads; reading all 64 means nothing capped it
    expect(pulled).toBeLessThan(40)
  })
})

describe('the ids it keeps', () => {
  it('drops blanks and non-strings, and dedupes while keeping first-seen order', async () => {
    stubFetch(
      jsonBody(
        '{"data":[{"id":"b"},{"id":"a"},{"id":"b"},{"id":""},{"id":"   "},{"id":7},{"name":"c"},null,{"id":"  a  "},{"id":"z"}]}',
      ),
    )
    expect(await listCustomModels(BASE_URL)).toEqual({
      models: ['b', 'a', 'z'],
      defaultModel: '',
    })
  })

  it('caps a pathological list at 2000', async () => {
    const data = Array.from({ length: 2400 }, (_unused, i) => ({ id: `m${i}` }))
    stubFetch(jsonBody(JSON.stringify({ data })))
    const catalog = await listCustomModels(BASE_URL)
    expect(catalog.models).toHaveLength(2000)
    expect(catalog.models[1999]).toBe('m1999')
  })
})

describe('listCustomModelsForIpc validates the payload', () => {
  it('passes a well-formed payload through', async () => {
    const mock = stubFetch(jsonBody(OMLX_BODY))
    const catalog = await listCustomModelsForIpc({ baseUrl: BASE_URL, apiKey: 'sk-secret' })
    expect(catalog.models).toHaveLength(16)
    expect(sentHeaders(mock).get('authorization')).toBe('Bearer sk-secret')
  })

  it('never throws on junk, and never sends a credential it was not given', async () => {
    const mock = stubFetch(jsonBody(OMLX_BODY))
    for (const junk of [undefined, null, 'string', 42, [], { baseUrl: 5 }, {}]) {
      expect(await listCustomModelsForIpc(junk)).toEqual({ models: [], defaultModel: '' })
    }
    expect(mock).not.toHaveBeenCalled()

    const keyless = await listCustomModelsForIpc({ baseUrl: BASE_URL, apiKey: 99 })
    expect(keyless.models).toHaveLength(16)
    expect(sentHeaders(mock).has('authorization')).toBe(false)
  })
})
