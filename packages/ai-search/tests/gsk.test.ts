import { describe, it, expect, afterEach, vi } from 'vitest'
import {
  gskChildEnv,
  setGskProxyUrl,
  parseGskOutput,
  parseGskWebSearch,
  parseGskImageSearch,
  parseGskGeneratedImage,
  parseGskPastProjects,
  extractGskText,
  parseToolCliNdjson,
  gskSlideGenerate,
  MAX_SLIDE_ARTIFACT_BYTES,
  MAX_TOOL_CLI_NDJSON_BYTES,
  summarizeGskFailure,
} from '../src/gsk'
import { ResponseTooLargeError } from '@genoffice/electron-utils/remote-image'
import { AiTimeoutError } from '@genoffice/ai-provider'

describe('parseGskOutput', () => {
  it('parses clean JSON', () => {
    expect(parseGskOutput('{"status":"ok"}')).toEqual({ status: 'ok' })
  })

  it('skips [INFO] noise lines before JSON', () => {
    const out = '[INFO] Calling /tools...\n[INFO] cache hit\n{"status":"ok","data":[1,2]}'
    expect(parseGskOutput(out)).toEqual({ status: 'ok', data: [1, 2] })
  })

  it('parses multi-line JSON after noise', () => {
    const out = '[INFO] x\n{\n "a": 1\n}'
    expect(parseGskOutput(out)).toEqual({ a: 1 })
  })

  it('skips trailing log lines after JSON', () => {
    const out = '{"status":"ok","data":[1,2]}\n[INFO] done in 120ms'
    expect(parseGskOutput(out)).toEqual({ status: 'ok', data: [1, 2] })
  })

  it('parses multi-line JSON surrounded by leading and trailing noise', () => {
    const out =
      '[INFO] Calling /tools...\n{\n "a": 1,\n "b": [1, 2]\n}\n[INFO] cache hit\n[INFO] done'
    expect(parseGskOutput(out)).toEqual({ a: 1, b: [1, 2] })
  })

  it('returns the root of a pretty-printed array followed by logs', () => {
    const out = '{\n  "items": [\n    { "id": 1 },\n    { "id": 2 }\n  ]\n}\n[INFO] done'
    expect(parseGskOutput(out)).toEqual({ items: [{ id: 1 }, { id: 2 }] })
  })

  it('throws when no JSON present', () => {
    expect(() => parseGskOutput('[INFO] nothing here')).toThrow()
  })

  it('recovers past a log line that only looks like an array opener', () => {
    expect(parseGskOutput('[INFO] progress {50%}\n{"status":"ok"}')).toEqual({ status: 'ok' })
  })

  it('takes the outer block, not the first inner one, of a large pretty payload', () => {
    const rows = Array.from({ length: 2000 }, (_, i) => `    { "id": ${i}, "t": "row ${i}" },`)
    const out = [
      '[INFO] Calling /tools...',
      '[INFO] cache hit',
      '{',
      '  "status": "ok",',
      '  "data": [',
      ...rows.slice(0, -1),
      rows.at(-1)!.replace(/,$/, ''),
      '  ]',
      '}',
      '[INFO] done in 900ms',
    ].join('\n')
    const parsed = parseGskOutput(out) as { status: string; data: unknown[] }
    expect(parsed.status).toBe('ok')
    expect(parsed.data).toHaveLength(2000)
    expect(parsed.data[0]).toEqual({ id: 0, t: 'row 0' })
  })

  it('locates a large payload in linear time', () => {
    const rows = Array.from({ length: 20_000 }, (_, i) => `  { "id": ${i} },`)
    const out = [
      '[INFO] starting',
      '{',
      '  "data": [',
      ...rows.slice(0, -1),
      rows.at(-1)!.replace(/,$/, ''),
      '  ]',
      '}',
      '[INFO] done',
    ].join('\n')
    const started = performance.now()
    const parsed = parseGskOutput(out) as { data: unknown[] }
    const elapsed = performance.now() - started
    expect(parsed.data).toHaveLength(20_000)
    // the previous nested slice-and-reparse scan needed minutes at this size
    expect(elapsed).toBeLessThan(5_000)
  })

  it('gives up on a hostile response in bounded time', () => {
    // Every line opens a block that never closes, so each candidate re-scans the
    // rest of the output: N lines of length L cost N x L. gsk output is
    // model-controlled and capped only by MAX_BUFFER, so the recovery scan has
    // to be bounded rather than quadratic.
    const out = Array.from(
      { length: 16_000 },
      (_, i) => `{"step":${i},"msg":"still rendering`,
    ).join('\n')
    const started = performance.now()
    expect(() => parseGskOutput(out)).toThrow(/No JSON found/)
    const elapsed = performance.now() - started
    // the unbounded scan needed ~25s at this size
    expect(elapsed).toBeLessThan(1_000)
  })

  it('still finds a payload buried behind unclosed log lines', () => {
    const out = [
      '{"msg":"still rendering',
      '{"msg":"still rendering',
      '{"status":"ok","data":{"n":7}}',
      '[INFO] done',
    ].join('\n')
    expect(parseGskOutput(out)).toEqual({ status: 'ok', data: { n: 7 } })
  })
})

describe('gskChildEnv', () => {
  afterEach(() => setGskProxyUrl(''))

  it('sets ELECTRON_RUN_AS_NODE and no proxy vars when no proxy is known', () => {
    const env = gskChildEnv({ PATH: '/bin' })
    expect(env.ELECTRON_RUN_AS_NODE).toBe('1')
    expect(env.NODE_USE_ENV_PROXY).toBeUndefined()
    expect(env.HTTPS_PROXY).toBeUndefined()
  })

  it('forwards the proxy registered by the main-process bootstrap', () => {
    setGskProxyUrl('http://127.0.0.1:7890')
    const env = gskChildEnv({ PATH: '/bin' })
    expect(env.NODE_USE_ENV_PROXY).toBe('1')
    expect(env.HTTPS_PROXY).toBe('http://127.0.0.1:7890')
    expect(env.HTTP_PROXY).toBe('http://127.0.0.1:7890')
  })

  it('falls back to inherited proxy env vars (terminal launch)', () => {
    const env = gskChildEnv({ https_proxy: 'http://10.0.0.1:8080' })
    expect(env.NODE_USE_ENV_PROXY).toBe('1')
    expect(env.HTTPS_PROXY).toBe('http://10.0.0.1:8080')
  })

  it('prefers the registered proxy over env vars', () => {
    setGskProxyUrl('http://127.0.0.1:7890')
    const env = gskChildEnv({ HTTPS_PROXY: 'http://10.0.0.1:8080' })
    expect(env.HTTPS_PROXY).toBe('http://127.0.0.1:7890')
  })

  it('scrubs lowercase/ALL_PROXY variants so they cannot override the selection', () => {
    setGskProxyUrl('http://127.0.0.1:7890')
    const env = gskChildEnv({
      https_proxy: 'socks5://127.0.0.1:1080',
      http_proxy: 'http://10.0.0.1:8080',
      all_proxy: 'socks5://127.0.0.1:1080',
    })
    expect(env.HTTPS_PROXY).toBe('http://127.0.0.1:7890')
    expect(env.https_proxy).toBeUndefined()
    expect(env.http_proxy).toBeUndefined()
    expect(env.all_proxy).toBeUndefined()
  })

  it('ignores SOCKS proxies (undici env proxy is http(s)-only)', () => {
    setGskProxyUrl('socks5://127.0.0.1:1080')
    const env = gskChildEnv({ ALL_PROXY: 'socks5://127.0.0.1:1080' })
    expect(env.NODE_USE_ENV_PROXY).toBeUndefined()
    expect(env.HTTPS_PROXY).toBeUndefined()
  })
})

describe('parseGskWebSearch', () => {
  it('maps organic_results and respects maxResults', () => {
    const raw = {
      status: 'ok',
      data: {
        organic_results: [
          { title: 'A', link: 'https://a.com', snippet: 'sa' },
          { title: 'B', link: 'https://b.com', snippet: 'sb' },
          { title: 'C', link: 'https://c.com', snippet: 'sc' },
        ],
      },
    }
    const r = parseGskWebSearch(raw, 2)
    expect(r.results).toEqual([
      { title: 'A', url: 'https://a.com', snippet: 'sa' },
      { title: 'B', url: 'https://b.com', snippet: 'sb' },
    ])
    expect(r.answer).toBeUndefined()
  })

  it('tolerates missing data', () => {
    expect(parseGskWebSearch({ status: 'ok' }, 5).results).toEqual([])
  })

  it('preserves full urls instead of clipping them to the snippet cap', () => {
    const long = `https://a.com/${'x'.repeat(3000)}`
    const raw = { data: { organic_results: [{ title: 'A', link: long, snippet: 's' }] } }
    const r = parseGskWebSearch(raw, 5)
    expect(r.results[0]!.url).toBe(long)
  })

  it('clamps maxResults and truncates long fields', () => {
    const big = 'x'.repeat(5000)
    const raw = {
      data: { organic_results: [{ title: big, link: 'https://a.com', snippet: big }] },
    }
    const r = parseGskWebSearch(raw, 1e9)
    expect(r.results).toHaveLength(1)
    expect(r.results[0]!.snippet.length).toBeLessThanOrEqual(2000)
    expect(parseGskWebSearch(raw, NaN).results).toHaveLength(1)
  })
})

describe('parseGskImageSearch', () => {
  it('maps image entries with numeric size coercion', () => {
    const raw = {
      status: 'ok',
      data: [
        {
          image_url: 'https://sspark.genspark.ai/img1',
          title: 'T1',
          source: 'Site',
          link: 'https://site.com/page',
          width: '1000',
          height: '688',
        },
      ],
    }
    const images = parseGskImageSearch(raw, 8)
    expect(images).toEqual([
      {
        title: 'T1',
        imageUrl: 'https://sspark.genspark.ai/img1',
        sourceUrl: 'https://site.com/page',
        source: 'Site',
        width: 1000,
        height: 688,
      },
    ])
  })

  it('filters copyright hosts and entries without url', () => {
    const raw = {
      data: [
        { image_url: 'https://media.gettyimages.com/x.jpg', title: 'g' },
        { title: 'no-url' },
        { image_url: 'https://ok.com/a.jpg', title: 'ok' },
      ],
    }
    const images = parseGskImageSearch(raw, 8)
    expect(images.map((i) => i.title)).toEqual(['ok'])
  })

  it('clamps maxResults like the web parser', () => {
    const entries = Array.from({ length: 25 }, (_, i) => ({
      image_url: `https://ok.com/${i}.jpg`,
      title: `t${i}`,
    }))
    const raw = { data: entries }
    expect(parseGskImageSearch(raw, NaN)).toHaveLength(6)
    expect(parseGskImageSearch(raw, 1e9)).toHaveLength(20)
    expect(parseGskImageSearch(raw, 0)).toHaveLength(1)
    expect(parseGskImageSearch(raw, -3)).toHaveLength(1)
  })

  it('keeps benign images whose path or query merely mentions a stock host', () => {
    const raw = {
      data: [
        { image_url: 'https://cdn.example.com/shutterstock-review.png', title: 'review' },
        { image_url: 'https://img.example.com/a.jpg?ref=shutterstock', title: 'query' },
        { image_url: 'https://media.gettyimages.com/x.jpg', title: 'blocked' },
      ],
    }
    const images = parseGskImageSearch(raw, 8)
    expect(images.map((i) => i.title)).toEqual(['review', 'query'])
  })
})

describe('parseGskPastProjects', () => {
  // real `gsk projects --artifact_types slides` shape (trimmed)
  const raw = {
    version: 1,
    status: 'ok',
    message: 'success',
    data: {
      projects: [
        {
          project_id: 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee',
          type: 'slides_agent_git',
          title: 'Product launch trailer presentation',
          ctime: '2026-07-29T07:09:43.706212',
        },
        {
          project_id: '12345678-90ab-4cde-8f01-234567890abc',
          type: 'slides_agent_git',
          title: 'Team collaboration deck request',
          ctime: '2026-07-23T08:39:04.464484',
        },
      ],
      total: 222,
      offset: 0,
      has_more: true,
      returned: 2,
    },
    session_state: {
      past_projects: {
        projects: [
          {
            project_id: 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee',
            project_url: '/agents?id=aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee',
            artifacts: [],
          },
        ],
      },
    },
  }

  it('maps projects, preferring session_state project_url and deriving the rest', () => {
    const page = parseGskPastProjects(raw)
    expect(page.total).toBe(222)
    expect(page.hasMore).toBe(true)
    expect(page.projects).toEqual([
      {
        projectId: 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee',
        type: 'slides_agent_git',
        title: 'Product launch trailer presentation',
        ctime: '2026-07-29T07:09:43.706212',
        projectUrl: '/agents?id=aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee',
      },
      {
        projectId: '12345678-90ab-4cde-8f01-234567890abc',
        type: 'slides_agent_git',
        title: 'Team collaboration deck request',
        ctime: '2026-07-23T08:39:04.464484',
        projectUrl: '/agents?id=12345678-90ab-4cde-8f01-234567890abc',
      },
    ])
  })

  it('skips entries without project_id and tolerates missing data', () => {
    const page = parseGskPastProjects({
      status: 'ok',
      data: { projects: [{ title: 'no id' }], has_more: false },
    })
    expect(page.projects).toEqual([])
    expect(page.total).toBe(0)
    expect(page.hasMore).toBe(false)
  })

  it('tolerates a completely empty response', () => {
    expect(parseGskPastProjects({ status: 'ok' })).toEqual({
      projects: [],
      total: 0,
      hasMore: false,
    })
  })
})

describe('parseGskGeneratedImage', () => {
  it('prefers no-watermark url', () => {
    const raw = {
      data: {
        generated_images: [
          {
            image_urls: ['https://cdn/wm.png'],
            image_urls_nowatermark: ['https://cdn/clean.png'],
            task_id: 't1',
          },
        ],
      },
    }
    expect(parseGskGeneratedImage(raw)).toEqual({ url: 'https://cdn/clean.png', taskId: 't1' })
  })

  it('falls back to image_urls, throws on empty', () => {
    expect(
      parseGskGeneratedImage({
        data: { generated_images: [{ image_urls: ['https://cdn/a.png'] }] },
      }).url,
    ).toBe('https://cdn/a.png')
    expect(() => parseGskGeneratedImage({ data: { generated_images: [] } })).toThrow()
  })
})

describe('extractGskText', () => {
  it('returns string data directly', () => {
    expect(extractGskText({ data: 'hello' })).toBe('hello')
  })

  it('picks known text fields', () => {
    expect(extractGskText({ data: { analysis: 'deep' } })).toBe('deep')
    expect(extractGskText({ data: { transcript: 'words' } })).toBe('words')
  })

  it('stringifies unknown shapes', () => {
    expect(extractGskText({ data: { foo: 1 } })).toBe('{"foo":1}')
  })
})

describe('gskSlideGenerate response caps', () => {
  const originalApiKey = process.env.GSK_API_KEY
  const slideResult = JSON.stringify({
    status: 'ok',
    data: { pptx_url: 'https://www.genspark.ai/api/files/deck.pptx', model: 'claude-opus-4-7' },
  })
  // a public address literal, so the download path needs no DNS in tests
  const downloadResult = JSON.stringify({
    status: 'ok',
    data: { download_url: 'https://8.8.8.8/deck.pptx' },
  })

  function stubSlideGenerate(artifact: () => Response, downloadResultBody = downloadResult) {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(ndjsonResponse(slideResult))
      .mockResolvedValueOnce(ndjsonResponse(downloadResultBody))
      .mockImplementationOnce(() => Promise.resolve(artifact()))
    vi.stubGlobal('fetch', fetchMock)
    return fetchMock
  }

  function ndjsonResponse(payload: string, contentLength?: number): Response {
    return new Response(payload, {
      status: 200,
      headers: {
        'content-type': 'application/x-ndjson',
        ...(contentLength === undefined ? {} : { 'content-length': String(contentLength) }),
      },
    })
  }

  afterEach(() => {
    vi.unstubAllGlobals()
    if (originalApiKey === undefined) delete process.env.GSK_API_KEY
    else process.env.GSK_API_KEY = originalApiKey
  })

  it('returns the downloaded slide bytes within both caps', async () => {
    process.env.GSK_API_KEY = 'test-key'
    stubSlideGenerate(() => new Response(new Uint8Array([1, 2, 3, 4])))
    await expect(gskSlideGenerate({ brief: 'a title slide' })).resolves.toEqual({
      bytes: new Uint8Array([1, 2, 3, 4]),
      model: 'claude-opus-4-7',
    })
  })

  it('refuses a tool_cli NDJSON body that declares more than the NDJSON cap', async () => {
    process.env.GSK_API_KEY = 'test-key'
    let cancelled = false
    const body = new ReadableStream<Uint8Array>({
      cancel() {
        cancelled = true
      },
    })
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        new Response(body, {
          status: 200,
          headers: {
            'content-type': 'application/x-ndjson',
            'content-length': String(MAX_TOOL_CLI_NDJSON_BYTES + 1),
          },
        }),
      ),
    )
    await expect(gskSlideGenerate({ brief: 'a title slide' })).rejects.toBeInstanceOf(
      ResponseTooLargeError,
    )
    expect(cancelled).toBe(true)
  })

  it('refuses a chunked tool_cli NDJSON body that streams past the NDJSON cap', async () => {
    process.env.GSK_API_KEY = 'test-key'
    const chunk = new Uint8Array(1024 * 1024)
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        for (let i = 0; i <= MAX_TOOL_CLI_NDJSON_BYTES / chunk.byteLength; i++) {
          controller.enqueue(chunk)
        }
        controller.close()
      },
    })
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(
          new Response(body, { status: 200, headers: { 'content-type': 'application/x-ndjson' } }),
        ),
    )
    await expect(gskSlideGenerate({ brief: 'a title slide' })).rejects.toThrow(
      /response larger than 8 MB/,
    )
  })

  it('refuses a slide artifact download larger than the artifact cap', async () => {
    process.env.GSK_API_KEY = 'test-key'
    stubSlideGenerate(
      () =>
        new Response(new Uint8Array([1, 2, 3]), {
          status: 200,
          headers: { 'content-length': String(MAX_SLIDE_ARTIFACT_BYTES + 1) },
        }),
    )
    await expect(gskSlideGenerate({ brief: 'a title slide' })).rejects.toBeInstanceOf(
      ResponseTooLargeError,
    )
  })

  // The download_url comes from the cloud response, so a compromised or spoofed
  // endpoint decides which host the main process dials. It must pass the same
  // SSRF gate as every other model-influenced download.
  it('never requests a cloud download URL that points at a private address', async () => {
    process.env.GSK_API_KEY = 'test-key'
    const fetchMock = stubSlideGenerate(
      () => new Response(new Uint8Array([1, 2, 3])),
      JSON.stringify({
        status: 'ok',
        data: { download_url: 'http://169.254.169.254/latest/meta-data/iam/security-credentials/' },
      }),
    )
    await expect(gskSlideGenerate({ brief: 'a title slide' })).rejects.toThrow(/blocked/i)
    const requested = fetchMock.mock.calls.map(([u]) => String(u))
    expect(requested.some((u) => u.includes('169.254.169.254'))).toBe(false)
  })

  it('revalidates every redirect hop of a cloud download URL', async () => {
    process.env.GSK_API_KEY = 'test-key'
    const fetchMock = stubSlideGenerate(
      () =>
        new Response(null, {
          status: 302,
          headers: { location: 'http://127.0.0.1:8080/internal.pptx' },
        }),
    )
    await expect(gskSlideGenerate({ brief: 'a title slide' })).rejects.toThrow(/blocked/i)
    const requested = fetchMock.mock.calls.map(([u]) => String(u))
    expect(requested.some((u) => u.includes('127.0.0.1'))).toBe(false)
  })
})

describe('parseToolCliNdjson', () => {
  it('skips heartbeat lines and returns the final status line', () => {
    const text =
      '{"version":1,"debug":true,"message":"Still processing... (5.0s)","heartbeat":1}\n' +
      '{"version":1,"debug":true,"message":"Still processing... (10.0s)","heartbeat":2}\n' +
      '{"version":1,"status":"ok","message":"success","data":{"pptx_url":"https://x/y","model":"claude-opus-4-7"}}'
    const r = parseToolCliNdjson(text)
    expect(r.status).toBe('ok')
    expect((r.data as { model: string }).model).toBe('claude-opus-4-7')
  })

  it('returns error result lines as-is', () => {
    const r = parseToolCliNdjson(
      '{"version":1,"status":"error","message":"deck_context must be an object","data":null}',
    )
    expect(r.status).toBe('error')
    expect(r.message).toMatch(/deck_context/)
  })

  it('throws when no result line exists', () => {
    expect(() => parseToolCliNdjson('{"heartbeat":1}\nnot json')).toThrow(/No result line/)
  })
})

describe('summarizeGskFailure', () => {
  // trimmed version of the gateway page Genspark serves for refused calls
  const HTML_403 = `<html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <title>Genspark</title>
  </head>
  <body>
    <div class="tt">Service unavailable. Please check your internet connection.</div>
    <form id="codeForm">
      <input type="text" id="codeInput" maxlength="8" />
      <button type="submit" class="submit-button">Submit</button>
    </form>
  </body>
  <script>function setCookie(event) { location.reload() }</script>
</html>`

  it('distills an HTML error page to its status and visible text', () => {
    const s = summarizeGskFailure(`HTTP 403: ${HTML_403}`)
    expect(s).toBe(
      'HTTP 403 (HTML error page): Service unavailable. Please check your internet connection.',
    )
    expect(s).not.toContain('<')
  })

  it('labels an HTML page that carries no status', () => {
    expect(summarizeGskFailure(HTML_403)).toBe(
      'an HTML error page: Service unavailable. Please check your internet connection.',
    )
  })

  it('keeps a short plain message as it is', () => {
    expect(summarizeGskFailure('HTTP 500: internal error')).toBe('HTTP 500: internal error')
    expect(summarizeGskFailure('deck_context must be an object')).toBe(
      'deck_context must be an object',
    )
  })

  it('keeps the [ERROR] lines, drops [INFO] chatter and crash noise', () => {
    const s = summarizeGskFailure(
      '[INFO] Uploading a.png...\n[INFO] Calling /file/upload_url...\n[ERROR] Failed to get upload URL: HTTP 403: Forbidden\nAssertion failed: !(handle->flags)',
    )
    expect(s).toBe('[ERROR] Failed to get upload URL: HTTP 403: Forbidden')
  })

  it('collapses multi-line plain text to one line', () => {
    expect(summarizeGskFailure('first line\nsecond line')).toBe('first line second line')
  })

  it('clips a long plain message', () => {
    const s = summarizeGskFailure('x'.repeat(400))
    expect(s.length).toBe(301)
    expect(s.endsWith('…')).toBe(true)
  })

  it('falls back on empty input and stringifies non-strings', () => {
    expect(summarizeGskFailure(undefined)).toBe('unknown error')
    expect(summarizeGskFailure(null, '')).toBe('')
    expect(summarizeGskFailure(404)).toBe('404')
  })

  it('never returns an empty string when the page has no readable text', () => {
    expect(summarizeGskFailure('HTTP 502: <html><body><script>x()</script></body></html>')).toBe(
      'HTTP 502 (HTML error page)',
    )
  })
})

describe('gskSlideGenerate cancellation and timeout', () => {
  const realFetch = globalThis.fetch
  afterEach(() => {
    globalThis.fetch = realFetch
    delete process.env.GSK_API_KEY
    vi.useRealTimers()
  })

  it('refuses an already-aborted signal before issuing the billed POST', async () => {
    process.env.GSK_API_KEY = 'test-key'
    let called = 0
    globalThis.fetch = vi.fn(async () => {
      called++
      return new Response('{}', { status: 500 })
    }) as unknown as typeof fetch

    const already = AbortSignal.abort()
    expect(already.aborted).toBe(true)
    // An abort listener added after the event fired is never invoked, so without an
    // explicit check the internal controller is never aborted and the request runs
    // to the full timeout — a billed slide_generate plus its artifact download.
    await expect(gskSlideGenerate({ brief: 'x', signal: already })).rejects.toThrow(/abort/i)
    expect(called).toBe(0)
  })

  it('surfaces its own deadline as AiTimeoutError, not a bare abort', async () => {
    process.env.GSK_API_KEY = 'test-key'
    // never settles: the call must end on the watchdog deadline, and must be typed so
    // the apps can map it to errorCode 'timeout' and show the localized message
    let sawAbort = false
    globalThis.fetch = vi.fn(
      (_url: unknown, init?: { signal?: AbortSignal }) =>
        new Promise<Response>((_resolve, reject) => {
          init?.signal?.addEventListener('abort', () => {
            sawAbort = true
            reject(new DOMException('The operation was aborted', 'AbortError'))
          })
        }),
    ) as unknown as typeof fetch

    vi.useFakeTimers()
    const call = gskSlideGenerate({ brief: 'x' }).then(
      () => null,
      (e: unknown) => e,
    )
    await vi.advanceTimersByTimeAsync(240_000)
    const err = await call
    vi.useRealTimers()

    expect(sawAbort).toBe(true)
    // AiTimeoutError, not a DOMException AbortError: the apps branch on this type
    expect(err).toBeInstanceOf(AiTimeoutError)
    expect((err as Error).message).toMatch(/timed out/i)
  })
})
