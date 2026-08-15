import { describe, it, expect, afterEach } from 'vitest'
import {
  gskChildEnv,
  setGskProxyUrl,
  parseGskOutput,
  parseGskWebSearch,
  parseGskImageSearch,
  parseGskGeneratedImage,
  parseGskConvertResult,
  parseGskPastProjects,
  extractGskText,
  parseToolCliNdjson,
} from '../src/gsk'

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

  it('throws when no JSON present', () => {
    expect(() => parseGskOutput('[INFO] nothing here')).toThrow()
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

describe('parseGskConvertResult', () => {
  it('extracts the markdown link from the result text', () => {
    const raw = {
      status: 'ok',
      data: {
        result:
          'Conversion complete. Download links:\n[report.docx](https://www.genspark.ai/api/files/s/JmS2WJHv)\n',
      },
    }
    expect(parseGskConvertResult(raw)).toBe('https://www.genspark.ai/api/files/s/JmS2WJHv')
  })

  it('falls back to a bare URL without markdown', () => {
    const raw = { status: 'ok', data: { result: 'Done: https://example.com/f.docx' } }
    expect(parseGskConvertResult(raw)).toBe('https://example.com/f.docx')
  })

  it('throws when the result has no link', () => {
    expect(() => parseGskConvertResult({ status: 'ok', data: { result: 'no link' } })).toThrow()
    expect(() => parseGskConvertResult({ status: 'ok' })).toThrow()
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
