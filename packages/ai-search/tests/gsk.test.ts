import { describe, it, expect } from 'vitest'
import {
  parseGskOutput,
  parseGskWebSearch,
  parseGskImageSearch,
  parseGskGeneratedImage,
  parseGskConvertResult,
  extractGskText,
  parseToolCliNdjson,
  parseGskLoginLine,
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

// sample lines mirror @genspark/cli login: info/errors on stderr with
// [INFO]/[ERROR] prefixes, success JSON envelope on stdout
describe('parseGskLoginLine', () => {
  it('extracts the auth URL', () => {
    expect(
      parseGskLoginLine('[INFO] Login URL: https://www.genspark.ai/cli-auth?code=abc'),
    ).toEqual({ kind: 'url', url: 'https://www.genspark.ai/cli-auth?code=abc' })
  })

  it('extracts expires_in from the waiting line', () => {
    expect(
      parseGskLoginLine(
        '[INFO] Waiting for authorization (expires in 300s, press Ctrl+C to cancel)...',
      ),
    ).toEqual({ kind: 'expires', expiresInSec: 300 })
  })

  it('detects success from the stderr info line and the stdout json envelope', () => {
    expect(parseGskLoginLine('[INFO] Login successful! API key saved.')).toEqual({
      kind: 'success',
    })
    expect(parseGskLoginLine('  "message": "Login successful",')).toEqual({ kind: 'success' })
  })

  it('classifies expiry and timeout as expired', () => {
    expect(parseGskLoginLine('[ERROR] Authorization expired. Please try again.')).toEqual({
      kind: 'error',
      reason: 'expired',
      message: 'Authorization expired. Please try again.',
    })
    expect(parseGskLoginLine('[ERROR] Authorization timed out. Please try again.')).toMatchObject({
      kind: 'error',
      reason: 'expired',
    })
  })

  it('classifies device-code request failures as network', () => {
    expect(parseGskLoginLine('[ERROR] fetch failed')).toMatchObject({
      kind: 'error',
      reason: 'network',
    })
    expect(parseGskLoginLine('[ERROR] HTTP 502: Bad Gateway')).toMatchObject({
      kind: 'error',
      reason: 'network',
    })
  })

  it('passes other errors through as-is', () => {
    expect(parseGskLoginLine('[ERROR] auth_url host does not match baseUrl host')).toEqual({
      kind: 'error',
      reason: 'other',
      message: 'auth_url host does not match baseUrl host',
    })
  })

  it('ignores noise lines', () => {
    expect(parseGskLoginLine('[INFO] Requesting device code...')).toBeNull()
    expect(parseGskLoginLine('[INFO] Opening browser for login...')).toBeNull()
    expect(
      parseGskLoginLine('[INFO] Still waiting for authorization... (295s remaining)'),
    ).toBeNull()
    expect(parseGskLoginLine('')).toBeNull()
  })
})
