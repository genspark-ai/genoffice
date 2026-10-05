import { afterEach, describe, expect, it, vi } from 'vitest'
import { analyzeMediaWithProvider, generateImageWithProvider } from '../src/media-protocols'

/**
 * A corporate proxy or vendor gateway can answer a media request with HTTP 200 and
 * its own web shell instead of JSON. Every other failure mode in this package reports
 * itself as `Image generation failed: …` / `Media analysis failed: …`; a raw
 * `resp.json()` on that body throws a bare `SyntaxError: Unexpected token '<'` and
 * surfaces it to the user instead.
 */

afterEach(() => {
  vi.unstubAllGlobals()
})

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0])

const HTML_SHELL =
  '<!doctype html>\n<html><head><title>403 Forbidden</title></head><body>Access denied</body></html>'

function htmlShell(): Response {
  return new Response(HTML_SHELL, { status: 200, headers: { 'content-type': 'text/html' } })
}

function jsonResponse(body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  })
}

/** the message a caller sees; a resolved call fails this test */
async function messageOf(promise: Promise<unknown>): Promise<string> {
  const err = await promise.then(
    () => {
      throw new Error('expected the media call to fail')
    },
    (e: unknown) => e,
  )
  expect(err).toBeInstanceOf(Error)
  return (err as Error).message
}

/** the package's own failure message — the labelled one, never a leaked SyntaxError */
function expectSoftFailure(message: string, label: string): void {
  expect(message.startsWith(label)).toBe(true)
  expect(message).toContain('the service returned a web page')
  expect(message).not.toMatch(/SyntaxError|Unexpected token/)
}

describe('a 200 media response whose body is not JSON', () => {
  it('fails OpenAI image generation and edits with the image-failure message', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => htmlShell()),
    )
    const config = { apiKey: 'sk-1', imageModel: 'gpt-image-2', analysisModel: '' }
    expectSoftFailure(
      await messageOf(generateImageWithProvider('openai', config, { prompt: 'a cat' })),
      'Image generation failed:',
    )
    expectSoftFailure(
      await messageOf(
        generateImageWithProvider('openai', config, {
          prompt: 'a cat',
          references: [{ bytes: PNG, mime: 'image/png' }],
        }),
      ),
      'Image edit failed:',
    )
  })

  it('fails DashScope/Qwen-Image with the image-failure message', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => htmlShell()),
    )
    expectSoftFailure(
      await messageOf(
        generateImageWithProvider(
          'qwen',
          {
            apiKey: 'k',
            baseUrl: 'https://dashscope-intl.aliyuncs.com/compatible-mode/v1',
            imageModel: 'qwen-image-plus',
            analysisModel: '',
          },
          { prompt: 'a fox' },
        ),
      ),
      'Image generation failed:',
    )
  })

  it('fails MiniMax with the image-failure message', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => htmlShell()),
    )
    expectSoftFailure(
      await messageOf(
        generateImageWithProvider(
          'minimax',
          { apiKey: 'k', imageModel: 'image-01', analysisModel: '' },
          { prompt: 'a fox' },
        ),
      ),
      'Image generation failed:',
    )
  })

  it('fails Gemini Imagen predict and native image output with the image-failure message', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => htmlShell()),
    )
    const config = { apiKey: 'AIza', imageModel: '', analysisModel: '' }
    expectSoftFailure(
      await messageOf(
        generateImageWithProvider(
          'gemini',
          { ...config, imageModel: 'imagen-4.0-generate-001' },
          {
            prompt: 'a dog',
          },
        ),
      ),
      'Image generation failed:',
    )
    expectSoftFailure(
      await messageOf(
        generateImageWithProvider(
          'gemini',
          { ...config, imageModel: 'gemini-3.1-flash-image' },
          { prompt: 'a dog' },
        ),
      ),
      'Image generation failed:',
    )
  })

  it('fails the OpenAI-compatible and Gemini analysis paths with the analysis-failure message', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => htmlShell()),
    )
    const media = [{ bytes: PNG, mime: 'image/png', name: 'logo.png' }]
    expectSoftFailure(
      await messageOf(
        analyzeMediaWithProvider(
          'openai',
          { apiKey: 'sk', imageModel: '', analysisModel: 'gpt-5.6-luna' },
          { media, requirements: 'describe' },
        ),
      ),
      'Media analysis failed:',
    )
    expectSoftFailure(
      await messageOf(
        analyzeMediaWithProvider(
          'gemini',
          { apiKey: 'AIza', imageModel: '', analysisModel: 'gemini-3.7-flash' },
          { media, requirements: 'describe' },
        ),
      ),
      'Media analysis failed:',
    )
  })

  it('fails the Gemini Files upload with the upload-failure message', async () => {
    const uploadUrl = 'https://generativelanguage.googleapis.com/upload/v1beta/files?upload_id=1'
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string) =>
        url === uploadUrl
          ? htmlShell()
          : new Response('{}', {
              status: 200,
              headers: { 'content-type': 'application/json', 'x-goog-upload-url': uploadUrl },
            }),
      ),
    )
    // one byte over the inline limit, so the media goes through the Files API
    const tooBig = { bytes: new Uint8Array(18 * 1024 * 1024 + 1), mime: 'video/mp4' }
    expectSoftFailure(
      await messageOf(
        analyzeMediaWithProvider(
          'gemini',
          { apiKey: 'AIza', imageModel: '', analysisModel: 'gemini-3.7-flash' },
          { media: [tooBig], requirements: 'summarize' },
        ),
      ),
      'Media upload failed:',
    )
  })

  it('fails the Gemini upload poll with the upload-failure message', async () => {
    const uploadUrl = 'https://generativelanguage.googleapis.com/upload/v1beta/files?upload_id=1'
    const start = new Response('{}', {
      status: 200,
      headers: { 'content-type': 'application/json', 'x-goog-upload-url': uploadUrl },
    })
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string) => {
        if (url === uploadUrl)
          return jsonResponse({ file: { state: 'PROCESSING', name: 'files/x' } })
        if (url.includes('upload/v1beta/files')) return start
        return htmlShell()
      }),
    )
    const tooBig = { bytes: new Uint8Array(18 * 1024 * 1024 + 1), mime: 'video/mp4' }
    expectSoftFailure(
      await messageOf(
        analyzeMediaWithProvider(
          'gemini',
          { apiKey: 'AIza', imageModel: '', analysisModel: 'gemini-3.7-flash' },
          { media: [tooBig], requirements: 'summarize' },
        ),
      ),
      'Media upload failed:',
    )
  })
})
