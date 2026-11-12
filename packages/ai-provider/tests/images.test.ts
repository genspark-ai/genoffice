import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentMessage } from '@genoffice/agent-core'
import { streamForProvider } from '../src/stream'
import { okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

const config = { apiKey: 'test-key', model: 'test-model' }

function callbacks() {
  return {
    signal: new AbortController().signal,
    onDelta: () => {},
    onToolCall: () => {},
  }
}

/** run one turn against a stubbed fetch and return the parsed request body */
async function requestBodyFor(
  provider: 'anthropic' | 'gemini',
  messages: AgentMessage[],
): Promise<any> {
  const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream([])))
  vi.stubGlobal('fetch', fetchMock)
  // the empty fixture stream legitimately rejects with "returned no content";
  // these tests only inspect the outgoing request body
  await streamForProvider(provider, config, 'sys', messages, [], 1024, callbacks()).catch(() => {})
  return JSON.parse(fetchMock.mock.calls[0][1].body as string)
}

const IMAGE = { base64: 'aGVsbG8=', mime: 'image/png' }

describe('anthropic user message with images', () => {
  it('sends a content array of text + image blocks', async () => {
    const body = await requestBodyFor('anthropic', [
      { role: 'user', text: 'look at this image', images: [IMAGE] },
    ])
    expect(body.messages[0]).toEqual({
      role: 'user',
      content: [
        { type: 'text', text: 'look at this image' },
        {
          type: 'image',
          source: { type: 'base64', media_type: 'image/png', data: 'aGVsbG8=' },
        },
      ],
    })
  })

  it('omits the text block when text is empty', async () => {
    const body = await requestBodyFor('anthropic', [{ role: 'user', text: '', images: [IMAGE] }])
    expect(body.messages[0].content).toHaveLength(1)
    expect(body.messages[0].content[0].type).toBe('image')
  })

  it('keeps plain string content when no images (existing behavior)', async () => {
    const body = await requestBodyFor('anthropic', [{ role: 'user', text: 'hi' }])
    expect(body.messages[0]).toEqual({ role: 'user', content: 'hi' })
  })
})

describe('gemini user message with images', () => {
  it('sends parts with text + inline_data', async () => {
    const body = await requestBodyFor('gemini', [
      { role: 'user', text: 'look at this image', images: [IMAGE] },
    ])
    expect(body.contents[0]).toEqual({
      role: 'user',
      parts: [
        { text: 'look at this image' },
        { inline_data: { mime_type: 'image/png', data: 'aGVsbG8=' } },
      ],
    })
  })

  it('keeps single text part when no images (existing behavior)', async () => {
    const body = await requestBodyFor('gemini', [{ role: 'user', text: 'hi' }])
    expect(body.contents[0]).toEqual({ role: 'user', parts: [{ text: 'hi' }] })
  })
})
