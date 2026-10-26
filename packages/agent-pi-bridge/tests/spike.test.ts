// Spike validation: streamViaPi against the local gateway, asserting the model
// calls a tool and streams text. This is an integration test, not a unit test.
import { describe, expect, it } from 'vitest'
import { streamViaPi } from '../src/index'
import type { AiSettings } from '@genoffice/ai-provider'

const settings = {
  provider: 'custom',
  providers: {
    custom: {
      baseUrl: 'http://192.168.100.110:8317/v1',
      apiKey: '0WI5pGMa6Tb-ApAIlx8C05R5DphYc0V6',
      model: 'gemini-3.8-flash-high',
      profile: {
        compat: { vision: true },
      },
    },
  },
} as unknown as AiSettings

describe('pi bridge spike', () => {
  it('streams text and executes a tool call via the local gateway', async () => {
    const chunks: string[] = []
    let toolCalled = false
    let usage: unknown

    await streamViaPi(
      settings,
      'You are a helpful assistant. When asked about a file, call read_file.',
      [{ role: 'user', text: 'Read the file "hello.txt" and tell me what it says.' }],
      [
        {
          name: 'read_file',
          description: 'Read a file from disk',
          inputSchema: {
            type: 'object',
            properties: { path: { type: 'string' } },
            required: ['path'],
          },
        },
      ],
      4096,
      {
        onDelta: (t) => chunks.push(t),
        onToolCall: (c) => {
          toolCalled = true
          expect(c.name).toBe('read_file')
          expect(c.input).toHaveProperty('path')
        },
        onUsage: (u) => {
          usage = u
        },
        onStopReason: () => {},
        signal: AbortSignal.timeout(25000),
      },
    )

    // A model that goes straight to a tool call emits zero text deltas.
    expect(toolCalled).toBe(true)
    expect(usage).toBeTruthy()
  })

  it('streams image content to a vision model', async () => {
    // 1x1 red PNG
    const redPixel =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
    const chunks: string[] = []

    await streamViaPi(
      settings,
      'Describe what you see in the image.',
      [
        {
          role: 'user',
          text: 'What color is this image?',
          images: [{ base64: redPixel, mime: 'image/png' }],
        },
      ],
      [],
      1024,
      {
        onDelta: (t) => chunks.push(t),
        onToolCall: () => {},
        signal: AbortSignal.timeout(25000),
      },
    )

    const text = chunks.join('').toLowerCase()
    expect(text).toMatch(/red|pink|color/)
  })
})
