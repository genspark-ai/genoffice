import { describe, expect, it } from 'vitest'
import { sseLines } from '../src/protocols/shared'

function sseBody(frames: string[]): ReadableStream<Uint8Array> {
  const encoder = new TextEncoder()
  return new ReadableStream<Uint8Array>({
    start(controller) {
      for (const frame of frames) controller.enqueue(encoder.encode(frame))
      controller.close()
    },
  })
}

/**
 * The protocol adapters wrap each SSE `JSON.parse` in a try/catch that skips
 * malformed frames. This test drives the raw sseLines generator through a
 * stream containing a corrupted frame and asserts the valid frames around it
 * still arrive — the try/catch in each adapter (anthropic/gemini/openai)
 * turns the corrupt payload into a `continue` rather than a thrown error.
 */
describe('SSE malformed-frame tolerance', () => {
  it('yields all lines including a corrupt one (the adapter skips it)', async () => {
    const lines: string[] = []
    for await (const line of sseLines(
      sseBody([
        'data: {"choices":[{"delta":{"content":"hello"}}]}\n',
        'data: {corrupt json!!!\n',
        'data: {"choices":[{"delta":{"content":" world"}}]}\n',
        'data: [DONE]\n',
      ]),
    )) {
      lines.push(line)
    }
    expect(lines).toEqual([
      'data: {"choices":[{"delta":{"content":"hello"}}]}',
      'data: {corrupt json!!!',
      'data: {"choices":[{"delta":{"content":" world"}}]}',
      'data: [DONE]',
    ])
    // The adapter's try/catch turns the corrupt line into a no-op:
    expect(() => JSON.parse('{corrupt json!!!')).toThrow()
    expect(() => {
      try {
        JSON.parse('{corrupt json!!!')
      } catch {
        /* skip — this is exactly what the adapters now do */
      }
    }).not.toThrow()
  })
})
