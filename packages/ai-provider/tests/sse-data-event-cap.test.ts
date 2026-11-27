import { describe, expect, it } from 'vitest'
import { MAX_SSE_DATA_EVENT_CHARS, sseDataEvents } from '../src/protocols/shared'

function sseBody(frames: string[]): ReadableStream<Uint8Array> {
  const encoder = new TextEncoder()
  return new ReadableStream<Uint8Array>({
    start(controller) {
      for (const frame of frames) controller.enqueue(encoder.encode(frame))
      controller.close()
    },
  })
}

describe('sseDataEvents fragment cap', () => {
  it('fails closed when a gateway streams fragments past the accumulated cap', async () => {
    // Each line is a JSON opener that never completes: parse fails on its own,
    // so the event is held fragment after fragment with no blank line to flush
    // it — exactly the buffer that must not grow without bound.
    const fragment = `{"a":"${'x'.repeat(4000)}`
    const perLine = fragment.length + 1 // + the '\n' separator used when joining
    const lineCount = Math.ceil((MAX_SSE_DATA_EVENT_CHARS + 1) / perLine) + 5
    const lines = Array.from({ length: lineCount }, () => `data: ${fragment}\n`)
    expect(lines.length * perLine).toBeGreaterThan(MAX_SSE_DATA_EVENT_CHARS)
    let cancelled = false
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        const encoder = new TextEncoder()
        for (const line of lines) controller.enqueue(encoder.encode(line))
        controller.close()
      },
      cancel() {
        cancelled = true
      },
    })
    await expect(async () => {
      for await (const _event of sseDataEvents(body)) {
        // must throw before the stream ends, not buffer the whole flood
      }
    }).rejects.toThrow(/buffer limit/)
    // The throw aborts the turn: sseLines cancels the underlying stream and
    // releases the reader on the way out.
    expect(cancelled).toBe(true)
    expect(() => body.getReader()).not.toThrow()
  })

  it('still joins and dispatches a normal-sized split data event', async () => {
    const events: unknown[] = []
    for await (const event of sseDataEvents(
      sseBody([
        'data: {"choices":[{"delta":{"content":\n',
        'data: "hello"}}]}\n',
        'data: [DONE]\n',
      ]),
    )) {
      events.push(event)
    }
    expect(events).toEqual([
      {
        raw: '{"choices":[{"delta":{"content":\n"hello"}}]}',
        json: { choices: [{ delta: { content: 'hello' } }] },
      },
      { raw: '[DONE]', json: undefined },
    ])
  })
})
