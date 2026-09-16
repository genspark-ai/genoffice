import { describe, expect, it } from 'vitest'
import { sseLines } from '../src/protocols/shared'

function sseBody(lines: string[]): ReadableStream<Uint8Array> {
  const encoder = new TextEncoder()
  return new ReadableStream<Uint8Array>({
    start(controller) {
      for (const line of lines) controller.enqueue(encoder.encode(line + '\n'))
      controller.close()
    },
  })
}

function hangingBody(): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(new TextEncoder().encode('data: first\n\n'))
      // Never closes — the stream stays open.
    },
  })
}

describe('sseLines', () => {
  it('yields SSE lines from a completing stream', async () => {
    const lines: string[] = []
    for await (const line of sseLines(sseBody(['data: a', 'data: b', '']))) {
      lines.push(line)
    }
    expect(lines).toEqual(['data: a', 'data: b', ''])
  })

  it('flushes a truncated multibyte tail instead of dropping it silently', async () => {
    const head = new TextEncoder().encode('data: \u4e2d')
    // A stream cut inside the next char (E4… with no completion bytes coming):
    // without a final decode() the buffered byte is discarded silently.
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(head)
        controller.enqueue(new Uint8Array([0xe4]))
        controller.close()
      },
    })
    const lines: string[] = []
    for await (const line of sseLines(body)) {
      lines.push(line)
    }
    expect(lines).toEqual(['data: \u4e2d\ufffd'])
  })

  it('releases the reader when the consumer abandons mid-stream', async () => {
    const body = hangingBody()
    const reader = body.getReader()
    reader.releaseLock() // give the lock back so sseLines can acquire it
    const gen = sseLines(body)
    const first = await gen.next()
    expect(first.value).toBe('data: first')
    // Abandon: .return() triggers the finally block, which must cancel and
    // release the reader so the underlying socket can be reused.
    await gen.return(undefined)
    // If the lock was not released this would throw.
    expect(() => body.getReader()).not.toThrow()
  })

  it('propagates consumer exceptions and still releases the reader', async () => {
    const body = hangingBody()
    const reader = body.getReader()
    reader.releaseLock()
    const gen = sseLines(body)
    await gen.next()
    // The consumer throws inside the for-await loop — JS calls gen.return(),
    // which runs the finally block.
    await expect(gen.throw(new Error('gateway error'))).rejects.toThrow('gateway error')
    expect(() => body.getReader()).not.toThrow()
  })
})
