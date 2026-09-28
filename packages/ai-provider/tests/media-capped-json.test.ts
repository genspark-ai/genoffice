import { describe, expect, it } from 'vitest'

const MB = 1024 * 1024

function streamingResponse(totalBytes: number, chunkBytes = 64 * 1024): Response {
  let sent = 0
  const chunk = 'A'.repeat(chunkBytes)
  const enc = new TextEncoder()
  const stream = new ReadableStream({
    pull(ctrl) {
      if (sent >= totalBytes) {
        ctrl.close()
        return
      }
      ctrl.enqueue(enc.encode(chunk))
      sent += chunkBytes
    },
  })
  return new Response(stream, { headers: { 'content-type': 'application/json' } })
}

// readCappedJson is module-private; exercised through a media call below is
// heavy, so this test pins the cap behaviour of the shared reader it delegates to.
import { readCappedResponseText } from '../src/protocols/shared'

describe('media response cap', () => {
  it('refuses a gateway response larger than the shared cap before buffering it', async () => {
    const before = process.memoryUsage().rss
    await expect(readCappedResponseText(streamingResponse(64 * MB))).rejects.toThrow(
      /exceeded the .*-byte limit/,
    )
    expect(process.memoryUsage().rss - before).toBeLessThan(32 * MB)
  })

  it('still reads small JSON bodies through the same path', async () => {
    const resp = new Response('{"data":[]}', { headers: { 'content-type': 'application/json' } })
    expect(await readCappedResponseText(resp)).toBe('{"data":[]}')
  })
})
