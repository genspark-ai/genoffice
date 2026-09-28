import { describe, expect, it } from 'vitest'
import { readImageSource } from '../src/formats/image-source'

const ctx = { cwd: process.cwd(), env: process.env } as never

describe('data: URL image cap', () => {
  it('refuses an oversized data URL before decoding it', async () => {
    // 70 MB of base64 encodes ~52 MB — over the 50 MB budget the remote
    // branch enforces; on main this allocated the full payload first
    // (measured: a 400 MB URL peaked at ~1 GB RSS).
    const big = `data:image/png;base64,${'A'.repeat(70 * 1024 * 1024)}`
    // The rejection itself must not allocate: on main the decoded payload
    // (~52 MB → ~1 GB peak with a 400 MB URL) appeared before any check.
    const before = process.memoryUsage().rss
    expect(await readImageSource(big, ctx)).toBeNull()
    // The input string itself accounts for ~70 MB; anything beyond a small
    // margin would be the decode we are guarding against.
    expect(process.memoryUsage().rss - before).toBeLessThan(80 * 1024 * 1024)
  })

  it('keeps small data URLs working', async () => {
    // 1x1 transparent PNG
    const png =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
    const src = await readImageSource(png, ctx)
    expect(src).not.toBeNull()
  })
})
