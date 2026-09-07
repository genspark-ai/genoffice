import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { startHwpLoopback } from '../src/main/loopback'

describe('startHwpLoopback', () => {
  it('serves the renderer index over http', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'hwp-loopback-'))
    await writeFile(join(dir, 'index.html'), '<html>ok</html>')
    const origin = await startHwpLoopback(dir)
    expect(origin.startsWith('http://127.0.0.1:')).toBe(true)
    const res = await fetch(origin)
    expect(res.ok).toBe(true)
    expect(await res.text()).toBe('<html>ok</html>')
  })
})
