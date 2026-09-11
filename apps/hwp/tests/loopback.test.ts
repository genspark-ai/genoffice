import { request } from 'node:http'
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { startHwpLoopback } from '../src/main/loopback'

function rawStatus(origin: string, path: string): Promise<number> {
  const url = new URL(origin)
  return new Promise((resolve, reject) => {
    const req = request(
      { hostname: url.hostname, port: url.port, path, method: 'GET' },
      (res) => {
        res.resume()
        res.on('end', () => resolve(res.statusCode ?? 0))
      },
    )
    req.on('error', reject)
    req.end()
  })
}

describe('startHwpLoopback', () => {
  it('serves the renderer index over http', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'hwp-loopback-'))
    await writeFile(join(dir, 'index.html'), '<html>ok</html>')
    const { origin, close } = await startHwpLoopback(dir)
    try {
      expect(origin.startsWith('http://127.0.0.1:')).toBe(true)
      const res = await fetch(origin)
      expect(res.ok).toBe(true)
      expect(await res.text()).toBe('<html>ok</html>')
    } finally {
      await close()
    }
  })

  it('returns 404 when the file and fallback are missing', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'hwp-loopback-empty-'))
    const { origin, close } = await startHwpLoopback(dir)
    try {
      const res = await fetch(origin)
      expect(res.status).toBe(404)
    } finally {
      await close()
    }
  })

  it('serves the studio print surface next to the embed', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'hwp-loopback-print-'))
    await writeFile(join(dir, 'index.html'), '<html>host</html>')
    await mkdir(join(dir, 'rhwp'))
    await writeFile(join(dir, 'rhwp', 'print.html'), '<html>print</html>')
    const { origin, close } = await startHwpLoopback(dir)
    try {
      const res = await fetch(new URL('rhwp/print.html', origin))
      expect(res.ok).toBe(true)
      expect(res.headers.get('content-type')).toContain('text/html')
      expect(await res.text()).toBe('<html>print</html>')
    } finally {
      await close()
    }
  })

  it('refuses path traversal and stock PWA files', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'hwp-loopback-guard-'))
    await writeFile(join(dir, 'index.html'), '<html>ok</html>')
    await writeFile(join(dir, 'sw.js'), 'stolen')
    const { origin, close } = await startHwpLoopback(dir)
    try {
      expect(await rawStatus(origin, '/../../../etc/passwd')).toBe(403)
      const pwa = await fetch(new URL('sw.js', origin))
      expect(pwa.status).toBe(404)
      expect(await pwa.text()).not.toContain('stolen')
    } finally {
      await close()
    }
  })
})
