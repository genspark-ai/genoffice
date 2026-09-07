import { createReadStream } from 'node:fs'
import { createServer } from 'node:http'
import { join, resolve } from 'node:path'
import {
  cacheControlFor,
  existingFile,
  isBlockedStudioAsset,
  mimeFor,
  safePathUnder,
} from '../shared/static-serve'

export interface HwpLoopback {
  origin: string
  close(): Promise<void>
}

/**
 * rhwp-studio's embed SDK rejects file:/opaque origins. Packaged Hangul tabs
 * therefore load the renderer + vendored studio over loopback HTTP.
 */
export function startHwpLoopback(root: string): Promise<HwpLoopback> {
  const base = resolve(root)
  const server = createServer((req, res) => {
    const urlPath = req.url ?? '/'
    if (isBlockedStudioAsset(urlPath)) {
      res.statusCode = 404
      res.end()
      return
    }
    const wanted = safePathUnder(base, urlPath)
    if (!wanted) {
      res.statusCode = 403
      res.end()
      return
    }
    const fallback = urlPath.split('?')[0].startsWith('/rhwp/')
      ? join(base, 'rhwp', 'index.html')
      : join(base, 'index.html')
    const file = existingFile(wanted) ?? existingFile(fallback)
    if (!file) {
      res.statusCode = 404
      res.end()
      return
    }
    res.setHeader('Content-Type', mimeFor(file))
    res.setHeader('Cache-Control', cacheControlFor(file))
    createReadStream(file).pipe(res)
  })
  return new Promise((resolveAddr, reject) => {
    server.once('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const addr = server.address()
      if (!addr || typeof addr === 'string') {
        reject(new Error('hwp: loopback bind failed'))
        return
      }
      resolveAddr({
        origin: `http://127.0.0.1:${addr.port}/`,
        close: () =>
          new Promise((resolveClose, rejectClose) => {
            server.close((err) => (err ? rejectClose(err) : resolveClose()))
          }),
      })
    })
  })
}
