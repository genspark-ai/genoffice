/**
 * Slides file-pick / file-add channels — same surface as the generic
 * `files:*` set, namespaced for the slides renderer.
 */
import { existsSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { basename, extname, join } from 'node:path'
import { FILES_DIR, registerHandle } from '../common/index.js'

const IMAGE_MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp' }

export function registerSlidesFileHandlers(): void {
  registerHandle('slides:files-add', async (_event: unknown, args: unknown) => {
    const paths = Array.isArray(args) ? args.filter((p): p is string => typeof p === 'string') : []
    return paths.map((path) => {
      if (!existsSync(path)) return { path, ok: false, error: 'file not found' }
      const stat = statSync(path)
      return { path, ok: stat.isFile(), name: basename(path), sizeBytes: stat.size }
    })
  })

  registerHandle('slides:files-pick', () => ({
    canceled: false,
    filePaths: [],
    message: '请使用 Web File API',
  }))

  registerHandle('slides:files-read-image', async (_event: unknown, path: unknown) => {
    if (typeof path === 'string' && existsSync(path)) {
      const ext = extname(path).slice(1).toLowerCase()
      if (!IMAGE_MIME[ext]) return { ok: false, error: 'not an image' }
      const bytes = readFileSync(path)
      if (bytes.length > 5 * 1024 * 1024) return { ok: false, error: 'image is too large' }
      return { ok: true, base64: bytes.toString('base64'), mime: IMAGE_MIME[ext], name: basename(path) }
    }
    return { ok: false, error: 'file not found' }
  })

  registerHandle('slides:files-read', async (_event: unknown, path: unknown, offset: unknown, maxChars: unknown) => {
    if (typeof path !== 'string' || !existsSync(path)) return { ok: false, error: 'file not found' }
    const ext = extname(path).slice(1).toLowerCase()
    if (IMAGE_MIME[ext]) return { ok: false, error: 'image has no text' }
    const text = readFileSync(path, 'utf8')
    const start = Math.max(0, Number.isFinite(Number(offset)) ? Math.floor(Number(offset)) : 0)
    const size = Math.min(48000, Math.max(1, Number.isFinite(Number(maxChars)) ? Math.floor(Number(maxChars)) : 1))
    return { ok: true, name: basename(path), totalChars: text.length, offset: start, text: text.slice(start, start + size) }
  })

  registerHandle('slides:files-add-pasted-image', async (_event: unknown, data: unknown, ext: unknown) => {
    const cleanExt = typeof ext === 'string' ? ext.toLowerCase().replace(/^\./, '') : ''
    const bytes = data instanceof ArrayBuffer ? Buffer.from(data) : ArrayBuffer.isView(data) ? Buffer.from(data.buffer, data.byteOffset, data.byteLength) : null
    if (!bytes || !IMAGE_MIME[cleanExt] || bytes.length === 0 || bytes.length > 20 * 1024 * 1024) return { accepted: [], rejected: ['invalid image'] }
    const name = `pasted-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.${cleanExt}`
    const path = join(FILES_DIR, name)
    writeFileSync(path, bytes)
    return { accepted: [{ path, name, ext: cleanExt, sizeBytes: bytes.length }], rejected: [] }
  })

  registerHandle('slides:pick-export-dir', () => ({ path: '/tmp/exports' }))

  registerHandle('slides:pick-export-pdf-path', () => ({ path: '/tmp/exports/presentation.pdf' }))
}
