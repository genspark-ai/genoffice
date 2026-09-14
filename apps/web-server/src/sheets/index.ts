/**
 * Sheets IPC channels — workbook open/has-queued/consume-new-blank.
 * Persistence uses `sheets-recent.json` for the recent-files list.
 */
import { existsSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { basename, extname, join } from 'node:path'
import { FILES_DIR, loadRecentSheets, registerHandle, saveRecentSheets } from '../common/index.js'

export function registerSheetsHandlers(): void {
  registerHandle('sheets:new-blank', async (_event: unknown, options: unknown) => {
    const opts = options as { xlsx?: ArrayBuffer; path?: string } | undefined
    const id = `sheet-${Date.now()}`
    const name = `表格-${new Date().toLocaleDateString()}.xlsx`
    const path = join(FILES_DIR, `${id}.xlsx`)

    if (opts?.xlsx) {
      writeFileSync(path, Buffer.from(opts.xlsx))
    }

    const recent = loadRecentSheets()
    recent.unshift({ id, path, name, openedAt: Date.now() })
    saveRecentSheets(recent)

    return { id, path, name }
  })

  registerHandle('sheets:has-queued-workbook', () => false)

  registerHandle('workbook:open-path', async (_event: unknown, filePath: unknown) => {
    if (!existsSync(filePath as string)) {
      throw new Error(`File not found: ${filePath}`)
    }

    const bytes = readFileSync(filePath as string)
    const name = basename(filePath as string)
    const id = `sheet-${Date.now()}`

    const recent = loadRecentSheets()
    recent.unshift({ id, path: filePath as string, name, openedAt: Date.now() })
    saveRecentSheets(recent)

    return {
      id,
      path: filePath,
      name,
      bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    }
  })

  registerHandle('sheets:consume-new-blank', () => ({ ok: true }))
  registerHandle('workbook:pending-edits', () => ({ ok: true }))

  const imageMime: Record<string, string> = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    webp: 'image/webp',
  }
  registerHandle('sheets:files-add', (_event: unknown, paths: unknown) => {
    const values = Array.isArray(paths) ? paths.filter((path): path is string => typeof path === 'string') : []
    return values.map((path) => {
      if (!existsSync(path)) return { path, ok: false, error: 'file not found' }
      const info = statSync(path)
      return { path, ok: info.isFile(), name: basename(path), sizeBytes: info.size }
    })
  })
  registerHandle('sheets:files-read', (_event: unknown, path: unknown, offset: unknown, maxChars: unknown) => {
    if (typeof path !== 'string' || !existsSync(path)) return { ok: false, error: 'file not found' }
    const ext = extname(path).slice(1).toLowerCase()
    if (imageMime[ext]) return { ok: false, error: 'image has no text' }
    const text = readFileSync(path, 'utf8')
    const start = Math.max(0, Number.isFinite(Number(offset)) ? Math.floor(Number(offset)) : 0)
    const size = Math.min(48000, Math.max(1, Number.isFinite(Number(maxChars)) ? Math.floor(Number(maxChars)) : 1))
    return { ok: true, name: basename(path), totalChars: text.length, offset: start, text: text.slice(start, start + size) }
  })
  registerHandle('sheets:files-read-image', (_event: unknown, path: unknown) => {
    if (typeof path !== 'string' || !existsSync(path)) return { ok: false, error: 'file not found' }
    const ext = extname(path).slice(1).toLowerCase()
    const mime = imageMime[ext]
    if (!mime) return { ok: false, error: 'not an image' }
    const bytes = readFileSync(path)
    if (bytes.length > 5 * 1024 * 1024) return { ok: false, error: 'image is too large' }
    return { ok: true, base64: bytes.toString('base64'), mime }
  })
  registerHandle('sheets:files-add-pasted-image', (_event: unknown, data: unknown, ext: unknown) => {
    const cleanExt = typeof ext === 'string' ? ext.toLowerCase().replace(/^\./, '') : ''
    const bytes = data instanceof ArrayBuffer ? Buffer.from(data) : ArrayBuffer.isView(data) ? Buffer.from(data.buffer, data.byteOffset, data.byteLength) : null
    if (!bytes || !imageMime[cleanExt] || bytes.length === 0 || bytes.length > 20 * 1024 * 1024) return { accepted: [], rejected: ['invalid image'] }
    const name = `pasted-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.${cleanExt}`
    const path = join(FILES_DIR, name)
    writeFileSync(path, bytes)
    return { accepted: [{ path, name, ext: cleanExt, sizeBytes: bytes.length }], rejected: [] }
  })

  registerHandle('workbook:open-for-merge', (_event: unknown, paths: unknown) => {
    if (!Array.isArray(paths) || paths.length === 0 || paths.length > 20) throw new Error('merge sources must be 1-20 files')
    return paths.map((path) => {
      if (typeof path !== 'string' || !existsSync(path)) throw new Error(`Merge source not found: ${String(path)}`)
      return { path, name: basename(path) }
    })
  })
}
