/**
 * HTML module channels — file read/save and pending consume for the
 * minimal HTML editor at apps/html/. Mirrors the shape of the other
 * module handlers (docs/sheets/slides/markdown/pdf).
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { registerHandle } from '../common/index.js'

export function registerHtmlHandlers(): void {
  registerHandle('html:consume-pending', () => null)

  registerHandle('html:read-file', async (_event: unknown, filePath: unknown) => {
    if (typeof filePath !== 'string' || !existsSync(filePath)) {
      throw new Error(`File not found: ${String(filePath)}`)
    }
    return readFileSync(filePath, 'utf8')
  })

  registerHandle('html:save-file', async (_event: unknown, path: unknown, content: unknown) => {
    if (typeof path !== 'string' || !path) {
      return { ok: false, error: 'invalid path' }
    }
    try {
      writeFileSync(path, typeof content === 'string' ? content : '', 'utf8')
      return { ok: true, path }
    } catch (e) {
      return { ok: false, error: (e as Error).message }
    }
  })
}
