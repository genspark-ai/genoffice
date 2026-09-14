/**
 * Markdown channels — single channel for reading markdown asset files.
 */
import { existsSync, readFileSync } from 'node:fs'
import { registerHandle } from '../common/index.js'

export function registerMarkdownHandlers(): void {
  registerHandle('markdown:consume-pending', () => null)
  registerHandle('markdown:dirty-changed', () => ({ ok: true }))

  registerHandle('markdown:read-file', async (_event: unknown, filePath: unknown) => {
    if (typeof filePath !== 'string' || !existsSync(filePath)) {
      throw new Error(`File not found: ${String(filePath)}`)
    }
    return readFileSync(filePath, 'utf8')
  })

  registerHandle('md-asset', async (_event: unknown, path?: unknown, type?: unknown) => {
    if (type === 'read' && typeof path === 'string' && existsSync(path)) {
      return { content: readFileSync(path, 'utf-8') }
    }
    return null
  })
}
