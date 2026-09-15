/**
 * Markdown channels — single channel for reading markdown asset files.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, extname, join } from 'node:path'
import { DATA_DIR, registerHandle } from '../common/index'

const MARKDOWN_ASSET_DIR = join(DATA_DIR, 'markdown-assets')
const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
}

function safeAssetPath(src: string): string | null {
  const name = basename(src)
  if (name !== src || !IMAGE_MIME[extname(name).slice(1).toLowerCase()]) return null
  return join(MARKDOWN_ASSET_DIR, name)
}

export function registerMarkdownHandlers(): void {
  registerHandle('markdown:consume-pending', () => null)
  registerHandle('markdown:dirty-changed', () => ({ ok: true }))

  registerHandle('markdown:save-image', (_event: unknown, request: unknown) => {
    const value = request as { base64?: unknown; ext?: unknown } | null
    const ext = typeof value?.ext === 'string' ? value.ext.toLowerCase().replace(/^\./, '') : ''
    const base64 = typeof value?.base64 === 'string' ? value.base64 : ''
    if (!IMAGE_MIME[ext] || !base64) return null
    const bytes = Buffer.from(base64, 'base64')
    if (bytes.length === 0 || bytes.length > 20 * 1024 * 1024) return null
    mkdirSync(MARKDOWN_ASSET_DIR, { recursive: true })
    const name = `image-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.${ext}`
    writeFileSync(join(MARKDOWN_ASSET_DIR, name), bytes)
    return `markdown-assets/${name}`
  })

  registerHandle('markdown:read-image', (_event: unknown, src: unknown) => {
    if (typeof src !== 'string') return null
    const path = safeAssetPath(src.split('/').pop() ?? '')
    if (!path || !existsSync(path)) return null
    return { base64: readFileSync(path).toString('base64'), mime: IMAGE_MIME[extname(path).slice(1).toLowerCase()] }
  })

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
