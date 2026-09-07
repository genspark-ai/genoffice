/**
 * Markdown channels — single channel for reading markdown asset files.
 */
import { existsSync, readFileSync } from 'node:fs'
import { registerHandle } from '../common/index.js'

export function registerMarkdownHandlers(): void {
  registerHandle('md-asset', async (_event: unknown, args: unknown) => {
    const { path, type } = args as { path: string; type: string }
    if (type === 'read' && existsSync(path)) {
      return { content: readFileSync(path, 'utf-8') }
    }
    return null
  })
}
