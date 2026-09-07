/**
 * Slides file-pick / file-add channels — same surface as the generic
 * `files:*` set, namespaced for the slides renderer.
 */
import { existsSync, readFileSync } from 'node:fs'
import { basename } from 'node:path'
import { registerHandle } from '../common/index.js'

export function registerSlidesFileHandlers(): void {
  registerHandle('slides:files-add', async (_event: unknown, args: unknown) => {
    const paths = (args as string[]) || []
    return paths.map(p => ({ path: p, ok: true }))
  })

  registerHandle('slides:files-pick', () => ({
    canceled: false,
    filePaths: [],
    message: '请使用 Web File API',
  }))

  registerHandle('slides:files-read-image', async (_event: unknown, path: unknown) => {
    if (existsSync(path as string)) {
      const bytes = readFileSync(path as string)
      return { base64: bytes.toString('base64'), name: basename(path as string) }
    }
    return null
  })

  registerHandle('slides:pick-export-dir', () => ({ path: '/tmp/exports' }))

  registerHandle('slides:pick-export-pdf-path', () => ({ path: '/tmp/exports/presentation.pdf' }))
}
