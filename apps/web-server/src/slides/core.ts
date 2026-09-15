/**
 * Core slides lifecycle channels — new-blank, recent, open, open-path,
 * save, save-as, export-pdf, consume-pending-open, font-download/install,
 * insert-model3d. Persistence uses `slides-recent.json`.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, join } from 'node:path'
import { FILES_DIR, loadRecentSlides, registerHandle, saveRecentSlides } from '../common/index'

export function registerSlidesCoreHandlers(): void {
  registerHandle('slides:new-blank', async (_event: unknown, options: unknown) => {
    const opts = options as { pptx?: ArrayBuffer; path?: string } | undefined
    const id = `slide-${Date.now()}`
    const name = `演示文稿-${new Date().toLocaleDateString()}.pptx`
    const path = join(FILES_DIR, `${id}.pptx`)

    if (opts?.pptx) {
      writeFileSync(path, Buffer.from(opts.pptx))
    }

    const recent = loadRecentSlides()
    recent.unshift({ id, path, name, openedAt: Date.now() })
    saveRecentSlides(recent)

    return {
      path: '',
      slides: [
        {
          widthPx: 1280,
          heightPx: 720,
          scale: 1,
          background: { kind: 'solid', color: '#FFFFFF' },
          nodes: [],
        },
      ],
      size: { cx: 12192000, cy: 6858000 },
      defaultFont: 'Aptos',
    }
  })

  registerHandle('slides:recent', () => loadRecentSlides())

  registerHandle('slides:open', async (_event: unknown, options: unknown) => {
    const opts = options as { pptx?: ArrayBuffer; path?: string } | undefined
    const id = `slide-${Date.now()}`
    const name = opts?.path ? basename(opts.path) : `演示文稿-${Date.now()}.pptx`

    return { id, path: opts?.path || '', name }
  })

  registerHandle('slides:open-path', async (_event: unknown, filePath: unknown) => {
    if (!existsSync(filePath as string)) {
      throw new Error(`File not found: ${filePath}`)
    }

    const bytes = readFileSync(filePath as string)
    const name = basename(filePath as string)
    const id = `slide-${Date.now()}`

    const recent = loadRecentSlides()
    recent.unshift({ id, path: filePath as string, name, openedAt: Date.now() })
    saveRecentSlides(recent)

    return {
      id,
      path: filePath,
      name,
      bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    }
  })

  registerHandle(
    'slides:save',
    async (_event: unknown, id?: unknown, path?: unknown, data?: unknown) => {
      if (data && typeof path === 'string') {
        writeFileSync(path, Buffer.from(data as ArrayBuffer))
      }
      return { ok: true, path: typeof path === 'string' ? path : undefined }
    },
  )

  registerHandle(
    'slides:save-as',
    async (_event: unknown, defaultName?: unknown, data?: unknown) => {
      const id = `slide-${Date.now()}`
      const name = (typeof defaultName === 'string' && defaultName) || `演示文稿.pptx`
      const path = join(FILES_DIR, `${id}.pptx`)

      if (data) {
        writeFileSync(path, Buffer.from(data as ArrayBuffer))
      }

      return { id, path, name }
    },
  )

  registerHandle('slides:export-pdf', () => ({
    ok: true,
    message: '请使用浏览器的打印功能导出 PDF',
  }))

  registerHandle('slides:consume-pending-open', () => null)
  registerHandle('slides:autosave-pref', () => undefined)

  registerHandle('slides:font-download', () => ({
    ok: true,
    message: 'Web 版本不支持字体下载',
  }))

  registerHandle('slides:font-install-local', () => ({ ok: true }))

  registerHandle('slides:insert-model3d', () => ({ ok: true }))
}
