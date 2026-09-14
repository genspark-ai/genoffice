/**
 * Docs IPC channels — recent, font-metrics, pick-image, settings,
 * open/save/read/print. Persistence is in-memory + `docs-recent.json`; the
 * `DOCS_RECENT` map referenced by `home:recents` is the same singleton
 * declared in `common/state.ts`.
 */
import { existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { basename, extname, isAbsolute, join, relative, resolve, sep } from 'node:path'
import {
  DATA_DIR,
  FILES_DIR,
  WEB_TEMP_ROOT,
  loadProjects,
  loadRecentDocs,
  registerHandle,
  saveProjects,
  saveRecentDocs,
} from '../common/index.js'

const MAX_PASTED_IMAGE_BYTES = 20 * 1024 * 1024
const closeState = {
  check: null as { dirty: boolean; autoSave: boolean; filePath: string | null } | null,
  saveOk: null as boolean | null,
}

function isWithin(root: string, target: string): boolean {
  const relativePath = relative(resolve(root), resolve(target))
  return (
    relativePath === '' ||
    (!relativePath.startsWith(`..${sep}`) && relativePath !== '..' && !isAbsolute(relativePath))
  )
}

function isManagedDocPath(filePath: string): boolean {
  return /\.docx$/i.test(filePath) && [DATA_DIR, WEB_TEMP_ROOT].some((root) => isWithin(root, filePath))
}

function bytesFrom(value: unknown): Buffer | null {
  if (value instanceof ArrayBuffer) return Buffer.from(value)
  if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  return null
}

export function registerDocsHandlers(): void {
  registerHandle('docs:view-menu-state', () => ({ ok: true }))

  registerHandle('docs:discard-password-intents', (_event: unknown, throughRevision: unknown) => ({
    ok:
      typeof throughRevision === 'number' &&
      Number.isSafeInteger(throughRevision) &&
      throughRevision >= 0,
  }))

  registerHandle('docs:recent', () => loadRecentDocs())

  registerHandle('docs:font-metrics', (_event: unknown, family: unknown) => ({
    family: family || 'sans-serif',
    ascent: 0.8,
    descent: 0.2,
    lineGap: 0.1,
    unitsPerEm: 1000,
  }))

  registerHandle('docs:pick-image', () => ({
    canceled: false,
    dataUrl: null,
    message: '请使用 Web File API 在前端选择图片',
  }))

  registerHandle('docs:get-settings', () => ({
    language: 'zh-CN',
    spellCheck: true,
    autoSave: true,
    autoSaveInterval: 30000,
    fontSize: 14,
    fontFamily: 'sans-serif',
  }))

  registerHandle('docs:save-settings', () => ({ ok: true }))

  registerHandle('docs:close-check-result', (_event: unknown, state: unknown) => {
    const value = state as { dirty?: unknown; autoSave?: unknown; filePath?: unknown } | null
    closeState.check = {
      dirty: value?.dirty === true,
      autoSave: value?.autoSave === true,
      filePath: typeof value?.filePath === 'string' ? value.filePath : null,
    }
  })

  registerHandle('docs:close-save-result', (_event: unknown, ok: unknown) => {
    closeState.saveOk = ok === true
  })

  registerHandle('docs:open', async (_event: unknown, options: unknown) => {
    const opts = options as { docx?: ArrayBuffer; path?: string } | undefined
    const id = `doc-${Date.now()}`

    if (opts?.docx) {
      const name = `文档-${new Date().toLocaleDateString()}.docx`
      const path = join(FILES_DIR, `${id}.docx`)
      writeFileSync(path, Buffer.from(opts.docx))

      const recent = loadRecentDocs()
      recent.unshift({ id, path, name, openedAt: Date.now(), modified: false })
      saveRecentDocs(recent)

      return { id, path, name }
    }

    return { id, path: '', name: '新文档.docx' }
  })

  registerHandle('docs:open-path', async (_event: unknown, filePath: unknown) => {
    if (typeof filePath !== 'string' || !isManagedDocPath(filePath)) {
      throw new Error('document path is outside the web storage area')
    }
    if (!existsSync(filePath as string)) {
      throw new Error(`File not found: ${filePath}`)
    }

    const bytes = readFileSync(filePath as string)
    const name = basename(filePath as string)
    const id = `doc-${Date.now()}`

    const recent = loadRecentDocs()
    recent.unshift({ id, path: filePath as string, name, openedAt: Date.now(), modified: false })
    saveRecentDocs(recent)

    return {
      id,
      path: filePath,
      name,
      bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    }
  })

  registerHandle('docs:read-path', async (_event: unknown, filePath: unknown) => {
    if (typeof filePath !== 'string' || !isManagedDocPath(filePath)) return null
    if (!existsSync(filePath as string)) {
      return null
    }
    const bytes = readFileSync(filePath as string)
    return {
      name: basename(filePath as string),
      bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    }
  })

  registerHandle(
    'docs:save',
    async (_event: unknown, filePath: unknown, data: unknown, _auto?: unknown) => {
      if (typeof filePath !== 'string' || !isManagedDocPath(filePath)) {
        return { ok: false, error: 'save target is outside the web storage area' }
      }
      const bytes = bytesFrom(data)
      if (!bytes || bytes.byteLength === 0) {
        return { ok: false, error: 'save data is empty or invalid' }
      }
      try {
        mkdirSync(resolve(filePath, '..'), { recursive: true })
        writeFileSync(filePath, bytes)
        const recent = loadRecentDocs().filter((doc) => doc.path !== filePath)
        recent.unshift({
          id: basename(filePath, extname(filePath)),
          path: filePath,
          name: basename(filePath),
          openedAt: Date.now(),
          modified: false,
        })
        saveRecentDocs(recent)
        return { ok: true }
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : String(error) }
      }
    },
  )

  registerHandle('docs:create-document', async (_event: unknown, request: unknown) => {
    const value = request as { type?: unknown; title?: unknown; content?: unknown } | null
    const type = value?.type
    const title = String(value?.title ?? 'Untitled')
      .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_')
      .trim()
      .slice(0, 80) || 'Untitled'
    const content = typeof value?.content === 'string' ? value.content : ''
    if (!content.trim()) return { ok: false, error: 'content must not be empty' }
    if (type !== 'md' && type !== 'html') {
      return { ok: false, error: `WEB_UNSUPPORTED: document type '${String(type)}' requires the desktop renderer` }
    }
    const filePath = join(DATA_DIR, `${title}.${type}`)
    try {
      writeFileSync(filePath, content, 'utf8')
      return { ok: true, path: filePath }
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  })

  registerHandle('docs:open-decrypt', () => ({
    ok: false,
    reason: 'error',
    error: 'WEB_UNSUPPORTED: password-protected DOCX requires the desktop renderer',
  }))

  registerHandle('docs:set-password', () => ({
    ok: false,
    error: 'WEB_UNSUPPORTED: DOCX password protection requires the desktop renderer',
  }))

  registerHandle('docs:ai-generate-image', () => ({
    ok: false,
    error: 'WEB_UNSUPPORTED: image generation is not available in the web docs bridge',
  }))

  registerHandle('files:add-pasted-image', async (_event: unknown, data: unknown, ext: unknown) => {
    const cleanExt = typeof ext === 'string' ? ext.toLowerCase().replace(/^\./, '') : ''
    if (!['png', 'jpg', 'jpeg', 'gif', 'webp'].includes(cleanExt)) {
      return { accepted: [], rejected: ['not an image'] }
    }
    const bytes = bytesFrom(data)
    if (!bytes || bytes.byteLength === 0) return { accepted: [], rejected: ['image data is empty'] }
    if (bytes.byteLength > MAX_PASTED_IMAGE_BYTES) {
      return { accepted: [], rejected: ['image is too large'] }
    }
    const name = `pasted-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.${cleanExt}`
    const path = join(FILES_DIR, name)
    writeFileSync(path, bytes)
    const stat = statSync(path)
    return {
      accepted: [{ path, name, ext: cleanExt, sizeBytes: stat.size }],
      rejected: [],
    }
  })

  registerHandle(
    'docs:save-new',
    async (_event: unknown, defaultName?: unknown, data?: unknown, projectId?: unknown) => {
      const name =
        (typeof defaultName === 'string' && defaultName) ||
        `文档-${new Date().toLocaleDateString()}.docx`
      const id = `doc-${Date.now()}`
      const path = join(FILES_DIR, `${id}.docx`)

      if (data) {
        writeFileSync(path, Buffer.from(data as ArrayBuffer))
      }

      if (typeof projectId === 'string' && projectId) {
        const projects = loadProjects()
        const project = projects.find((p) => p.id === projectId)
        if (project && !project.files.includes(`${id}.docx`)) {
          project.files.push(`${id}.docx`)
          project.updatedAt = Date.now()
          saveProjects(projects)
        }
      }

      return { id, path, name }
    },
  )

  registerHandle('docs:print', () => ({
    ok: true,
    message: '请使用浏览器的打印功能 (Ctrl+P 或 Cmd+P)',
  }))

  registerHandle('docs:consume-new-blank', () => false)
  registerHandle('docs:consume-pending-open', () => null)
  registerHandle('docs:consume-ai-doc-content', () => null)
  registerHandle('docs:write-recovery', () => ({ ok: true }))
  registerHandle('docs:password-intent-revision', () => 0)
}
