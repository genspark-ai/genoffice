/**
 * Docs IPC channels — recent, font-metrics, pick-image, settings,
 * open/save/read/print. Persistence is in-memory + `docs-recent.json`; the
 * `DOCS_RECENT` map referenced by `home:recents` is the same singleton
 * declared in `common/state.ts`.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, join } from 'node:path'
import {
  FILES_DIR,
  loadProjects,
  loadRecentDocs,
  registerHandle,
  saveProjects,
  saveRecentDocs,
} from '../common/index.js'

export function registerDocsHandlers(): void {
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
    if (!existsSync(filePath as string)) {
      return null
    }
    const bytes = readFileSync(filePath as string)
    return {
      name: basename(filePath as string),
      bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    }
  })

  registerHandle('docs:save-new', async (_event: unknown, args: unknown) => {
    const { defaultName, data, projectId } = args as { defaultName?: string; data?: ArrayBuffer; projectId?: string }
    const name = defaultName || `文档-${new Date().toLocaleDateString()}.docx`
    const id = `doc-${Date.now()}`
    const path = join(FILES_DIR, `${id}.docx`)

    if (data) {
      writeFileSync(path, Buffer.from(data))
    }

    if (projectId) {
      const projects = loadProjects()
      const project = projects.find(p => p.id === projectId)
      if (project && !project.files.includes(`${id}.docx`)) {
        project.files.push(`${id}.docx`)
        project.updatedAt = Date.now()
        saveProjects(projects)
      }
    }

    return { id, path, name }
  })

  registerHandle('docs:print', () => ({
    ok: true,
    message: '请使用浏览器的打印功能 (Ctrl+P 或 Cmd+P)',
  }))

  registerHandle('docs:consume-new-blank', () => ({ ok: true }))
  registerHandle('docs:consume-pending-open', () => null)
  registerHandle('docs:consume-ai-doc-content', () => ({ ok: true }))
  registerHandle('docs:write-recovery', () => ({ ok: true }))
  registerHandle('docs:password-intent-revision', () => 0)
}
