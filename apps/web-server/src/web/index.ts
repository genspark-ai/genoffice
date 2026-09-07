/**
 * Web-platform helper channels — temp file read/write, temp dir creation,
 * and the `web:save-file` helper used by browser pickers to land a chosen
 * file into the project's `FILES_DIR` and (optionally) attach it to a
 * project.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, join } from 'node:path'
import { tmpdir } from 'node:os'
import {
  FILES_DIR,
  loadProjects,
  registerHandle,
  saveProjects,
  WEB_TEMP_ROOT,
} from '../common/index.js'

export function registerWebHandlers(): void {
  registerHandle('web:write-temp-file', async (_event: unknown, request: unknown) => {
    const record = request as { name?: unknown; bytes?: unknown } | null
    if (!record || typeof record.name !== 'string' || !(record.bytes instanceof ArrayBuffer)) {
      throw new Error('web:write-temp-file expects { name: string, bytes: ArrayBuffer }')
    }
    const safeName = basename(record.name).replace(/[^\w.\- ]+/g, '_') || 'file'
    const dir = mkdirSync(join(WEB_TEMP_ROOT, `upload-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`), { recursive: true })
    const filePath = join(dir, safeName)
    writeFileSync(filePath, Buffer.from(record.bytes))
    return filePath
  })

  registerHandle('web:read-file-bytes', async (_event: unknown, path: unknown) => {
    if (!existsSync(path as string)) {
      throw new Error(`File not found: ${path}`)
    }
    const bytes = readFileSync(path as string)
    return {
      name: basename(path as string),
      bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    }
  })

  registerHandle('web:make-temp-dir', async () =>
    mkdirSync(join(tmpdir(), `genoffice-dir-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`), { recursive: true }),
  )

  registerHandle('web:save-file', async (_event: unknown, request: unknown) => {
    const { name, bytes, projectId } = request as {
      name?: string
      bytes?: ArrayBuffer
      projectId?: string
    }
    if (!name || !bytes) throw new Error('web:save-file expects { name, bytes, projectId? }')

    const fileId = `${Date.now()}-${name}`
    const filePath = join(FILES_DIR, fileId)
    writeFileSync(filePath, Buffer.from(bytes))

    if (projectId) {
      const projects = loadProjects()
      const project = projects.find(p => p.id === projectId)
      if (project && !project.files.includes(fileId)) {
        project.files.push(fileId)
        project.updatedAt = Date.now()
        saveProjects(projects)
      }
    }

    return { id: fileId, path: filePath, name }
  })
}
