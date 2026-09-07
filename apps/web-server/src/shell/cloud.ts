/**
 * Cloud channels — file upload/list/download/delete/sign-url. Backed by
 * the local `CLOUD_FILES` map; bytes live in `FILES_DIR`. Phase 3 (LUM-551)
 * will replace this with real S3/MinIO/OSS.
 */
import { existsSync, readFileSync, unlinkSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { CLOUD_FILES, FILES_DIR, registerHandle } from '../common/index.js'

export function registerCloudHandlers(): void {
  registerHandle('cloud:upload', async (_event: unknown, args: unknown) => {
    const { name, bytes, mimeType, public: isPublic } = args as {
      name: string
      bytes: ArrayBuffer
      mimeType: string
      public?: boolean
    }

    const id = `cloud-${Date.now()}-${name}`
    const url = `/cloud/files/${id}`

    const filePath = join(FILES_DIR, id)
    writeFileSync(filePath, Buffer.from(bytes))

    CLOUD_FILES.set(id, {
      id,
      name,
      size: bytes.byteLength,
      type: mimeType,
      url,
      createdAt: Date.now(),
      updatedAt: Date.now(),
      public: isPublic || false,
    })

    return { ok: true, id, url, name, size: bytes.byteLength }
  })

  registerHandle('cloud:list', (_event: unknown, args: unknown) => {
    const { type, search } = args as { type?: string; search?: string }

    let files = [...CLOUD_FILES.values()]

    if (type) {
      files = files.filter(f => f.type.startsWith(type))
    }

    if (search) {
      const searchLower = search.toLowerCase()
      files = files.filter(f => f.name.toLowerCase().includes(searchLower))
    }

    return files.map(f => ({
      id: f.id,
      name: f.name,
      size: f.size,
      type: f.type,
      url: f.url,
      createdAt: f.createdAt,
      updatedAt: f.updatedAt,
    }))
  })

  registerHandle('cloud:download', async (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    const file = CLOUD_FILES.get(id)
    if (!file) return null

    const filePath = join(FILES_DIR, id)
    if (!existsSync(filePath)) return null

    const bytes = readFileSync(filePath)
    return {
      name: file.name,
      type: file.type,
      bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    }
  })

  registerHandle('cloud:delete', (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    if (!CLOUD_FILES.has(id)) return { ok: false, error: 'File not found' }

    const filePath = join(FILES_DIR, id)
    if (existsSync(filePath)) unlinkSync(filePath)

    CLOUD_FILES.delete(id)
    return { ok: true }
  })

  registerHandle('cloud:get-url', (_event: unknown, args: unknown) => {
    const { id, expires } = args as { id: string; expires?: number }
    const file = CLOUD_FILES.get(id)
    if (!file) return null

    const expiry = expires || 3600
    const signedUrl = `${file.url}?token=${Date.now()}&expires=${Date.now() + expiry * 1000}`

    return { url: signedUrl, expires: expiry }
  })
}
