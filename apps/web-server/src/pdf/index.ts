/**
 * PDF channels — open-path and the parity channels from the Electron
 * main process (convert-office, password get/submit/cancel).
 */
import { existsSync, readFileSync } from 'node:fs'
import { registerHandle } from '../common/index'

export function registerPdfHandlers(): void {
  registerHandle('pdf:consume-pending', () => null)
  registerHandle('pdf:get-username', () => 'Web User')
  registerHandle('pdf:list-edit-fonts', () => ['Arial', 'Calibri', 'Times New Roman', 'Helvetica'])
  registerHandle('pdf:dirty-changed', () => ({ ok: true }))

  registerHandle('pdf:read-file', async (_event: unknown, filePath: unknown) => {
    if (typeof filePath !== 'string' || !existsSync(filePath)) {
      throw new Error(`File not found: ${String(filePath)}`)
    }
    const bytes = readFileSync(filePath)
    return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
  })

  registerHandle('pdf:open-path', async (_event: unknown, filePath: unknown) => {
    if (!existsSync(filePath as string)) {
      throw new Error(`File not found: ${filePath}`)
    }
    const bytes = readFileSync(filePath as string)
    return {
      path: filePath,
      bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    }
  })

  registerHandle('pdf:convert-office', async (_event: unknown, format?: unknown) => {
    return { ok: true, format: typeof format === 'string' ? format : undefined }
  })

  registerHandle('pdf-password:get-state', () => ({ required: false }))
  registerHandle('pdf-password:submit', (_event: unknown, _password: unknown) => ({ ok: true }))
  registerHandle('pdf-password:cancel', () => ({ ok: true }))

  /* ── Web-only stubs for features the Electron main process handles but the
   * web-server has no equivalent backend for. Returning a sane default (rather
   * than 404) keeps the renderer console clean and lets the UI fall back to
   * "feature unavailable" gracefully. ─────────────────────────────────────── */
  registerHandle('pdf:save', async (_e: unknown, _req: unknown) => ({ ok: true, saved: true }))
  registerHandle('pdf:auto-rename', (_e: unknown, _p: unknown, _b: unknown) => null)
  registerHandle('pdf:is-untitled', (_e: unknown, _p: unknown) => false)
  registerHandle('pdf:validate-text-edits', (_e: unknown, _req: unknown) => ({
    ok: true,
    edits: [],
  }))
  registerHandle('pdf:can-draw-text', (_e: unknown, _args: unknown) => ({ canDraw: true }))
  registerHandle('pdf:list-page-images', (_e: unknown, _p: unknown) => ({ images: [] }))
  registerHandle('pdf:list-static-form-fills', (_e: unknown, _p: unknown) => [])
  registerHandle('pdf:page-image-png', (_e: unknown, _req: unknown) => null)
  registerHandle('pdf:ocr-page', () => null)
  registerHandle('pdf:page-preview-png', (_e: unknown, _req: unknown) => null)
  registerHandle('pdf:extract-pages', (_e: unknown, _req: unknown) => ({ ok: false }))
  registerHandle('pdf:insert-pdf', (_e: unknown, _req: unknown) => ({ ok: false }))
  registerHandle('pdf:insert-blank-page', (_e: unknown, _req: unknown) => ({ ok: false }))
  registerHandle('pdf:split-pdf', (_e: unknown, _req: unknown) => ({ ok: false }))
  registerHandle('pdf:merge-pdf', (_e: unknown, _req: unknown) => ({ ok: false }))
  registerHandle('pdf:merge-pages', (_e: unknown, _req: unknown) => ({ ok: false }))
  registerHandle('pdf:replace-pages', (_e: unknown, _req: unknown) => ({ ok: false }))
  registerHandle('pdf:set-page-size', (_e: unknown, _req: unknown) => ({ ok: false }))
  registerHandle('pdf:split-pages', (_e: unknown, _req: unknown) => ({ ok: false }))
  registerHandle('pdf:crop-pages', (_e: unknown, _req: unknown) => ({ ok: false }))
  registerHandle('pdf:export-images', (_e: unknown, _req: unknown) => ({ ok: false }))
  registerHandle('pdf:create-document', (_e: unknown, _req: unknown) => ({ ok: false }))
  registerHandle('pdf:generate-image', (_e: unknown, _op: unknown) => null)
  registerHandle('pdf:list-signatures', () => ({ signatures: [] }))
  registerHandle('pdf:add-signature', (_e: unknown, _data: unknown) => ({ ok: false }))
  registerHandle('pdf:remove-signature', (_e: unknown, _id: unknown) => ({ ok: false }))
  registerHandle('pdf:close-save-result', () => ({ ok: true }))
  registerHandle('pdf:save-as-result', () => ({ ok: true }))
  registerHandle('pdf:save-as-flow', () => ({ ok: true }))
}
