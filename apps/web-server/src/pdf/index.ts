/**
 * PDF channels — open-path and the parity channels from the Electron
 * main process (convert-office, password get/submit/cancel).
 */
import { existsSync, readFileSync } from 'node:fs'
import { registerHandle } from '../common/index.js'

export function registerPdfHandlers(): void {
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

  registerHandle('pdf:convert-office', async (_event: unknown, args: unknown) => {
    const { format } = (args || {}) as { format: string }
    return { ok: true, format }
  })

  registerHandle('pdf-password:get-state', () => ({ required: false }))
  registerHandle('pdf-password:submit', (_event: unknown, _password: unknown) => ({ ok: true }))
  registerHandle('pdf-password:cancel', () => ({ ok: true }))
}
