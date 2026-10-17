/**
 * Attachment staging: temp file landing plus content-addressed store.
 * Files land in staging first; the quarantine scan (P5) promotes them to
 * the store. Hostile-file parsing never runs in the renderer. Size policy
 * mirrors the spec: outgoing warning above 25 MB, inline preview capped
 * at the 50 MB markdown precedent.
 */

import { createHash } from 'node:crypto'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import type { StagedAttachment } from './types'

/** Outgoing attachments above this size warn in compose (per-driver caps apply). */
export const OUTGOING_WARN_BYTES = 25 * 1024 * 1024
/** Inline preview cap; staging rejects anything larger. */
export const INLINE_PREVIEW_CAP_BYTES = 50 * 1024 * 1024

/** Control characters rejected from stored filenames (stripping them is the point). */
// eslint-disable-next-line no-control-regex -- control characters are rejected on purpose
const CONTROL_CHARS_RE = /[\x00-\x1f\x7f]/g

/** Keep store filenames portable: strip separators and control chars. */
export function safeFilename(filename: string): string {
  const base = filename.split(/[\\/]/).pop() ?? 'unnamed'
  const clean = base.replace(CONTROL_CHARS_RE, '').trim().replace(/^\.+/, '')
  return clean === '' ? 'unnamed' : clean
}

/** Content address for store placement: sha256 hex of the bytes. */
export function contentAddress(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

/** Size verdict for compose: ok, warn (large), or reject (over preview cap). */
export function checkAttachmentSize(size: number): 'ok' | 'warn' | 'reject' {
  if (size > INLINE_PREVIEW_CAP_BYTES) return 'reject'
  if (size > OUTGOING_WARN_BYTES) return 'warn'
  return 'ok'
}

export interface StageInput {
  filename: string
  mime: string
  bytes: Uint8Array
  /** Staging root (temp). Never the final store: quarantine scan first. */
  stageDir: string
}

/**
 * Write bytes to a content-addressed file under stageDir and return the
 * staged handle for SendRequest. Throws on oversize input.
 */
export function stageAttachment(input: StageInput): StagedAttachment {
  const verdict = checkAttachmentSize(input.bytes.byteLength)
  if (verdict === 'reject') {
    throw new Error(`attachment exceeds ${INLINE_PREVIEW_CAP_BYTES} byte preview cap`)
  }
  const digest = contentAddress(input.bytes)
  const filename = safeFilename(input.filename)
  mkdirSync(input.stageDir, { recursive: true })
  const path = join(input.stageDir, `${digest}-${filename}`)
  writeFileSync(path, input.bytes)
  return { filename, mime: input.mime, size: input.bytes.byteLength, path }
}
