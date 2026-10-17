import type { PipelineProgress, QuarantineReason, QuarantineVerdict } from '../main/quarantine'

/** Re-exported for renderer use (type-only; erased at build, never a runtime dep). */
export type { PipelineProgress, QuarantineReason, QuarantineVerdict } from '../main/quarantine'

/**
 * Quarantine IPC channel names (H5 owns this file). New-file discipline:
 * the H4-hotspot `ipc.ts` / `ipc-schemas.ts` stay untouched, so quarantine
 * traffic gets its own channel table. This file stays zod-free so the
 * preload can import it (compose-channels precedent); zod schemas live in
 * `./quarantine-schemas`.
 */
export const QUARANTINE_CHANNELS = {
  /** Pre-send/pre-store scan over renderer-held bytes (no disk side effects). */
  scan: 'mail:quarantine-scan',
  /** Receive/compose ingest: staged-temp --> scan --> store pipeline. */
  ingest: 'mail:quarantine-ingest',
  /** Text preview for clean-verdict stored files (execs never inline). */
  preview: 'mail:quarantine-preview',
  /** OS-open a stored file (explicit confirm required for quarantined). */
  open: 'mail:quarantine-open',
} as const

export type QuarantineChannel = (typeof QUARANTINE_CHANNELS)[keyof typeof QUARANTINE_CHANNELS]

/** Scan verdict wire shape (verdict + reason code, never file bytes). */
export interface QuarantineScanWire {
  verdict: QuarantineVerdict
  reason: QuarantineReason
  filename: string
  sizeWarn: boolean
  detail?: string
}

/** Pipeline outcome wire shape (refs only; full blobs never cross IPC). */
export interface QuarantineIngestWire extends QuarantineScanWire {
  /** Disk home for clean/quarantined; absent when blocked. */
  path?: string
  progress: PipelineProgress[]
}

/** Text preview wire shape (clean verdicts only; execs are refused). */
export interface QuarantinePreviewWire {
  allowed: boolean
  reason: QuarantineReason
  text?: string
  kind?: 'text' | 'image' | 'unsupported'
  detail?: string
}

/** API exposed by the preload to the renderer (window.mailQuarantine). */
export interface QuarantineApi {
  /**
   * Scan renderer-held bytes before staging (quarantine scan verdict, no
   * disk writes). Hostile-file parsing still runs in main only.
   */
  scan(filename: string, mime: string, data: Uint8Array): Promise<QuarantineScanWire>
  /** Run the full staged-temp --> scan --> store pipeline on bytes. */
  ingest(
    filename: string,
    mime: string,
    data: Uint8Array,
    opts?: { accountId?: string; msgUid?: string },
  ): Promise<QuarantineIngestWire>
  /** Text preview of a stored clean file (blocked/execs are refused). */
  preview(path: string): Promise<QuarantinePreviewWire>
  /**
   * OS-open a stored file. Quarantined files need an explicit renderer-side
   * confirm first (confirmed=true); blocked files have no path to open.
   */
  openAttachment(path: string, confirmed: boolean): Promise<{ opened: boolean }>
}
