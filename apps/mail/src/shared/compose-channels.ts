import type { StagedAttachment } from '@genoffice/mail-protocols'

/**
 * Compose-session channel names (Lane C owns this file; Lane S/R consume,
 * never edit). The canonical compose rows (`mail:compose-send`,
 * `mail:save-draft`, `mail:outbox-status`) stay in `ipc.ts` (Lane S owns);
 * only compose-session EXTRAS live here: attachment staging, outbox
 * introspection for the intent-recording stub, and renderer-initiated
 * external opens. This file stays zod-free so the preload can import it
 * (sheets precedent); zod schemas live in `./compose-schemas`.
 */
export const COMPOSE_CHANNELS = {
  stageAttachment: 'mail:compose-stage-attachment',
  outboxList: 'mail:compose-outbox-list',
  openExternal: 'mail:compose-open-external',
} as const

export type ComposeChannel = (typeof COMPOSE_CHANNELS)[keyof typeof COMPOSE_CHANNELS]

/**
 * Outbox intent-recording stub (Lane C shape; real replay is the sync
 * phase, recorded as a limitation — do not build replay here). Entries are
 * send intents the main process accepted; `failed` entries carry the
 * provider error for surfacing, never secrets.
 */
export interface OutboxEntry {
  id: string
  to: string[]
  subject: string
  state: 'queued' | 'failed'
  /** ISO timestamp of acceptance. */
  queuedAt: string
  /** Provider error text (send-side only, no secrets). */
  error?: string
}

/** API exposed by the preload to the renderer (window.mailCompose). */
export interface ComposeApi {
  /**
   * Stage renderer-held bytes to a path-canonical StagedAttachment on disk
   * (quarantine scan happens downstream; staging never executes content).
   */
  stageAttachment(filename: string, mime: string, data: Uint8Array): Promise<StagedAttachment>
  /** Intent-recording stub listing (no replay; sync phase owns replay). */
  outboxList(accountId: string): Promise<OutboxEntry[]>
  /**
   * Open an external/compose URL from the renderer. Main validates with zod
   * then `safeExternalUrl` (per-call mailto allowlist at the call-site);
   * anything else is rejected, never opened.
   */
  openExternal(url: string): Promise<{ opened: boolean }>
}
