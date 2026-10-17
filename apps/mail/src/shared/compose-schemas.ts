import { z } from 'zod'

/**
 * Zod schemas for the compose-session channels (Lane C owns; sheets
 * `desktop-api.ts` precedent: `z.object({...}).strict()` throughout).
 * Main-process handlers validate with these; the compose preload bridge
 * keeps a hand-rolled validator and stays zod-free. Canonical send/draft
 * schemas live in `./ipc-schemas` (Lane S owns) and are reused, never
 * redefined here.
 */

export const composeStageAttachmentSchema = z
  .object({
    filename: z.string().min(1).max(256),
    mime: z.string().min(1).max(128),
    data: z.instanceof(Uint8Array),
  })
  .strict()

export const composeOutboxListSchema = z
  .object({
    accountId: z.string().min(1).max(256),
  })
  .strict()

export const composeOutboxStatusSchema = z
  .object({
    accountId: z.string().min(1).max(256),
  })
  .strict()

/**
 * Renderer-initiated external open (Alya XR-1 record): the URL is
 * zod-validated here first, then gated by `safeExternalUrl` with a
 * per-call mailto allowlist at the handler call-site. The main-process
 * default stays http/https-only; this schema never widens it.
 */
export const composeOpenExternalSchema = z
  .object({
    url: z.string().min(1).max(2048),
  })
  .strict()

export type ComposeStageAttachmentWire = z.infer<typeof composeStageAttachmentSchema>
export type ComposeOpenExternalWire = z.infer<typeof composeOpenExternalSchema>
