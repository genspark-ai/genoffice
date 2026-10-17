import { z } from 'zod'

/**
 * Zod schemas for the quarantine channels (H5 owns; compose-schemas
 * precedent: `z.object({...}).strict()` throughout). Main-process handlers
 * validate with these; the quarantine preload bridge keeps a hand-rolled
 * validator and stays zod-free.
 */

export const quarantineScanSchema = z
  .object({
    filename: z.string().min(1).max(256),
    mime: z.string().min(1).max(128),
    data: z.instanceof(Uint8Array),
  })
  .strict()

export const quarantineIngestSchema = z
  .object({
    filename: z.string().min(1).max(256),
    mime: z.string().min(1).max(128),
    data: z.instanceof(Uint8Array),
    accountId: z.string().min(1).max(256).optional(),
    msgUid: z.string().min(1).max(256).optional(),
  })
  .strict()

export const quarantinePreviewSchema = z
  .object({
    path: z.string().min(1).max(4096),
  })
  .strict()

export const quarantineOpenSchema = z
  .object({
    path: z.string().min(1).max(4096),
    confirmed: z.boolean(),
  })
  .strict()

export type QuarantineScanWire = z.infer<typeof quarantineScanSchema>
export type QuarantineIngestWire = z.infer<typeof quarantineIngestSchema>
export type QuarantinePreviewWire = z.infer<typeof quarantinePreviewSchema>
export type QuarantineOpenWire = z.infer<typeof quarantineOpenSchema>
