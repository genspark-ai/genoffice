import { z } from 'zod'

import { mailMessageRefSchema } from './ipc-schemas'

/**
 * Zod schemas for the read-path channels that have no Lane S schema
 * (merge lane owns this file; Lane S `ipc-schemas.ts` is untouched).
 * Sheets `desktop-api.ts` precedent: `z.object({...}).strict()`
 * throughout. Main-process handlers validate with these; the preload
 * bridge keeps a hand-rolled validator and stays zod-free.
 */

export const mailAddAccountSchema = z
  .object({
    label: z.string().min(1).max(256),
  })
  .strict()

export const mailRemoveAccountSchema = z
  .object({
    accountId: z.string().min(1).max(256),
  })
  .strict()

export const mailListFoldersSchema = z
  .object({
    accountId: z.string().min(1).max(256),
  })
  .strict()

export const mailFetchBodySchema = z
  .object({
    ref: mailMessageRefSchema,
  })
  .strict()

export type MailAddAccountWire = z.infer<typeof mailAddAccountSchema>
export type MailFetchBodyWire = z.infer<typeof mailFetchBodySchema>
