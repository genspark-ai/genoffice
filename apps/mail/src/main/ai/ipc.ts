import { MAIL_CHANNELS } from '../../shared/ipc'
import { mailAiRetrieveSchema } from '../../shared/ipc-schemas'
import { aiRetrieve } from './retrieve'
import type { AiRetrieveDeps } from './retrieve'

/**
 * mail:ai-retrieve IPC registration (P5 spec section 6). The payload is
 * zod-validated end-to-end (strict) before any store opens; the preload
 * stays zod-free with hand-rolled guards (sheets precedent). Errors carry
 * mail-prefixed summaries only — bodies, snippets, and secrets never
 * cross this boundary. Main-process only.
 */

/** Minimal handler table so tests can inject a fake bus (ipcMain-compatible). */
export interface AiIpcBus {
  handle(channel: string, listener: (payload: unknown) => Promise<unknown>): void
}

function badRequest(message: string): Error {
  return new Error(`mail: ${message}`)
}

let handlersRegistered = false

/** Test-only reset (fresh registration on the next register call). */
export function resetAiRetrieveHandlers(): void {
  handlersRegistered = false
}

/** Register the single AI invoke channel over the injected store deps. */
export function registerAiRetrieveHandlers(bus: AiIpcBus, deps: AiRetrieveDeps): void {
  if (handlersRegistered) return
  handlersRegistered = true
  bus.handle(MAIL_CHANNELS.aiRetrieve, async (payload) => {
    const parsed = mailAiRetrieveSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad ai-retrieve request')
    return aiRetrieve(
      {
        scope: parsed.data.scope,
        op: parsed.data.op,
        query: parsed.data.query,
        folderId: parsed.data.folderId,
        limit: parsed.data.limit,
        refs: parsed.data.refs,
      },
      deps,
    )
  })
}
