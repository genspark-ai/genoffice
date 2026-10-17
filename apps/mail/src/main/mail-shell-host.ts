import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { stageAttachment } from '@genoffice/mail-protocols'
import { FakeComposeProvider, registerComposeHandlers } from './compose-handlers'
import type { ComposeIpcBus } from './compose-handlers'
import { registerMailReadHandlers } from './mail-read-handlers'
// H5 quarantine (additive; merge-owned file, registration call only).
import { registerQuarantineHandlers } from './quarantine-handlers'
// P5 AI retrieval (W owns this import + call only): scoped retrieval over
// the Lane-T store with redaction-backed prompt hygiene.
import { defaultAiRetrieveDeps, registerAiRetrieveHandlers } from './ai'

/**
 * Shell-hosted mail wiring (merge lane owns this file). The shell main
 * process calls `registerMailShellHandlers` once with an ipcMain-backed
 * bus: read handlers serve the fake demo mailbox and compose handlers run
 * behind the fake compose provider (FAKES ONLY — never real accounts).
 * Standalone mail mode (`apps/mail/src/main/index.ts`) is untouched.
 */
export interface MailShellHostDeps {
  /** OS open for validated external URLs (shell.openExternal in prod). */
  openExternal: (url: string) => Promise<void> | void
  /** OS open for stored attachment files (shell.openPath in prod). */
  openPath?: (path: string) => unknown
  /** Staging root override (defaults to the temp compose dir). */
  stageDir?: string
}

export function registerMailShellHandlers(bus: ComposeIpcBus, deps: MailShellHostDeps): void {
  registerMailReadHandlers(bus)
  registerAiRetrieveHandlers(bus, defaultAiRetrieveDeps())
  registerComposeHandlers(bus, {
    provider: new FakeComposeProvider(),
    stage: stageAttachment,
    stageDir: deps.stageDir ?? join(tmpdir(), 'genoffice-mail-compose'),
    openExternal: deps.openExternal,
  })
  registerQuarantineHandlers(bus, {
    openPath: deps.openPath,
  })
}
