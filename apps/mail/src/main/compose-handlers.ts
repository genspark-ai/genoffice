import { existsSync, statSync } from 'node:fs'
import { resolve, sep } from 'node:path'
import { safeExternalUrl } from '@genoffice/electron-utils'
import type {
  MailBody,
  MailEnvelope,
  MailEvent,
  MailFlags,
  MailFolder,
  MailMessageRef,
  MailProvider,
  MailQuery,
  SendRequest,
  StagedAttachment,
  StageInput,
} from '@genoffice/mail-protocols'
import { MAIL_CHANNELS } from '../shared/ipc'
import { mailSaveDraftSchema, sendRequestSchema } from '../shared/ipc-schemas'
import { COMPOSE_CHANNELS } from '../shared/compose-channels'
import type { OutboxEntry } from '../shared/compose-channels'
import {
  composeOpenExternalSchema,
  composeOutboxListSchema,
  composeOutboxStatusSchema,
  composeStageAttachmentSchema,
} from '../shared/compose-schemas'
import { getMailStore } from './mail-store'
import { resolveSendCaps, scanAttachmentPath, scanFileName } from './quarantine'
import type { MailDriverKind } from './quarantine'

// Compose + drafts main-process handlers (Lane C owns this file — never
// edit `mail-main.ts`). Every handler validates with zod end-to-end before
// touching the provider or the store; drivers run main-only and secrets
// never touch the renderer. UI dev calls MailProvider FAKES ONLY (no real
// accounts); the real drivers and replay engine arrive with the sync phase.

/** Minimal handler table so tests can inject a fake bus (ipcMain-compatible). */
export interface ComposeIpcBus {
  handle(channel: string, listener: (payload: unknown) => Promise<unknown>): void
}

export interface ComposeSendCaps {
  warnTotalBytes: number
  maxTotalBytes: number
  maxAttachmentBytes: number
  maxRecipients: number
}

/**
 * Default send caps mirror the IMAP/SMTP driver
 * (`packages/mail-protocols/src/imap-smtp-driver.ts`: MAX_TOTAL_BYTES,
 * MAX_ATTACHMENT_BYTES, MAX_RECIPIENTS; warn cap is OUTGOING_WARN_BYTES
 * from `attachments.ts`). Per-account driver caps resolve in the sync
 * phase; until then the main process enforces these and the renderer
 * pre-warns at the same thresholds.
 */
export const DEFAULT_COMPOSE_SEND_CAPS: ComposeSendCaps = {
  warnTotalBytes: 25 * 1024 * 1024,
  maxTotalBytes: 50 * 1024 * 1024,
  maxAttachmentBytes: 25 * 1024 * 1024,
  maxRecipients: 100,
}

export interface ComposeHandlerDeps {
  /** MailProvider behind compose-send (FAKES ONLY in P3 UI dev). */
  provider: MailProvider
  /** Canonical staging (`stageAttachment` from mail-protocols, injected). */
  stage: (input: StageInput) => StagedAttachment
  /** Staging root (temp; quarantine scan promotes from here downstream). */
  stageDir: string
  /** OS open for validated external URLs (shell.openExternal in prod). */
  openExternal: (url: string) => Promise<void> | void
  sendCaps?: ComposeSendCaps
  /**
   * Driver for per-driver cap resolution (H5 quarantine): when sendCaps is
   * unset, the scan + send caps resolve for this driver (Q-5). Defaults to
   * the IMAP/SMTP mirror, so unset behavior is unchanged.
   */
  driver?: MailDriverKind
  /** Draft persistence (defaults to the shared in-memory stub store). */
  saveDraft?: (accountId: string, threadKey: string, text: string) => void
  /** ISO clock for outbox entries (injectable for tests). */
  clock?: () => string
}

const busesRegistered = new WeakSet<object>()

function badRequest(message: string): Error {
  return new Error(`mail: ${message}`)
}

function countRecipients(req: SendRequest): number {
  return req.to.length + (req.cc?.length ?? 0) + (req.bcc?.length ?? 0)
}

/**
 * Normalize attachments to path-canonical form before the provider call.
 * A staged file already on disk streams from disk (memory payloads are
 * stripped so sends never carry full blobs); a missing path with a memory
 * payload is staged via the canonical `stageAttachment` (data-fallback
 * contract); a missing path without bytes is rejected.
 *
 * Merge-gate M1 (t_7d6ffa29): a renderer-supplied disk path is honored
 * only when it resolves inside stageDir — otherwise a compromised
 * renderer could exfiltrate arbitrary files through compose-send.
 */
function normalizeAttachments(
  req: SendRequest,
  stage: (input: StageInput) => StagedAttachment,
  stageDir: string,
): StagedAttachment[] {
  const root = resolve(stageDir)
  const out: StagedAttachment[] = []
  for (const att of req.attachments ?? []) {
    if (typeof att.path === 'string' && att.path.length > 0 && existsSync(att.path)) {
      const confined = resolve(att.path)
      if (confined !== root && !confined.startsWith(root + sep)) {
        throw badRequest('staged attachment outside stage dir')
      }
      const onDisk = statSync(att.path).size
      if (onDisk !== att.size) throw badRequest('staged attachment changed on disk')
      const { filename, mime, size, path } = att
      out.push({ filename, mime, size, path })
      continue
    }
    if (att.data instanceof Uint8Array) {
      // H5 quarantine: the original renderer-supplied name is scanned
      // BEFORE staging sanitizes it, so traversal/ADS tricks cannot wash
      // through safeFilename into the send path.
      const nameCheck = scanFileName(att.filename)
      if (!nameCheck.ok) {
        throw badRequest(
          `attachment blocked (${nameCheck.reason ?? 'hostile-name'}): ${att.filename}`,
        )
      }
      const staged = stage({ filename: att.filename, mime: att.mime, bytes: att.data, stageDir })
      out.push({
        filename: staged.filename,
        mime: staged.mime,
        size: staged.size,
        path: staged.path,
      })
      continue
    }
    throw badRequest(`attachment missing: ${att.filename}`)
  }
  return out
}

/**
 * UI-dev MailProvider fake behind compose-send (FAKES ONLY — never real
 * accounts or creds). Records sends and mints outbox ids; failure
 * injection is test-only. Real drivers wire in through `deps.provider`
 * in the sync phase.
 */
export class FakeComposeProvider implements MailProvider {
  readonly sent: SendRequest[] = []
  failSendsWith?: string

  async connect(): Promise<void> {}
  async disconnect(): Promise<void> {}
  async listFolders(): Promise<MailFolder[]> {
    return []
  }
  async list(): Promise<MailEnvelope[]> {
    return []
  }
  async search(_folderId: string, _query: MailQuery): Promise<MailEnvelope[]> {
    return []
  }
  async fetch(ref: MailMessageRef): Promise<MailBody> {
    void ref
    return { text: '' }
  }
  async send(req: SendRequest): Promise<{ outboxId: string }> {
    if (this.failSendsWith !== undefined) throw new Error(this.failSendsWith)
    this.sent.push(req)
    return { outboxId: `compose-outbox-${this.sent.length}` }
  }
  async move(): Promise<void> {}
  async setFlags(): Promise<void> {}
  async delete(): Promise<void> {}
  watch(_handler: (ev: MailEvent) => void): { stop(): void } {
    return { stop() {} }
  }
  flagsOf(_ref: MailMessageRef): MailFlags {
    return { seen: false, flagged: false, answered: false, draft: false }
  }
}

export function registerComposeHandlers(bus: ComposeIpcBus, deps: ComposeHandlerDeps): void {
  if (busesRegistered.has(bus)) return
  busesRegistered.add(bus)

  const scanCaps = deps.sendCaps ?? resolveSendCaps(deps.driver)
  // Q-5: the send-path caps ARE the driver-resolved caps (explicit sendCaps
  // still win). Unset driver resolves to the IMAP/SMTP mirror, so default
  // behavior is unchanged.
  const caps = scanCaps
  const clock = deps.clock ?? (() => new Date().toISOString())
  // Intent-recording stub ledger (Lane C shape; replay is the sync phase).
  const ledger: OutboxEntry[] = []
  const saveDraft =
    deps.saveDraft ??
    ((accountId, threadKey, text) => getMailStore().saveDraft(accountId, threadKey, text))

  bus.handle(MAIL_CHANNELS.composeSend, async (payload) => {
    const parsed = sendRequestSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad send request')
    const req = parsed.data
    if (countRecipients(req) > caps.maxRecipients) throw badRequest('too many recipients')
    const attachments = normalizeAttachments(req, deps.stage, deps.stageDir)
    // H5 quarantine (Q-2/Q-4/Q-5): every staged attachment is scanned from
    // disk before the provider call. Blocked hostile/suspicious verdicts
    // (traversal, hostile names, EICAR-style vectors, executable magic,
    // zip-bomb heuristics) reject the send; quarantined verdicts
    // (executables, scripts, macros, MIME mismatch) still send but are
    // never inline-previewed downstream. Over-cap scan verdicts fall
    // through to the established caps checks below so the send-path error
    // contract (per-account cap messages) stays byte-identical.
    for (const att of attachments) {
      const scanned = scanAttachmentPath(att.path, att.filename, att.mime, att.size, scanCaps)
      if (scanned.verdict === 'blocked' && scanned.reason !== 'over-cap') {
        throw badRequest(`attachment blocked (${scanned.reason}): ${scanned.filename}`)
      }
    }
    const total = req.text.length + attachments.reduce((sum, att) => sum + att.size, 0)
    if (total > caps.maxTotalBytes) throw badRequest('attachments exceed the size cap')
    for (const att of attachments) {
      if (att.size > caps.maxAttachmentBytes) {
        throw badRequest(`attachment exceeds the per-account cap: ${att.filename}`)
      }
    }
    const wire: SendRequest = { ...req, attachments }
    try {
      const { outboxId } = await deps.provider.send(wire)
      ledger.push({
        id: outboxId,
        to: req.to,
        subject: req.subject,
        state: 'queued',
        queuedAt: clock(),
      })
      return { outboxId }
    } catch (error) {
      const message = error instanceof Error ? error.message : 'send failed'
      ledger.push({
        id: `compose-failed-${ledger.length + 1}`,
        to: req.to,
        subject: req.subject,
        state: 'failed',
        queuedAt: clock(),
        error: message,
      })
      throw badRequest('queued send failed')
    }
  })

  bus.handle(MAIL_CHANNELS.saveDraft, async (payload) => {
    const parsed = mailSaveDraftSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad draft')
    saveDraft(parsed.data.accountId, parsed.data.threadKey, parsed.data.text)
    return undefined
  })

  const statusOf = (): { queued: number; failed: number } => ({
    queued: ledger.filter((entry) => entry.state === 'queued').length,
    failed: ledger.filter((entry) => entry.state === 'failed').length,
  })

  bus.handle(MAIL_CHANNELS.outboxStatus, async (payload) => {
    const parsed = composeOutboxStatusSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad outbox status request')
    return statusOf()
  })

  bus.handle(COMPOSE_CHANNELS.outboxList, async (payload) => {
    const parsed = composeOutboxListSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad outbox list request')
    return ledger.map((entry) => ({ ...entry }))
  })

  bus.handle(COMPOSE_CHANNELS.stageAttachment, async (payload) => {
    const parsed = composeStageAttachmentSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad stage request')
    try {
      const staged = deps.stage({
        filename: parsed.data.filename,
        mime: parsed.data.mime,
        bytes: parsed.data.data,
        stageDir: deps.stageDir,
      })
      const { filename, mime, size, path } = staged
      return { filename, mime, size, path }
    } catch {
      throw badRequest('attachments exceed the size cap')
    }
  })

  // Alya XR-1 record: the URL is zod-validated above, then gated by
  // safeExternalUrl with a per-call mailto allowlist at THIS call-site.
  // The safeExternalUrl default (http/https-only) is never changed.
  bus.handle(COMPOSE_CHANNELS.openExternal, async (payload) => {
    const parsed = composeOpenExternalSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad external url')
    const target = safeExternalUrl(parsed.data.url, {
      allowedProtocols: ['http:', 'https:', 'mailto:'],
    })
    if (target === null) throw badRequest('blocked external url')
    await deps.openExternal(target)
    return { opened: true }
  })
}
