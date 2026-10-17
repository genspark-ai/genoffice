import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve, sep } from 'node:path'
import { stageAttachment } from '@genoffice/mail-protocols'
import { QUARANTINE_CHANNELS } from '../shared/quarantine-channels'
import type {
  QuarantineIngestWire,
  QuarantinePreviewWire,
  QuarantineScanWire,
} from '../shared/quarantine-channels'
import {
  quarantineIngestSchema,
  quarantineOpenSchema,
  quarantinePreviewSchema,
  quarantineScanSchema,
} from '../shared/quarantine-schemas'
import {
  confinePath,
  extractPreviewText,
  resolveSendCaps,
  runAttachmentPipeline,
  scanAttachmentBytes,
} from './quarantine'
import type { MailDriverKind, PipelineDirs, PipelineProgress, PreviewExtractor } from './quarantine'

// Quarantine IPC handlers (H5 owns this file). Same fake-bus shape as the
// compose handlers so vitest can drive them without Electron. Every handler
// validates with zod end-to-end; verdicts and refs cross IPC, never blobs.

/** Minimal handler table so tests can inject a fake bus (ipcMain-compatible). */
export interface QuarantineIpcBus {
  handle(channel: string, listener: (payload: unknown) => Promise<unknown>): void
}

export interface QuarantineHandlerDeps {
  /** Pipeline disk roots (stage/quarantine/store sibling roots). */
  dirs?: Partial<PipelineDirs>
  /** Driver for per-driver cap resolution (Q-5); defaults to IMAP mirror. */
  driver?: MailDriverKind
  /** OS open for stored files (shell.openPath in prod, which resolves a string). */
  openPath?: (path: string) => unknown
  /** Preview text extractor (file-parse in prod, fakes in tests). */
  extractText?: PreviewExtractor
}

const busesRegistered = new WeakSet<object>()

function badRequest(message: string): Error {
  return new Error(`mail: ${message}`)
}

export function defaultQuarantineDirs(): PipelineDirs {
  const root = join(tmpdir(), 'genoffice-mail-quarantine')
  return {
    stageDir: join(root, 'stage'),
    quarantineDir: join(root, 'quarantine'),
    storeDir: join(root, 'store'),
  }
}

export function registerQuarantineHandlers(
  bus: QuarantineIpcBus,
  deps: QuarantineHandlerDeps = {},
): void {
  if (busesRegistered.has(bus)) return
  busesRegistered.add(bus)

  const base = defaultQuarantineDirs()
  const dirs: PipelineDirs = {
    stageDir: resolve(deps.dirs?.stageDir ?? base.stageDir),
    quarantineDir: resolve(deps.dirs?.quarantineDir ?? base.quarantineDir),
    storeDir: resolve(deps.dirs?.storeDir ?? base.storeDir),
  }
  const caps = resolveSendCaps(deps.driver)

  bus.handle(QUARANTINE_CHANNELS.scan, async (payload) => {
    const parsed = quarantineScanSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad quarantine scan request')
    if (parsed.data.data.byteLength === 0) throw badRequest('bad quarantine scan request')
    const result = scanAttachmentBytes(
      { filename: parsed.data.filename, mime: parsed.data.mime, bytes: parsed.data.data },
      caps,
    )
    const wire: QuarantineScanWire = {
      verdict: result.verdict,
      reason: result.reason,
      filename: result.filename,
      sizeWarn: result.sizeWarn,
      ...(result.detail !== undefined ? { detail: result.detail } : {}),
    }
    return wire
  })

  bus.handle(QUARANTINE_CHANNELS.ingest, async (payload) => {
    const parsed = quarantineIngestSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad quarantine ingest request')
    if (parsed.data.data.byteLength === 0) throw badRequest('bad quarantine ingest request')
    const progress: PipelineProgress[] = []
    const outcome = runAttachmentPipeline(
      { filename: parsed.data.filename, mime: parsed.data.mime, bytes: parsed.data.data },
      dirs,
      {
        stage: stageAttachment,
        onProgress: (state) => {
          progress.push(state)
        },
      },
      caps,
    )
    const wire: QuarantineIngestWire = {
      verdict: outcome.verdict,
      reason: outcome.reason,
      filename: outcome.filename,
      sizeWarn: outcome.sizeWarn,
      ...(outcome.detail !== undefined ? { detail: outcome.detail } : {}),
      ...(outcome.path !== undefined ? { path: outcome.path } : {}),
      progress,
    }
    return wire
  })

  bus.handle(QUARANTINE_CHANNELS.preview, async (payload) => {
    const parsed = quarantinePreviewSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad quarantine preview request')
    const confined = confinePath(parsed.data.path, dirs.storeDir)
    if (!existsSync(confined)) throw badRequest('unknown stored attachment')
    const outcome = await extractPreviewText(
      confined,
      parsed.data.path,
      dirs.storeDir,
      'clean',
      deps.extractText,
    )
    const wire: QuarantinePreviewWire = {
      allowed: outcome.allowed,
      reason: outcome.reason,
      ...(outcome.text !== undefined ? { text: outcome.text } : {}),
      ...(outcome.kind !== undefined ? { kind: outcome.kind } : {}),
      ...(outcome.detail !== undefined ? { detail: outcome.detail } : {}),
    }
    return wire
  })

  bus.handle(QUARANTINE_CHANNELS.open, async (payload) => {
    const parsed = quarantineOpenSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad quarantine open request')
    let confined: string | null = null
    for (const root of [dirs.storeDir, dirs.quarantineDir]) {
      try {
        confined = confinePath(parsed.data.path, root)
        break
      } catch {
        confined = null
      }
    }
    if (confined === null || !existsSync(confined)) throw badRequest('unknown stored attachment')
    const inQuarantine =
      confined === dirs.quarantineDir || confined.startsWith(dirs.quarantineDir + sep)
    if (inQuarantine && parsed.data.confirmed !== true) {
      throw badRequest('quarantined file needs an explicit confirm')
    }
    if (deps.openPath === undefined) throw badRequest('attachment open unavailable')
    await deps.openPath(confined)
    return { opened: true }
  })
}
