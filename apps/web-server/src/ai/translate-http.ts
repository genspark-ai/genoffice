/**
 * HTTP surface for the AI translation pipeline.
 *
 * Adds three endpoints to the standalone web-server so the Dataflare parent's
 * bridge forwarder (and any other HTTP caller) can drive `translateBatch` /
 * `translateBatchStream` without going through the IPC layer:
 *
 *   POST /api/ai/translate              — non-streaming batch (sync result)
 *   POST /api/ai/translate/stream       — SSE: start / unit / quality / complete / error
 *   POST /api/ai/translate/stream/cancel — abort an in-flight stream by requestId
 *
 * Wire shape intentionally matches the legacy Dataflare `StructuredTranslationStreamEvent`
 * so the existing `apps/docs/src/renderer/web-bridge.ts#aiTranslateBatchStream`
 * consumer can switch the target URL with no field changes:
 *
 *   { type, requestId, status, sourceLanguage, targetLanguage,
 *     totalUnits, completedUnits, progress, unit, quality, warnings, message }
 *
 * The unit payload mirrors the same field names the bridge already parses
 * (`payload.unit.unitId`, `payload.unit.sourceText`, `payload.unit.translatedText`,
 * `payload.unit.status`, `payload.unit.matchedTerms`, `payload.unit.warnings`,
 * `payload.unit.errorMessage`).
 *
 * Settings: the caller may include an `AiSettings` blob to override the
 * server's persisted settings (mirrors `ai:translate-batch`). Falls back to
 * `aiSettings` from `./chat.js`.
 */
import type { IncomingMessage, ServerResponse } from 'node:http'

import {
  AiSettings,
  type AiProviderConfig,
  type AiProviderId,
} from '@genoffice/ai-provider'
import {
  type TranslateBatchUnitResult,
  translateBatch,
  translateBatchStream,
} from '@genoffice/translation-core'

import { aiSettings as defaultSettings } from './chat.js'

interface TranslateUnitRequest {
  unitId?: string
  kind?: string
  sourceText?: string
  order?: number
  path?: string
  metadata?: Record<string, unknown>
  range?: { from?: number; to?: number; scope?: string } | null
}

interface TranslateBatchHttpRequest {
  requestId?: string
  idempotencyKey?: string
  documentId?: string
  documentType?: string
  scene?: string
  sourceLanguage?: string
  targetLanguage?: string
  preserveFormatting?: boolean
  memoryEnabled?: boolean
  qualityCheck?: boolean
  glossaryCategory?: string
  units?: TranslateUnitRequest[]
  settings?: AiSettings
}

const TRANSLATE_STREAM_SESSIONS = new Map<string, AbortController>()

function pickSettings(req: TranslateBatchHttpRequest): AiSettings {
  return req.settings || defaultSettings
}

function resolveProvider(req: TranslateBatchHttpRequest): {
  provider: AiProviderId
  config: AiProviderConfig | undefined
} {
  const settings = pickSettings(req)
  const provider = settings.provider
  const config = settings.providers?.[provider]
  return { provider, config }
}

type CoreScope = 'selection' | 'document' | 'paragraph' | 'cell' | 'table'

function castScope(raw: string | undefined): CoreScope | undefined {
  return raw === 'selection' ||
    raw === 'document' ||
    raw === 'paragraph' ||
    raw === 'cell' ||
    raw === 'table'
    ? raw
    : undefined
}

function castRange(raw: TranslateUnitRequest['range']): {
  from?: number
  to?: number
  scope?: CoreScope
} | null {
  if (!raw) return null
  return { from: raw.from, to: raw.to, scope: castScope(raw.scope) }
}

function toCoreUnits(raw: TranslateUnitRequest[] | undefined): Array<{
  unitId: string
  kind: 'paragraph' | 'heading' | 'list-item' | 'table-cell' | 'document'
  sourceText: string
  order: number
  path?: string
  metadata?: Record<string, unknown>
  range?: { from?: number; to?: number; scope?: CoreScope } | null
}> {
  return (raw ?? []).map((u) => ({
    unitId: u.unitId ?? '',
    kind: (u.kind ?? 'paragraph') as 'paragraph' | 'heading' | 'list-item' | 'table-cell' | 'document',
    sourceText: u.sourceText ?? '',
    order: u.order ?? 0,
    path: u.path,
    metadata: u.metadata,
    range: castRange(u.range),
  }))
}

function writeSseEvent(
  response: ServerResponse,
  name: string,
  payload: unknown,
): void {
  try {
    response.write(`event: ${name}\n`)
    response.write(`data: ${JSON.stringify(payload)}\n\n`)
  } catch {
    /* socket may already be closed */
  }
}

async function readBody(request: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    request.on('data', (c: Buffer) => chunks.push(c))
    request.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')))
    request.on('error', reject)
  })
}

function sendJson(response: ServerResponse, status: number, payload: unknown): void {
  response.writeHead(status, { 'Content-Type': 'application/json' })
  response.end(JSON.stringify(payload))
}

/**
 * Handle POST /api/ai/translate — non-streaming batch.
 * Mirrors the `ai:translate-batch` IPC handler exactly; returns the final
 * TranslateBatchResponse shape with HTTP JSON.
 */
export async function handleTranslateBatchHttp(
  request: IncomingMessage,
  response: ServerResponse,
): Promise<void> {
  try {
    const raw = await readBody(request)
    const body = (raw ? JSON.parse(raw) : {}) as TranslateBatchHttpRequest
    const { provider, config } = resolveProvider(body)
    if (!config) {
      sendJson(response, 400, {
        error: { message: `AI provider "${provider}" not configured`, code: 'PROVIDER_NOT_CONFIGURED' },
      })
      return
    }
    const result = await translateBatch(
      {
        units: toCoreUnits(body.units),
        sourceLang: body.sourceLanguage,
        targetLang: body.targetLanguage ?? '',
        preserveFormat: body.preserveFormatting,
        scene: body.scene,
        memoryEnabled: body.memoryEnabled,
        qualityCheck: body.qualityCheck,
        glossaryCategory: body.glossaryCategory,
      },
      { provider, config },
    )
    sendJson(response, 200, result)
  } catch (error) {
    sendJson(response, 500, {
      error: { message: (error as Error)?.message ?? String(error) },
    })
  }
}

/**
 * Handle POST /api/ai/translate/stream — SSE pipeline.
 *
 * Emits `start` (with totalUnits), one `unit` event per settled unit, an
 * optional `quality` event after the batch completes, a `complete` event with
 * the final counts, and an `error` event if translation fails. The wire shape
 * matches Dataflare's `StructuredTranslationStreamEvent` so the existing
 * `web-bridge.ts` consumer can drop-in replace its target URL.
 */
export async function handleTranslateStreamHttp(
  request: IncomingMessage,
  response: ServerResponse,
): Promise<void> {
  const requestId = `stream-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
  let body: TranslateBatchHttpRequest = {}
  try {
    const raw = await readBody(request)
    body = (raw ? JSON.parse(raw) : {}) as TranslateBatchHttpRequest
  } catch (error) {
    sendJson(response, 400, { error: { message: 'invalid JSON body' } })
    return
  }

  const overrideRequestId = body.requestId?.trim()
  const effectiveRequestId = overrideRequestId || requestId

  response.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    Connection: 'keep-alive',
    'X-Request-Id': effectiveRequestId,
  })

  const abort = new AbortController()
  TRANSLATE_STREAM_SESSIONS.set(effectiveRequestId, abort)
  request.on('close', () => {
    abort.abort()
    TRANSLATE_STREAM_SESSIONS.delete(effectiveRequestId)
    try {
      response.end()
    } catch {
      /* ignore */
    }
  })

  const units = toCoreUnits(body.units)
  if (units.length === 0) {
    writeSseEvent(response, 'error', {
      type: 'error',
      requestId: effectiveRequestId,
      message: 'expected non-empty `units` array',
    })
    TRANSLATE_STREAM_SESSIONS.delete(effectiveRequestId)
    try {
      response.end()
    } catch {
      /* ignore */
    }
    return
  }

  const { provider, config } = resolveProvider(body)
  if (!config) {
    writeSseEvent(response, 'error', {
      type: 'error',
      requestId: effectiveRequestId,
      message: `AI provider "${provider}" not configured`,
    })
    TRANSLATE_STREAM_SESSIONS.delete(effectiveRequestId)
    try {
      response.end()
    } catch {
      /* ignore */
    }
    return
  }

  const startedAt = Date.now()
  writeSseEvent(response, 'start', {
    type: 'start',
    requestId: effectiveRequestId,
    sourceLanguage: body.sourceLanguage ?? 'auto',
    targetLanguage: body.targetLanguage ?? '',
    totalUnits: units.length,
  })

  let completed = 0
  let okCount = 0
  let memoryHitCount = 0
  let failedCount = 0

  try {
    const response_ = await translateBatchStream(
      {
        units,
        sourceLang: body.sourceLanguage,
        targetLang: body.targetLanguage ?? '',
        preserveFormat: body.preserveFormatting,
        scene: body.scene,
        memoryEnabled: body.memoryEnabled,
        qualityCheck: body.qualityCheck,
        glossaryCategory: body.glossaryCategory,
      },
      { provider, config },
      {
        concurrency: 25,
        onUnit: ({ index, total, result }) => {
          if (abort.signal.aborted) return
          completed += 1
          if (result.status === 'translated') okCount += 1
          else if (result.status === 'memory-hit') memoryHitCount += 1
          else if (result.status === 'failed') failedCount += 1
          const unitPayload: TranslateBatchUnitResult & {
            matchedTerms?: string[]
            warnings?: string[]
          } = {
            unitId: result.unitId,
            sourceText: result.sourceText,
            status: result.status,
            warnings: result.warnings,
            range: result.range,
          }
          if (result.translatedText !== undefined) unitPayload.translatedText = result.translatedText
          if (result.errorMessage) unitPayload.errorMessage = result.errorMessage
          writeSseEvent(response, 'unit', {
            type: 'unit',
            requestId: effectiveRequestId,
            unit: unitPayload,
            completedUnits: completed,
            totalUnits: total,
            progress: completed / Math.max(total, 1),
          })
        },
      },
    )
    if (response_.quality) {
      writeSseEvent(response, 'quality', {
        type: 'quality',
        requestId: effectiveRequestId,
        completedUnits: completed,
        quality: {
          overallScore: response_.quality.overallScore ?? 0,
          warnings: response_.quality.warnings ?? [],
          passed: (response_.quality.overallScore ?? 0) >= 0.85,
        },
      })
    }
    const finalStatus = response_.ok
      ? failedCount > 0
        ? 'partial'
        : 'completed'
      : 'failed'
    writeSseEvent(response, 'complete', {
      type: 'complete',
      requestId: effectiveRequestId,
      status: finalStatus,
      totalUnits: units.length,
      completedUnits: completed,
      okCount,
      memoryHitCount,
      failedCount,
      warnings: response_.quality?.warnings ?? [],
      elapsedMs: Date.now() - startedAt,
      ...(response_.error ? { errorMessage: response_.error } : {}),
    })
  } catch (error) {
    if (!abort.signal.aborted) {
      writeSseEvent(response, 'error', {
        type: 'error',
        requestId: effectiveRequestId,
        message: (error as Error)?.message ?? String(error),
      })
    }
  } finally {
    TRANSLATE_STREAM_SESSIONS.delete(effectiveRequestId)
    try {
      response.end()
    } catch {
      /* ignore */
    }
  }
}

/**
 * Handle POST /api/ai/translate/stream/cancel — abort an in-flight stream.
 */
export async function handleTranslateStreamCancelHttp(
  request: IncomingMessage,
  response: ServerResponse,
): Promise<void> {
  try {
    const raw = await readBody(request)
    const { requestId } = (raw ? JSON.parse(raw) : {}) as { requestId?: string }
    if (!requestId) {
      sendJson(response, 400, { error: { message: 'requestId required' } })
      return
    }
    const session = TRANSLATE_STREAM_SESSIONS.get(requestId)
    if (session) {
      session.abort()
      TRANSLATE_STREAM_SESSIONS.delete(requestId)
    }
    sendJson(response, 200, { ok: true, aborted: Boolean(session) })
  } catch (error) {
    sendJson(response, 500, {
      error: { message: (error as Error)?.message ?? String(error) },
    })
  }
}
