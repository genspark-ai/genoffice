/**
 * Normalizer for the translation HTTP response used by the embedded
 * (Dataflare-hosted) branches of `web-bridge.ts`.
 *
 * Wire format history:
 *
 *   - 旧链路：`/crmapi/ai/translation/v1/translate`（Dataflare Spring Boot）
 *     返回 `{ code: 0, msg?, data: { units, quality } }`。
 *   - 现链路：`/office-engine/api/ai/translate`（GenOffice web-server，
 *     见 `apps/web-server/src/ai/translate-http.ts`）返回
 *     `{ ok: true, units, quality }`，没有 `code` / `data` 信封。
 *
 * URL 迁移时调用方仍在解析旧信封（`body.code === 0` + `body.data.units`），
 * 导致嵌入模式下 `translatedText` 永远解析不到、单条与批量翻译都报
 * "Dataflare translation failed"。这里两种结构都接受，避免再出现同类回归。
 */

export interface DataflareTranslateUnit {
  unitId: string
  sourceText: string
  translatedText?: string
  status?: string
  matchedTerms?: string[]
  warnings?: string[]
  errorMessage?: string
}

export interface DataflareTranslateQuality {
  overallScore?: number
  warnings?: string[]
}

export interface DataflareTranslatePayload {
  ok: boolean
  units: DataflareTranslateUnit[]
  /** 旧信封里的 requestId；GenOffice 现链路由调用方自己生成，可缺省。 */
  requestId?: string
  quality?: DataflareTranslateQuality
  error?: string
}

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === 'object' ? (value as Record<string, unknown>) : null
}

function asStringArray(value: unknown): string[] | undefined {
  if (!Array.isArray(value)) return undefined
  return value.filter((item): item is string => typeof item === 'string')
}

function normalizeUnit(raw: unknown): DataflareTranslateUnit | null {
  const unit = asRecord(raw)
  if (!unit) return null
  const unitId = typeof unit.unitId === 'string' ? unit.unitId : ''
  const sourceText = typeof unit.sourceText === 'string' ? unit.sourceText : ''
  return {
    unitId,
    sourceText,
    translatedText: typeof unit.translatedText === 'string' ? unit.translatedText : undefined,
    status: typeof unit.status === 'string' ? unit.status : undefined,
    matchedTerms: asStringArray(unit.matchedTerms),
    warnings: asStringArray(unit.warnings),
    errorMessage: typeof unit.errorMessage === 'string' ? unit.errorMessage : undefined,
  }
}

/**
 * Accepts both the GenOffice (`{ ok, units, quality }`) and the legacy
 * Dataflare (`{ code, data: { units, quality } }`) payloads.
 *
 * `ok` means "服务端声明成功且至少带回一个 unit"，调用方仍需自行判断
 * 每个 unit 是否真的带回了 `translatedText`（部分成功时 `ok` 仍为 true）。
 */
export function parseDataflareTranslateResponse(body: unknown): DataflareTranslatePayload {
  const root = asRecord(body)
  if (!root) {
    return { ok: false, units: [], error: 'empty translation response' }
  }

  const envelope = asRecord(root.data)
  const payload = envelope ?? root
  const legacyCode = typeof root.code === 'number' ? root.code : undefined
  const envelopeOk = envelope ? legacyCode === 0 : root.ok !== false

  const units = (Array.isArray(payload.units) ? payload.units : [])
    .map(normalizeUnit)
    .filter((unit): unit is DataflareTranslateUnit => unit !== null)

  const quality = asRecord(payload.quality)
  const message =
    (typeof root.msg === 'string' && root.msg) ||
    (typeof root.error === 'string' && root.error) ||
    (typeof payload.error === 'string' && payload.error) ||
    undefined

  const requestId =
    (typeof root.requestId === 'string' && root.requestId) ||
    (typeof payload.requestId === 'string' && payload.requestId) ||
    undefined

  return {
    ok: envelopeOk && units.length > 0,
    units,
    requestId: requestId || undefined,
    quality: quality
      ? {
          overallScore: typeof quality.overallScore === 'number' ? quality.overallScore : undefined,
          warnings: asStringArray(quality.warnings),
        }
      : undefined,
    error: message || undefined,
  }
}
