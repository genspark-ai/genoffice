// The declarative profile of a user-hosted endpoint, resolved into the request
// shaping the protocol layer consumes. This is what lets a self-hosted gateway
// (or a regional mirror of a domestic vendor) be described entirely in the
// settings file: wire protocol, static headers, credential indirection, a
// relaxed idle budget for slow local servers, the vendor switches no shipped
// model-name heuristic can know, and a per-model capability list for releases
// the catalog has not caught up with.
//
// Resolution follows one documented fallback order every field shares:
//   model declaration → endpoint switches → shipped heuristic → old default.
// An endpoint that declares nothing therefore behaves exactly as it did before
// this existed, so a pre-profile settings file needs no migration and no
// behavior change.
import { modelHasFixedSampling } from './model-heuristics'
import { appendPath, normalizeBaseUrl, stripTrailingV1 } from './url-compose'
import type {
  AiCustomCompat,
  AiCustomModelProfile,
  AiCustomProfile,
  AiCustomProtocol,
} from './types'

/**
 * Compose the request-time base URL for a user-hosted endpoint, per protocol.
 *
 * The openai-compatible branch keeps the historic contract untouched — the
 * URL is used as stored (the UI's `https://…/v1` placeholder teaches that
 * shape) — so every endpoint saved before profiles existed keeps working byte
 * for byte. The two branches profiles add are new, so they may be lenient
 * where the copy-paste habits of gateways disagree:
 *
 * - `anthropic`: the protocol appends `v1/messages`, so a pasted
 *   `https://gw.example/v1` (the shape vendor docs print) would otherwise
 *   double the segment. Strip the trailing `/v1` and let the protocol put it
 *   back exactly once.
 * - `gemini`: the protocol appends `models/<id>:generateContent` under a
 *   versioned root. A bare host gets the documented `/v1beta`; a host with
 *   any path (a gateway mounted at `/v1`, `/api/google/v1beta`, …) is used
 *   as stored.
 *
 * The model-list probe composes its route from the same function, so the
 * settings pane can never ask a different address than the one a chat turn
 * uses.
 */
export function customProtocolBaseUrl(protocol: AiCustomProtocol, raw: string): string {
  const base = normalizeBaseUrl(raw, '')
  if (protocol === 'anthropic') return stripTrailingV1(base)
  if (protocol === 'gemini') {
    const url = new URL(base)
    if (url.pathname.replace(/\/+$/, '') === '') {
      return appendPath(base.replace(/\/+$/, ''), '/v1beta')
    }
    return base
  }
  return base
}
/** The protocols a user-hosted endpoint may declare, in picker order. */
export const AI_CUSTOM_PROTOCOLS: readonly AiCustomProtocol[] = [
  'openai-compatible',
  'anthropic',
  'gemini',
] as const

const CUSTOM_PROTOCOLS: ReadonlySet<string> = new Set<string>(AI_CUSTOM_PROTOCOLS)

/** Wire field limits a stored profile may carry; oversized values are user error, not data. */
export const MAX_PROFILE_HEADERS = 16
export const MAX_HEADER_NAME_LENGTH = 128
export const MAX_HEADER_VALUE_LENGTH = 512
export const MAX_PROFILE_MODELS = 200
/** matches the env-var naming a shell can actually export */
export const MAX_ENV_NAME_LENGTH = 128
/** a body-extras blob past this is not a vendor switch, it is payload stuffing */
export const MAX_BODY_EXTRAS_JSON = 65_536

const str = (v: unknown): string => (typeof v === 'string' ? v.trim() : '')
const bool = (v: unknown): boolean | undefined => (typeof v === 'boolean' ? v : undefined)
/** A positive integer, floored before the range check so 0.5 floors to 0 and is dropped */
const posInt = (v: unknown): number | undefined => {
  if (typeof v !== 'number' || !Number.isFinite(v)) return undefined
  const n = Math.floor(v)
  return n > 0 ? n : undefined
}

/**
 * A string→string record (headers), dropping any pair that is malformed or
 * not representable as an HTTP header: `new Headers` rejects CRLF and control
 * bytes in values and non-token characters in names, so a stored profile can
 * never smuggle a second header line into a request.
 */
function stringRecord(raw: unknown, maxKeys: number): Record<string, string> | undefined {
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) return undefined
  const out: Record<string, string> = {}
  for (const [key, value] of Object.entries(raw as Record<string, unknown>)) {
    if (Object.keys(out).length >= maxKeys) break
    const name = str(key)
    const text = typeof value === 'string' ? value.trim() : ''
    if (!name || !text) continue
    const candidate = {
      name: name.slice(0, MAX_HEADER_NAME_LENGTH),
      value: text.slice(0, MAX_HEADER_VALUE_LENGTH),
    }
    try {
      new Headers([[candidate.name, candidate.value]])
    } catch {
      continue
    }
    out[candidate.name] = candidate.value
  }
  return Object.keys(out).length > 0 ? out : undefined
}

/** The compat switches, keeping only keys that carry a real boolean. */
function compatOf(raw: unknown): AiCustomCompat | undefined {
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) return undefined
  const src = raw as Record<string, unknown>
  const out: AiCustomCompat = {}
  for (const field of [
    'omitTemperature',
    'echoReasoning',
    'vision',
    'useMaxCompletionTokens',
  ] as const) {
    const value = bool(src[field])
    if (value !== undefined) out[field] = value
  }
  return Object.keys(out).length > 0 ? out : undefined
}

function modelProfileOf(raw: unknown): AiCustomModelProfile | undefined {
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) return undefined
  const src = raw as Record<string, unknown>
  const id = str(src.id)
  if (!id) return undefined
  const alias = str(src.alias)
  const name = str(src.name)
  const vision = bool(src.vision)
  const omitTemperature = bool(src.omitTemperature)
  const echoReasoning = bool(src.echoReasoning)
  const useMaxCompletionTokens = bool(src.useMaxCompletionTokens)
  const maxTokens = posInt(src.maxTokens)
  const contextWindow = posInt(src.contextWindow)
  return {
    id: id.slice(0, 256),
    ...(alias ? { alias: alias.slice(0, 256) } : {}),
    ...(name ? { name: name.slice(0, 128) } : {}),
    ...(vision !== undefined ? { vision } : {}),
    ...(omitTemperature !== undefined ? { omitTemperature } : {}),
    ...(echoReasoning !== undefined ? { echoReasoning } : {}),
    ...(useMaxCompletionTokens !== undefined ? { useMaxCompletionTokens } : {}),
    ...(maxTokens !== undefined ? { maxTokens } : {}),
    ...(contextWindow !== undefined ? { contextWindow } : {}),
  }
}

/**
 * A stored profile narrowed to the shape this module can honor. Unknown or
 * malformed fields are dropped rather than rejecting the payload: a
 * hand-edited settings file with one bad key must not disable the endpoint's
 * other settings. Returns undefined for an empty/absent profile so callers can
 * tell "no profile" from "profile with defaults".
 */
export function normalizeCustomProfile(raw: unknown): AiCustomProfile | undefined {
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) return undefined
  const src = raw as Record<string, unknown>
  const out: AiCustomProfile = {}
  if (typeof src.api === 'string' && CUSTOM_PROTOCOLS.has(src.api.trim())) {
    out.api = src.api.trim() as AiCustomProtocol
  }
  const apiKeyEnv = str(src.apiKeyEnv)
  if (apiKeyEnv) out.apiKeyEnv = apiKeyEnv.slice(0, MAX_ENV_NAME_LENGTH)
  const headers = stringRecord(src.headers, MAX_PROFILE_HEADERS)
  if (headers) out.headers = headers
  const idleTimeoutMs = posInt(src.idleTimeoutMs)
  if (idleTimeoutMs !== undefined) out.idleTimeoutMs = idleTimeoutMs
  const compat = compatOf(src.compat)
  if (compat) out.compat = compat
  const bodyExtras =
    typeof src.bodyExtras === 'object' && src.bodyExtras !== null && !Array.isArray(src.bodyExtras)
      ? (src.bodyExtras as Record<string, unknown>)
      : undefined
  if (bodyExtras && Object.keys(bodyExtras).length > 0) {
    let size: number
    try {
      size = JSON.stringify(bodyExtras).length
    } catch {
      size = Number.POSITIVE_INFINITY // circular: refuse the whole blob
    }
    if (size <= MAX_BODY_EXTRAS_JSON) out.bodyExtras = bodyExtras
  }
  if (Array.isArray(src.modelProfiles)) {
    const models = src.modelProfiles
      .map(modelProfileOf)
      .filter((m): m is AiCustomModelProfile => m !== undefined)
      .slice(0, MAX_PROFILE_MODELS)
    if (models.length > 0) out.modelProfiles = models
  }
  if (Array.isArray(src.probedProfiles)) {
    const probed = src.probedProfiles
      .map(modelProfileOf)
      .filter((m): m is AiCustomModelProfile => m !== undefined)
      .slice(0, MAX_PROFILE_MODELS)
    if (probed.length > 0) out.probedProfiles = probed
  }
  return Object.keys(out).length > 0 ? out : undefined
}

/**
 * The endpoint's declaration for a model — the user's own first, then the
 * probe's cache — matched on the configured id: exact first (the user spelled
 * the id they use), then case-insensitive, then a prefix match so a
 * declaration for `qwen3.5-27b` also covers `qwen3.5-27b-coder`.
 */
export function declaredModelProfile(
  profile: AiCustomProfile | undefined,
  model: string,
): AiCustomModelProfile | undefined {
  if (!model) return undefined
  const lower = model.toLowerCase()
  const match = (list: AiCustomModelProfile[] | undefined): AiCustomModelProfile | undefined => {
    if (!list?.length) return undefined
    return (
      list.find((m) => m.id === model) ??
      list.find((m) => m.id.toLowerCase() === lower) ??
      list.find((m) => lower.startsWith(m.id.toLowerCase()))
    )
  }
  return match(profile?.modelProfiles) ?? match(profile?.probedProfiles)
}

/** A model switch contradicts nothing but the heuristic: the model speaks first, then the endpoint. */
function pick(declared: boolean | undefined, endpoint: boolean | undefined): boolean | undefined {
  return declared !== undefined ? declared : endpoint
}

/** Every request-shaping decision resolved from one endpoint profile + model. */
export interface ResolvedCustomProfile {
  /** the wire protocol to speak, defaulted for a pre-profile endpoint */
  protocol: AiCustomProtocol
  /** the id to put on the wire when the endpoint spells the configured one differently */
  wireModel?: string | undefined
  /** omit our hard-coded temperature (declared, else the fixed-sampling family) */
  omitTemperature: boolean
  /** tri-state: true/false is an instruction to the protocol; undefined means
   * "not declared, apply the shipped interleaved-thinking heuristic" */
  echoReasoning?: boolean | undefined
  /** send the output cap under the renamed OpenAI field */
  useMaxCompletionTokens: boolean
  /** images may be sent here — `vision` only when the user declared it, so an
   * endpoint that says nothing keeps the pre-profile behavior byte for byte */
  vision?: boolean | undefined
  /** the model's own output ceiling, when declared */
  maxTokens?: number | undefined
  headers?: Record<string, string> | undefined
  idleTimeoutMs?: number | undefined
  bodyExtras?: Record<string, unknown> | undefined
}

/**
 * Resolve one custom turn's shaping from its profile. Pass the model id that is
 * about to go on the wire; an empty model resolves to the fall-through defaults.
 */
export function resolveCustomProfile(
  profile: AiCustomProfile | undefined,
  model: string,
): ResolvedCustomProfile {
  const compat = profile?.compat
  const declared = declaredModelProfile(profile, model)
  const visionDeclared = pick(declared?.vision, compat?.vision)
  return {
    protocol: profile?.api ?? 'openai-compatible',
    ...(declared?.alias ? { wireModel: declared.alias } : {}),
    omitTemperature:
      pick(declared?.omitTemperature, compat?.omitTemperature) ?? modelHasFixedSampling(model),
    // Deliberately tri-state: undefined means "not declared, use the shipped
    // family heuristic" (the protocol's own fallback), and a declared false is
    // an order to stay silent even on a family the heuristic would echo.
    ...(pick(declared?.echoReasoning, compat?.echoReasoning) !== undefined
      ? { echoReasoning: pick(declared?.echoReasoning, compat?.echoReasoning) }
      : {}),
    useMaxCompletionTokens:
      pick(declared?.useMaxCompletionTokens, compat?.useMaxCompletionTokens) === true,
    // Only a declaration makes this field exist: an endpoint that says nothing
    // keeps the pre-profile resolved shape byte for byte, and the callers'
    // shipped fallbacks (the text-only model list) answer unchanged.
    ...(visionDeclared !== undefined ? { vision: visionDeclared } : {}),
    ...(declared?.maxTokens !== undefined ? { maxTokens: declared.maxTokens } : {}),
    ...(profile?.headers ? { headers: profile.headers } : {}),
    ...(profile?.idleTimeoutMs !== undefined ? { idleTimeoutMs: profile.idleTimeoutMs } : {}),
    ...(profile?.bodyExtras ? { bodyExtras: profile.bodyExtras } : {}),
  }
}
