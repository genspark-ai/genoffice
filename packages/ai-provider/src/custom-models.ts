// Live model list for the `custom` provider — a user-hosted endpoint (OMLX,
// Ollama, LM Studio, vLLM, OpenRouter…) — over the listing route its declared
// protocol speaks: `GET {baseUrl}/models` for OpenAI-compatible servers,
// `GET {baseUrl}/v1/models` for Anthropic-shaped gateways, `GET
// {baseUrl}/models` with the Google header for Gemini-shaped ones. A server
// that answers the wrong probe simply yields the empty catalogue, so the
// protocol-aware form degrades to the old behavior rather than breaking.
// `codex` gets the same list from its app-server's `model/list`; the settings
// screen turns either into a picker.
//
// Not `aiFetch`: that retries through the rescue fetch and warns on a network
// failure, which is right for a chat turn and wrong for a background probe of a
// local server that is very often simply not running. The user agent it would
// have added is still set, since gateways reject a bare `node`.
import { customProtocolBaseUrl } from './custom-profile'
import type { AiCustomModelProfile, AiCustomProtocol, CodexModelCatalog } from './types'
import { withUserAgent } from './fetch'
import { endpointUrl, readCappedResponseText } from './protocols/shared'
import { createStreamWatchdog } from './watchdog'

/** a server this slow is not usable for chat either, and the settings field must stay responsive */
const REQUEST_TIMEOUT_MS = 5000

/** ids past this are dropped; the picker is a settings control, not a catalog browser */
const MAX_MODELS = 2000

/** a model list is a few hundred KB at worst; past this the endpoint is not what it claims */
const MAX_BODY_BYTES = 2 * 1024 * 1024

/** "no list" — the caller leaves the free-text model box in place */
function emptyCatalog(): CodexModelCatalog {
  return { models: [], defaultModel: '' }
}

/** Pull the id out of one list entry in any of the three protocols' shapes. */
function entryId(entry: unknown): string | undefined {
  if (typeof entry === 'string') return entry
  if (typeof entry !== 'object' || entry === null) return undefined
  const e = entry as Record<string, unknown>
  // OpenAI: { id }; Anthropic: { id } already namespaced as "claude-…";
  // Gemini: { name: "models/gemini-3.5-flash" } — strip the resource prefix.
  const geminiName =
    typeof e.name === 'string' && e.name.startsWith('models/')
      ? e.name.slice('models/'.length)
      : undefined
  const id = typeof e.id === 'string' ? e.id : geminiName
  return id
}

/**
 * OpenRouter publishes every model a second time as `<id>:batch`: the
 * asynchronous batch tier, answered minutes to hours later and never
 * streamed. A chat picker cannot use it, so the variant is dropped; the
 * interactive `:free` / `:nitro` / `:thinking` variants stay.
 */
const isBatchVariant = (id: string): boolean => /:batch$/i.test(id)

const posInt = (v: unknown): number | undefined =>
  typeof v === 'number' && Number.isFinite(v) && v > 0 ? Math.floor(v) : undefined

/**
 * The capability declaration one list entry carries, when the endpoint
 * publishes any. Field names differ per shape:
 *
 * - OpenAI-compatible gateways (the rich ones — vLLM, and aggregators that
 *   serve many labs behind one key) answer `supports_images`,
 *   `max_output_tokens`, `context_length` and a display `name`;
 * - Gemini answers `inputTokenLimit` / `outputTokenLimit` and a resource
 *   `name`, and says nothing about image input;
 * - Anthropic answers a `display_name` and nothing else.
 *
 * Anything absent stays absent: the model declaration then falls through to
 * the endpoint switches and the shipped heuristics, exactly as before.
 */
function profileOf(entry: unknown): AiCustomModelProfile | undefined {
  const id = entryId(entry)?.trim()
  if (!id) return undefined
  if (typeof entry !== 'object' || entry === null) return { id }
  const e = entry as Record<string, unknown>
  const name = typeof e.name === 'string' && !e.name.startsWith('models/') ? e.name : undefined
  const displayName = typeof e.display_name === 'string' ? e.display_name : undefined
  const vision = typeof e.supports_images === 'boolean' ? e.supports_images : undefined
  const maxTokens = posInt(e.max_output_tokens) ?? posInt(e.outputTokenLimit)
  const contextWindow = posInt(e.context_length) ?? posInt(e.inputTokenLimit)
  return {
    id,
    ...(displayName || name ? { name: (displayName ?? name)!.slice(0, 128) } : {}),
    ...(vision !== undefined ? { vision } : {}),
    ...(maxTokens !== undefined ? { maxTokens } : {}),
    ...(contextWindow !== undefined ? { contextWindow } : {}),
  }
}

/** The listing route a protocol speaks, from the same base the chat request uses. */
function probeSpec(
  api: AiCustomProtocol,
  base: string,
  key: string,
): { url: string; headers: Record<string, string>; listKey: 'data' | 'models' } {
  if (api === 'anthropic') {
    // anthropicTurn appends v1/messages to the base; the list rides the same v1 route
    return {
      url: endpointUrl(base, 'v1/models'),
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      listKey: 'data',
    }
  }
  if (api === 'gemini') {
    // geminiTurn appends models/<id>:generateContent to a versioned root
    return {
      url: endpointUrl(base, 'models'),
      headers: { 'x-goog-api-key': key },
      listKey: 'models',
    }
  }
  return {
    url: endpointUrl(base, 'models'),
    headers: key ? { Authorization: `Bearer ${key}` } : {},
    listKey: 'data',
  }
}

/**
 * OpenRouter's `/models/user` is the public catalog filtered by the account's
 * own provider allow/ignore settings — the list the key can actually route to.
 * Only that host has the route; a keyless probe (or a refused key) falls back
 * to the public `/models`.
 */
function keyedCatalogUrl(base: string, key: string): string | undefined {
  if (!key) return undefined
  try {
    if (new URL(base).hostname !== 'openrouter.ai') return undefined
  } catch {
    return undefined
  }
  return endpointUrl(base, 'models/user')
}

/**
 * The model ids a user-hosted endpoint advertises, probed on the route its
 * declared protocol speaks (absent = the historic OpenAI-compatible shape).
 *
 * Never throws and never rejects: a refused, unreachable, slow, oversized,
 * non-JSON or unexpectedly shaped answer all return an empty catalogue, because
 * the only consumer is a settings screen deciding between a dropdown and a text
 * box. An endpoint that cannot answer must not break the field the user types
 * into, and must not make noise about it either.
 *
 * The key is sent under the protocol's own header (Bearer, x-api-key, or
 * x-goog-api-key) only when one is given — an empty `Bearer ` is a malformed
 * credential that strict servers reject, and local servers usually need no key
 * at all.
 *
 * Server order is preserved, duplicate, non-string and blank ids are dropped,
 * and `defaultModel` is always `''`: `/models` carries no `isDefault` marker to
 * read, and real lists mix embedding and conversion models in with the chat
 * ones, so the first entry is a guess that can be actively wrong.
 *
 * Callers hand it already-coerced arguments; `listCustomModelsForIpc` owns the
 * validation of anything arriving from outside.
 */
export async function listCustomModels(
  baseUrl: string,
  apiKey?: string,
  api?: AiCustomProtocol,
): Promise<CodexModelCatalog> {
  if (!baseUrl.trim()) return emptyCatalog()
  const key = (apiKey ?? '').trim()

  const watchdog = createStreamWatchdog(undefined, REQUEST_TIMEOUT_MS)
  try {
    return await watchdog.guard(async () => {
      // Inside the guard: an unparsable base URL is the old behavior's "empty
      // catalogue", not a rejection. The base is composed exactly as the
      // request path composes it (customProtocolBaseUrl), so the probe can
      // never ask a route the chat turn would not use.
      const base = customProtocolBaseUrl(api ?? 'openai-compatible', baseUrl)
      const probe = probeSpec(api ?? 'openai-compatible', base, key)
      const get = (url: string) =>
        fetch(
          url,
          withUserAgent({
            headers: { Accept: 'application/json', ...probe.headers },
            signal: watchdog.signal,
          }),
        )
      const keyedUrl =
        (api ?? 'openai-compatible') === 'openai-compatible' && keyedCatalogUrl(base, key)
      let response = keyedUrl ? await get(keyedUrl) : undefined
      if (!response?.ok) response = await get(probe.url)
      if (!response.ok) return emptyCatalog()
      const text = await readCappedResponseText(response, { maxBytes: MAX_BODY_BYTES })
      const body = JSON.parse(text) as { data?: unknown; models?: unknown } | null
      // OpenAI and Anthropic answer { data: [...] }; Gemini answers { models: [...] }
      const list = body?.[probe.listKey]
      if (!Array.isArray(list)) return emptyCatalog()
      const seen = new Set<string>()
      const profiles: AiCustomModelProfile[] = []
      for (const entry of list) {
        const id = entryId(entry)?.trim()
        if (!id || seen.has(id) || isBatchVariant(id)) continue
        seen.add(id)
        const profile = profileOf(entry)
        // an entry that declares nothing beyond its id is not a declaration
        if (profile && Object.keys(profile).length > 1) profiles.push(profile)
        if (seen.size >= MAX_MODELS) break
      }
      return {
        models: [...seen],
        defaultModel: '',
        ...(profiles.length > 0 ? { profiles } : {}),
      }
    })
  } catch {
    return emptyCatalog()
  }
}

/**
 * `listCustomModels` behind an `unknown` IPC payload.
 *
 * The validation lives here rather than in the main process so the hook there
 * stays one line: `{ baseUrl: string; apiKey?: string; api?: string }`, anything
 * else coerced to "no base URL", which returns the empty catalogue.
 */
export function listCustomModelsForIpc(input: unknown): Promise<CodexModelCatalog> {
  const raw = (input ?? {}) as { baseUrl?: unknown; apiKey?: unknown; api?: unknown }
  const api =
    raw.api === 'anthropic' || raw.api === 'gemini' || raw.api === 'openai-compatible'
      ? raw.api
      : undefined
  return listCustomModels(
    typeof raw.baseUrl === 'string' ? raw.baseUrl : '',
    typeof raw.apiKey === 'string' ? raw.apiKey : undefined,
    api,
  )
}
