import type { AiConnectionStatus } from './connection'

export interface OllamaModelsResult {
  models: Array<{ name: string; parameterSize?: string; modifiedAt?: string }>
  error?: string
}

/**
 * Derive a user-facing connection status from an Ollama model-list result,
 * using the same status vocabulary as the connection test so the settings
 * dialog shows one consistent set of states. Network-level failures mean the
 * server is not running; HTTP-level failures mean the endpoint is wrong.
 * Never echoes the raw error string — the caller maps the status to a
 * localized message.
 */
export function ollamaListStatus(result: OllamaModelsResult): AiConnectionStatus {
  if (result.models.length > 0 || !result.error) return 'connected'
  const low = result.error.toLowerCase()
  if (/(econnrefused|econnreset|fetch failed|getaddrinfo|network|timed? ?out)/.test(low)) {
    return 'not-running'
  }
  if (/(status|http|4\d\d|5\d\d|unauthorized|forbidden)/.test(low)) return 'invalid'
  return 'unknown'
}

const OLLAMA_DEFAULT_MODELS_URL = 'http://localhost:11434'

export async function listOllamaModels(
  baseUrl?: string,
): Promise<OllamaModelsResult> {
  const root = (baseUrl ?? OLLAMA_DEFAULT_MODELS_URL)
    .replace(/\/v1\/?$/, '')
    .replace(/\/+$/, '')
  try {
    const resp = await fetch(`${root}/api/tags`, {
      method: 'GET',
      signal: AbortSignal.timeout(5_000),
    })
    if (!resp.ok) {
      return { models: [], error: `Ollama returned ${resp.status}` }
    }
    const data = (await resp.json()) as {
      models?: Array<{
        name: string
        size?: number
        modified_at?: string
        details?: { parameter_size?: string }
      }>
    }
    return {
      models: (data.models ?? []).map((m) => ({
        name: m.name,
        ...(m.details?.parameter_size ? { parameterSize: m.details.parameter_size } : {}),
        ...(m.modified_at ? { modifiedAt: m.modified_at } : {}),
      })),
    }
  } catch (err) {
    return { models: [], error: err instanceof Error ? err.message : String(err) }
  }
}
