export interface OllamaModelsResult {
  models: Array<{ name: string; parameterSize?: string; modifiedAt?: string }>
  error?: string
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
