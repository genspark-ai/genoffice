import { listAgyModelsForIpc } from './agy-cli'
import { AI_PROVIDERS } from './providers'
import { getProviderAdapter } from './registry'
import { listCustomModels, listGeminiModelsForIpc } from './custom-models'
import { withUserAgent } from './fetch'
import { endpointUrl, readCappedResponseText } from './protocols/shared'
import type { AiProviderConfig, AiProviderId, CodexModelCatalog } from './types'

const EMPTY: CodexModelCatalog = { models: [], defaultModel: '' }

/** Background discovery uses the unsaved key and endpoint, without issuing a chat call. */
export async function listProviderModelsForIpc(input: unknown): Promise<CodexModelCatalog> {
  const raw = input as { provider?: unknown; config?: unknown } | null
  if (!AI_PROVIDERS.some((entry) => entry.id === raw?.provider)) return EMPTY
  const provider = raw!.provider as AiProviderId
  if (provider === 'agy') return listAgyModelsForIpc(raw!.config)
  if (provider === 'local-vlm') {
    const meta = AI_PROVIDERS.find((entry) => entry.id === 'local-vlm')
    return { models: meta?.models ?? [], defaultModel: meta?.defaultModel ?? '' }
  }
  if (provider === 'genspark' || provider === 'codex') return EMPTY
  const stored = raw!.config as Partial<AiProviderConfig> | null
  const config: AiProviderConfig = {
    apiKey: typeof stored?.apiKey === 'string' ? stored.apiKey.trim() : '',
    model: typeof stored?.model === 'string' ? stored.model : '',
    baseUrl: typeof stored?.baseUrl === 'string' ? stored.baseUrl.trim() || undefined : undefined,
  }
  if (!config.apiKey && provider !== 'custom' && provider !== 'openrouter' && provider !== 'opper')
    return EMPTY
  try {
    const endpoint = getProviderAdapter(provider).resolveEndpoint(config)
    if (provider === 'gemini') {
      return await listGeminiModelsForIpc({ apiKey: config.apiKey, baseUrl: endpoint.baseUrl })
    }
    if (provider === 'opper' && !config.baseUrl) {
      const models = new Set<string>()
      const signal = AbortSignal.timeout(5000)
      for (let offset = 0; offset < 2000; offset += 100) {
        const response = await fetch(
          `https://api.opper.ai/v3/models?type=llm&limit=100&offset=${offset}`,
          withUserAgent({ headers: { Accept: 'application/json' }, signal }),
        )
        if (!response.ok) return EMPTY
        const body = JSON.parse(
          await readCappedResponseText(response, { maxBytes: 2 * 1024 * 1024 }),
        ) as {
          models?: Array<{ id?: unknown }>
          total?: number
        }
        if (!Array.isArray(body.models)) return EMPTY
        for (const entry of body.models) {
          if (typeof entry?.id === 'string' && entry.id.trim()) models.add(entry.id.trim())
        }
        if (body.models.length < 100 || offset + 100 >= (body.total ?? Infinity)) break
      }
      return { models: [...models], defaultModel: '' }
    }
    if (endpoint.protocol !== 'anthropic') {
      return await listCustomModels(endpoint.baseUrl, config.apiKey)
    }
    // Anthropic-compatible APIs use a different auth header and cursor pagination.
    const models = new Set<string>()
    const signal = AbortSignal.timeout(5000)
    let afterId = ''
    for (let page = 0; page < 20; page++) {
      const url = new URL(endpointUrl(endpoint.baseUrl, 'v1/models'))
      url.searchParams.set('limit', '1000')
      if (afterId) url.searchParams.set('after_id', afterId)
      const response = await fetch(
        url,
        withUserAgent({
          headers: {
            Accept: 'application/json',
            'x-api-key': config.apiKey,
            'anthropic-version': '2023-06-01',
          },
          signal,
        }),
      )
      if (!response.ok) return EMPTY
      const body = JSON.parse(
        await readCappedResponseText(response, { maxBytes: 2 * 1024 * 1024 }),
      ) as {
        data?: Array<{ id?: unknown }>
        has_more?: boolean
        last_id?: string
      }
      if (!Array.isArray(body.data)) return EMPTY
      for (const entry of body.data) {
        if (typeof entry?.id === 'string' && entry.id.trim()) models.add(entry.id.trim())
        if (models.size >= 2000) return { models: [...models], defaultModel: '' }
      }
      if (!body.has_more) return { models: [...models], defaultModel: '' }
      if (!body.last_id || body.last_id === afterId) return EMPTY
      afterId = body.last_id
    }
  } catch {
    // Unsupported listing, network failures and invalid keys leave manual entry available.
  }
  return EMPTY
}
