import { AI_PROVIDERS } from './providers'
import { assertValidAiEndpoint } from './endpoint-policy'
import type {
  AiModel,
  AiModelCatalog,
  AiModelDiscoveryOptions,
  AiProviderConfig,
  AiProviderId,
} from './types'

const DEFAULT_TIMEOUT_MS = 15_000
const MAX_MODEL_PAGES = 100
const MAX_DISCOVERED_MODELS = 10_000

function metaFor(provider: AiProviderId) {
  return AI_PROVIDERS.find((meta) => meta.id === provider)
}

function endpointFor(
  provider: AiProviderId,
  config: AiProviderConfig,
  options?: AiModelDiscoveryOptions,
): string {
  const meta = metaFor(provider)
  const baseUrl = options?.baseUrl ?? config.baseUrl ?? meta?.defaultBaseUrl
  if (!baseUrl) throw new Error(`Provider ${provider} requires a Base URL`)
  return assertValidAiEndpoint(baseUrl)
}

function authHeaders(provider: AiProviderId, apiKey: string): Record<string, string> {
  if (!apiKey) return {}
  if (provider === 'anthropic') {
    return {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true',
    }
  }
  if (provider === 'gemini') return { 'x-goog-api-key': apiKey }
  return { Authorization: `Bearer ${apiKey}` }
}

async function fetchJson(
  url: string,
  headers: Record<string, string>,
  options: AiModelDiscoveryOptions | undefined,
  init: Omit<RequestInit, 'headers' | 'signal'> = {},
): Promise<unknown> {
  const controller = new AbortController()
  const timeout = setTimeout(() => controller.abort(), options?.timeoutMs ?? DEFAULT_TIMEOUT_MS)
  const onAbort = () => controller.abort()
  options?.signal?.addEventListener('abort', onAbort, { once: true })
  try {
    let response: Response
    try {
      response = await fetch(url, {
        method: 'GET',
        ...init,
        headers: { Accept: 'application/json', ...headers },
        signal: controller.signal,
      })
    } catch (error) {
      if (controller.signal.aborted) {
        if (options?.signal?.aborted) throw new Error('Model discovery canceled.', { cause: error })
        const seconds = Math.ceil((options?.timeoutMs ?? DEFAULT_TIMEOUT_MS) / 1_000)
        throw new Error(`Model discovery timed out after ${seconds} seconds. Try again.`, {
          cause: error,
        })
      }
      throw error
    }
    const text = await response.text()
    let body: unknown
    try {
      body = text ? (JSON.parse(text) as unknown) : undefined
    } catch {
      body = text
    }
    if (!response.ok) {
      const detail = typeof body === 'string' ? body : JSON.stringify(body)
      throw new Error(`Model discovery HTTP ${response.status}: ${detail.slice(0, 500)}`)
    }
    return body
  } finally {
    clearTimeout(timeout)
    options?.signal?.removeEventListener('abort', onAbort)
  }
}

function responseModels(body: unknown): unknown[] {
  if (Array.isArray(body)) return body
  if (Array.isArray((body as { data?: unknown[] })?.data)) {
    return (body as { data: unknown[] }).data
  }
  if (Array.isArray((body as { models?: unknown[] })?.models)) {
    return (body as { models: unknown[] }).models
  }
  if (Array.isArray((body as { results?: unknown[] })?.results)) {
    return (body as { results: unknown[] }).results
  }
  return []
}

function modelListUrl(provider: AiProviderId, baseUrl: string, path: string): string {
  const url = new URL(`${baseUrl}${path.startsWith('/') ? path : `/${path}`}`)
  if (provider === 'gemini') url.searchParams.set('pageSize', '1000')
  if (provider === 'anthropic' || provider === 'replicate') url.searchParams.set('limit', '100')
  return url.toString()
}

function nextModelPageUrl(
  provider: AiProviderId,
  currentUrl: string,
  body: unknown,
): string | undefined {
  if (!body || typeof body !== 'object') return undefined
  const record = body as Record<string, unknown>
  const next = typeof record.next === 'string' ? record.next : undefined
  if (next) return next
  const current = new URL(currentUrl)
  if (provider === 'gemini' && typeof record.nextPageToken === 'string') {
    current.searchParams.set('pageToken', record.nextPageToken)
    return current.toString()
  }
  if (provider === 'anthropic' && record.has_more === true && typeof record.last_id === 'string') {
    current.searchParams.set('after_id', record.last_id)
    return current.toString()
  }
  return undefined
}

function sameOriginPage(nextUrl: string, firstUrl: string): string {
  const parsed = new URL(nextUrl)
  if (parsed.hash || parsed.username || parsed.password) {
    throw new Error('Model pagination returned an unsafe URL')
  }
  const next = new URL(assertValidAiEndpoint(`${parsed.origin}${parsed.pathname}`))
  next.search = parsed.search
  const first = new URL(firstUrl)
  if (next.origin !== first.origin) throw new Error('Model pagination changed API origin')
  return next.toString()
}

async function fetchAllModelPages(
  provider: AiProviderId,
  firstUrl: string,
  headers: Record<string, string>,
  options?: AiModelDiscoveryOptions,
): Promise<unknown[]> {
  const models: unknown[] = []
  const visited = new Set<string>()
  let url: string | undefined = firstUrl
  for (let page = 0; url && page < MAX_MODEL_PAGES; page += 1) {
    if (visited.has(url)) throw new Error('Model pagination returned a repeated cursor')
    visited.add(url)
    const body = await fetchJson(url, headers, options)
    models.push(...responseModels(body))
    if (models.length >= MAX_DISCOVERED_MODELS) return models.slice(0, MAX_DISCOVERED_MODELS)
    const next = nextModelPageUrl(provider, url, body)
    url = next ? sameOriginPage(next, firstUrl) : undefined
  }
  if (url) throw new Error(`Model catalog exceeded ${MAX_MODEL_PAGES} pages`)
  return models
}

function stringArray(value: unknown): string[] {
  return Array.isArray(value)
    ? value.filter((item): item is string => typeof item === 'string')
    : []
}

function normalizeOpenAiModel(
  provider: AiProviderId,
  raw: Record<string, unknown>,
): AiModel | undefined {
  const id = typeof raw.id === 'string' ? raw.id : ''
  if (!id) return undefined
  const architecture = (raw.architecture ?? {}) as Record<string, unknown>
  const inputModalities = stringArray(architecture.input_modalities ?? raw.input_modalities)
  const outputModalities = stringArray(architecture.output_modalities ?? raw.output_modalities)
  const supportedParameters = stringArray(raw.supported_parameters)
  const supportsTools =
    supportedParameters.some((parameter) =>
      /^(tools|tool_choice|parallel_tool_calls)$/.test(parameter),
    ) || raw.supportsTools === true
  const hasImageInput = inputModalities.some((modality) => /image/i.test(modality))
  const rawCapability =
    raw.capabilities && typeof raw.capabilities === 'object'
      ? (raw.capabilities as Record<string, unknown>)
      : {}
  const family = [id, raw.type, raw.category, raw.task]
    .filter((value): value is string => typeof value === 'string')
    .join(' ')
  const inferredImageOutput =
    metaFor(provider)?.imageModels?.includes(id) === true ||
    /(?:^|[/_. -])(dall-e|flux|imagen|gpt-image|grok(?:-[\w.]+)*-image|grok-imagine|imagegen|seedream|stable-diffusion|recraft|ideogram|hidream|qwen-image|nano-banana)(?:$|[/_. -])/i.test(
      family,
    ) ||
    /text[-_ ]to[-_ ]image|image[-_ ]generation/i.test(family)
  const hasImageOutput =
    outputModalities.some((modality) => /image/i.test(modality)) || inferredImageOutput
  const nonChat = /embed|embedding|moderation|rerank|speech|transcri|ocr/i.test(id)
  const supportsChat =
    typeof rawCapability.completion_chat === 'boolean'
      ? rawCapability.completion_chat
      : !nonChat && !hasImageOutput
  const metadata = { ...raw }
  const displayName = typeof raw.name === 'string' ? raw.name : undefined
  const createdAt = typeof raw.created === 'number' ? raw.created : undefined
  const contextLength =
    typeof raw.context_length === 'number'
      ? raw.context_length
      : typeof raw.contextLength === 'number'
        ? raw.contextLength
        : undefined
  return {
    id,
    ...(displayName ? { displayName } : {}),
    ...(createdAt !== undefined ? { createdAt } : {}),
    ...(contextLength !== undefined ? { contextLength } : {}),
    inputModalities,
    outputModalities,
    capabilities: {
      ...(supportsChat ? { chat: true } : {}),
      ...(supportsTools ? { tools: true } : {}),
      ...(hasImageInput ? { vision: true } : {}),
      ...(hasImageOutput || (provider === 'openrouter' && outputModalities.includes('image'))
        ? { imageGeneration: true }
        : {}),
    },
    providerMetadata: metadata,
  }
}

function normalizeGeminiModel(raw: Record<string, unknown>): AiModel | undefined {
  const rawName = typeof raw.name === 'string' ? raw.name : ''
  const id = rawName.replace(/^models\//, '')
  if (!id) return undefined
  const methods = stringArray(raw.supportedGenerationMethods)
  const modalities = stringArray(raw.outputModalities)
  const supportsImage =
    modalities.some((modality) => /image/i.test(modality)) ||
    /(?:^|[-/])(imagen|.*image)(?:$|[-/])/i.test(id)
  const supportsTools = raw.supportsFunctionCalling === true || raw.functionCalling === true
  const displayName = typeof raw.displayName === 'string' ? raw.displayName : undefined
  const contextLength = typeof raw.inputTokenLimit === 'number' ? raw.inputTokenLimit : undefined
  return {
    id,
    ...(displayName ? { displayName } : {}),
    ...(contextLength !== undefined ? { contextLength } : {}),
    inputModalities: stringArray(raw.inputModalities),
    outputModalities: modalities,
    capabilities: {
      ...(methods.includes('generateContent') ? { chat: true } : {}),
      ...(supportsTools ? { tools: true } : {}),
      ...(stringArray(raw.inputModalities).some((modality) => /image/i.test(modality))
        ? { vision: true }
        : {}),
      ...(supportsImage ? { imageGeneration: true } : {}),
    },
    providerMetadata: { ...raw },
  }
}

function normalizeAnthropicModel(raw: Record<string, unknown>): AiModel | undefined {
  const id = typeof raw.id === 'string' ? raw.id : ''
  if (!id) return undefined
  const displayName = typeof raw.display_name === 'string' ? raw.display_name : undefined
  const createdAt =
    typeof raw.created_at === 'string' ? Date.parse(raw.created_at) / 1000 : undefined
  return {
    id,
    ...(displayName ? { displayName } : {}),
    ...(createdAt !== undefined && Number.isFinite(createdAt) ? { createdAt } : {}),
    capabilities: { chat: true },
    providerMetadata: { ...raw },
  }
}

function normalizeReplicateModel(raw: Record<string, unknown>): AiModel | undefined {
  const owner = typeof raw.owner === 'string' ? raw.owner : ''
  const name = typeof raw.name === 'string' ? raw.name : ''
  const id = owner && name ? `${owner}/${name}` : typeof raw.id === 'string' ? raw.id : ''
  if (!id) return undefined
  const description = typeof raw.description === 'string' ? raw.description : undefined
  return {
    id,
    ...(description ? { displayName: description } : {}),
    capabilities: { imageGeneration: true },
    providerMetadata: { ...raw },
  }
}

function normalizeFalModel(raw: Record<string, unknown>): AiModel | undefined {
  const id = typeof raw.endpoint_id === 'string' ? raw.endpoint_id : ''
  if (!id) return undefined
  const metadata =
    raw.metadata && typeof raw.metadata === 'object'
      ? (raw.metadata as Record<string, unknown>)
      : {}
  const displayName = typeof metadata.display_name === 'string' ? metadata.display_name : undefined
  const category = typeof metadata.category === 'string' ? metadata.category : ''
  return {
    id,
    ...(displayName ? { displayName } : {}),
    capabilities: {
      ...(/image/i.test(category) ? { imageGeneration: true } : {}),
    },
    providerMetadata: { ...raw },
  }
}

export function fallbackModelCatalog(provider: AiProviderId, now = Date.now()): AiModelCatalog {
  const meta = metaFor(provider)
  const models = (meta?.models ?? []).map((id): AiModel => ({
    id,
    capabilities: {
      chat: true,
      ...(meta?.capabilities?.includes('image-generation') &&
      (meta.imageModels?.includes(id) || /image|imagen|dall|flux|stable/i.test(id))
        ? { imageGeneration: true }
        : {}),
    },
  }))
  return {
    provider,
    ...(meta?.defaultBaseUrl ? { baseUrl: meta.defaultBaseUrl } : {}),
    models,
    source: 'fallback',
    fetchedAt: now,
  }
}

/** Fetch every page of a provider's model catalog and normalize models for task filtering. */
export async function discoverModels(
  provider: AiProviderId,
  config: AiProviderConfig,
  options?: AiModelDiscoveryOptions,
): Promise<AiModelCatalog> {
  const meta = metaFor(provider)
  if (!meta || provider === 'genspark') return fallbackModelCatalog(provider)
  const baseUrl = endpointFor(provider, config, options)
  const path = meta.modelListPath ?? '/models'
  const rawModels = await fetchAllModelPages(
    provider,
    modelListUrl(provider, baseUrl, path),
    authHeaders(provider, config.apiKey),
    options,
  )
  const models = rawModels
    .filter((item): item is Record<string, unknown> => Boolean(item && typeof item === 'object'))
    .map((item) =>
      provider === 'gemini'
        ? normalizeGeminiModel(item)
        : provider === 'anthropic'
          ? normalizeAnthropicModel(item)
          : provider === 'replicate'
            ? normalizeReplicateModel(item)
            : normalizeOpenAiModel(provider, item),
    )
    .filter((item): item is AiModel => Boolean(item))
  const unique = [...new Map(models.map((model) => [model.id, model])).values()]
  return { provider, baseUrl, models: unique, source: 'remote', fetchedAt: Date.now() }
}

/** OpenRouter's image catalog is separate from `/models` and includes output-image
 * capabilities that the text catalog does not guarantee. */
export async function discoverOpenRouterImageModels(
  config: AiProviderConfig,
  options?: AiModelDiscoveryOptions,
): Promise<AiModelCatalog> {
  const provider: AiProviderId = 'openrouter'
  const baseUrl = endpointFor(provider, config, options)
  const body = await fetchJson(
    `${baseUrl}/images/models`,
    authHeaders(provider, config.apiKey),
    options,
  )
  const rawModels = Array.isArray((body as { data?: unknown[] })?.data)
    ? ((body as { data: unknown[] }).data as unknown[])
    : []
  const models = rawModels
    .filter((item): item is Record<string, unknown> => Boolean(item && typeof item === 'object'))
    .flatMap((item): AiModel[] => {
      const model = normalizeOpenAiModel(provider, item)
      if (!model) return []
      return [
        {
          ...model,
          capabilities: { ...model.capabilities, imageGeneration: true },
        },
      ]
    })
  return { provider, baseUrl, models, source: 'remote', fetchedAt: Date.now() }
}

/** Discover every active fal image endpoint, following the platform API's cursor pagination. */
export async function discoverFalImageModels(
  config: AiProviderConfig,
  options?: AiModelDiscoveryOptions,
): Promise<AiModelCatalog> {
  const provider: AiProviderId = 'fal'
  const firstUrl = 'https://api.fal.ai/v1/models?limit=100&category=text-to-image&status=active'
  const models: AiModel[] = []
  const visited = new Set<string>()
  let url: string | undefined = firstUrl
  for (let page = 0; url && page < MAX_MODEL_PAGES; page += 1) {
    if (visited.has(url)) throw new Error('fal model pagination returned a repeated cursor')
    visited.add(url)
    const body = (await fetchJson(url, { Authorization: `Key ${config.apiKey}` }, options)) as {
      models?: unknown[]
      next_cursor?: unknown
      has_more?: unknown
    }
    models.push(
      ...(body.models ?? [])
        .filter((item): item is Record<string, unknown> =>
          Boolean(item && typeof item === 'object'),
        )
        .flatMap((item) => {
          const model = normalizeFalModel(item)
          return model?.capabilities.imageGeneration ? [model] : []
        }),
    )
    if (models.length >= MAX_DISCOVERED_MODELS) break
    if (body.has_more !== true || typeof body.next_cursor !== 'string') break
    const next = new URL(firstUrl)
    next.searchParams.set('cursor', body.next_cursor)
    url = next.toString()
  }
  return {
    provider,
    baseUrl: 'https://api.fal.ai/v1',
    models: [...new Map(models.map((model) => [model.id, model])).values()],
    source: 'remote',
    fetchedAt: Date.now(),
  }
}

/** Discover Runware's merged text-to-image catalog with offset pagination. */
export async function discoverRunwareImageModels(
  config: AiProviderConfig,
  options?: AiModelDiscoveryOptions,
): Promise<AiModelCatalog> {
  const provider: AiProviderId = 'runware'
  const baseUrl = endpointFor(provider, config, options)
  const models: AiModel[] = []
  const search = options?.search?.trim() ?? ''
  const maxPages = Math.max(1, Math.min(options?.maxPages ?? MAX_MODEL_PAGES, MAX_MODEL_PAGES))
  for (let offset = 0, page = 0; page < maxPages; page += 1, offset += 100) {
    const body = (await fetchJson(
      baseUrl,
      { ...authHeaders(provider, config.apiKey), 'Content-Type': 'application/json' },
      options,
      {
        method: 'POST',
        body: JSON.stringify([
          {
            taskType: 'modelSearch',
            taskUUID: globalThis.crypto.randomUUID(),
            search,
            ...(search ? {} : { source: 'merged', capabilities: ['text-to-image'] }),
            offset,
            limit: 100,
          },
        ]),
      },
    )) as { data?: Array<{ results?: unknown[]; totalResults?: number }> }
    const result = body.data?.[0]
    const pageModels = (result?.results ?? [])
      .filter((item): item is Record<string, unknown> => Boolean(item && typeof item === 'object'))
      .flatMap((item): AiModel[] => {
        const id = typeof item.air === 'string' ? item.air : ''
        if (!id) return []
        return [
          {
            id,
            ...(typeof item.name === 'string' ? { displayName: item.name } : {}),
            capabilities: { imageGeneration: true },
            providerMetadata: { ...item },
          },
        ]
      })
    models.push(...pageModels)
    if (!pageModels.length || models.length >= (result?.totalResults ?? 0)) break
    if (models.length >= MAX_DISCOVERED_MODELS) break
  }
  return {
    provider,
    baseUrl,
    models: [...new Map(models.map((model) => [model.id, model])).values()],
    source: 'remote',
    fetchedAt: Date.now(),
  }
}
