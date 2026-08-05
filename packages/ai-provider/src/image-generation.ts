import { AI_PROVIDERS } from './providers'
import { assertValidAiEndpoint } from './endpoint-policy'
import type {
  AiGeneratedImage,
  AiImageGenerationOptions,
  AiImageGenerationRequest,
  AiProviderConfig,
  AiProviderId,
} from './types'

const DEFAULT_TIMEOUT_MS = 60_000

function wait(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(new DOMException('The operation was aborted', 'AbortError'))
      return
    }
    const timer = setTimeout(resolve, ms)
    signal?.addEventListener(
      'abort',
      () => {
        clearTimeout(timer)
        reject(new DOMException('The operation was aborted', 'AbortError'))
      },
      { once: true },
    )
  })
}

function baseUrlFor(
  provider: AiProviderId,
  config: AiProviderConfig,
  options?: AiImageGenerationOptions,
): string {
  const meta = AI_PROVIDERS.find((item) => item.id === provider)
  const baseUrl = options?.baseUrl ?? config.baseUrl ?? meta?.defaultBaseUrl
  if (!baseUrl) throw new Error(`Provider ${provider} requires a Base URL`)
  return assertValidAiEndpoint(baseUrl)
}

function validateRequest(
  request: AiImageGenerationRequest,
): Required<Pick<AiImageGenerationRequest, 'prompt'>> {
  const prompt = request.prompt.trim()
  if (!prompt) throw new Error('Image prompt is required')
  if (prompt.length > 20_000) throw new Error('Image prompt is too long')
  if (
    request.count !== undefined &&
    (!Number.isInteger(request.count) || request.count < 1 || request.count > 10)
  ) {
    throw new Error('Image count must be an integer from 1 to 10')
  }
  return { prompt }
}

function headersFor(provider: AiProviderId, apiKey: string): Record<string, string> {
  if (!apiKey) return { 'Content-Type': 'application/json' }
  if (provider === 'gemini') return { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey }
  if (provider === 'fal')
    return { 'Content-Type': 'application/json', Authorization: `Key ${apiKey}` }
  return { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` }
}

async function postJson(
  url: string,
  provider: AiProviderId,
  apiKey: string,
  body: unknown,
  options?: AiImageGenerationOptions,
): Promise<unknown> {
  const controller = new AbortController()
  const timeout = setTimeout(() => controller.abort(), options?.timeoutMs ?? DEFAULT_TIMEOUT_MS)
  const onAbort = () => controller.abort()
  options?.signal?.addEventListener('abort', onAbort, { once: true })
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: headersFor(provider, apiKey),
      body: JSON.stringify(body),
      signal: controller.signal,
    })
    const text = await response.text()
    let parsed: unknown
    try {
      parsed = text ? (JSON.parse(text) as unknown) : undefined
    } catch {
      parsed = text
    }
    if (!response.ok) {
      const detail = typeof parsed === 'string' ? parsed : JSON.stringify(parsed)
      throw new Error(`Image generation HTTP ${response.status}: ${detail.slice(0, 500)}`)
    }
    return parsed
  } finally {
    clearTimeout(timeout)
    options?.signal?.removeEventListener('abort', onAbort)
  }
}

async function fetchJsonWithTimeout(
  url: string,
  init: RequestInit,
  options?: AiImageGenerationOptions,
): Promise<unknown> {
  const controller = new AbortController()
  const timeout = setTimeout(() => controller.abort(), options?.timeoutMs ?? DEFAULT_TIMEOUT_MS)
  const onAbort = () => controller.abort()
  options?.signal?.addEventListener('abort', onAbort, { once: true })
  try {
    const response = await fetch(url, { ...init, signal: controller.signal })
    const text = await response.text()
    let body: unknown
    try {
      body = text ? (JSON.parse(text) as unknown) : undefined
    } catch {
      body = text
    }
    if (!response.ok) {
      const detail = typeof body === 'string' ? body : JSON.stringify(body)
      throw new Error(`Image generation HTTP ${response.status}: ${detail.slice(0, 500)}`)
    }
    return body
  } finally {
    clearTimeout(timeout)
    options?.signal?.removeEventListener('abort', onAbort)
  }
}

function imageResultsFromUrls(
  provider: AiProviderId,
  model: string,
  urls: string[],
  mimeType = 'image/png',
): AiGeneratedImage[] {
  const images = urls
    .filter((url) => /^https:\/\//i.test(url))
    .map((url) => ({ provider, model, mimeType, url }))
  if (!images.length) throw new Error(`${provider} returned no usable image URLs`)
  return images
}

function dimensions(request: AiImageGenerationRequest): { width: number; height: number } {
  const match = request.size?.match(/^(\d{2,5})x(\d{2,5})$/i)
  if (match) return { width: Number(match[1]), height: Number(match[2]) }
  switch (request.aspectRatio) {
    case '16:9':
      return { width: 1344, height: 768 }
    case '9:16':
      return { width: 768, height: 1344 }
    case '4:3':
      return { width: 1152, height: 896 }
    case '3:4':
      return { width: 896, height: 1152 }
    default:
      return { width: 1024, height: 1024 }
  }
}

async function generateRunwareImage(
  config: AiProviderConfig,
  request: AiImageGenerationRequest,
  options?: AiImageGenerationOptions,
): Promise<AiGeneratedImage[]> {
  const baseUrl = baseUrlFor('runware', config, options)
  const model = request.model ?? config.model
  const { width, height } = dimensions(request)
  const body = await postJson(
    baseUrl,
    'runware',
    config.apiKey,
    [
      {
        taskType: 'imageInference',
        taskUUID: globalThis.crypto.randomUUID(),
        model,
        positivePrompt: request.prompt.trim(),
        width,
        height,
        numberResults: request.count ?? 1,
        outputType: ['URL'],
      },
    ],
    options,
  )
  const data = Array.isArray((body as { data?: unknown[] })?.data)
    ? ((body as { data: unknown[] }).data as unknown[])
    : []
  const urls = data.flatMap((item) => {
    const url = (item as { imageURL?: unknown })?.imageURL
    return typeof url === 'string' ? [url] : []
  })
  return imageResultsFromUrls('runware', model, urls)
}

interface ReplicatePrediction {
  status?: string
  output?: unknown
  error?: unknown
  urls?: { get?: string }
}

function replicateOutputUrls(output: unknown): string[] {
  if (typeof output === 'string') return [output]
  if (Array.isArray(output))
    return output.filter((item): item is string => typeof item === 'string')
  if (output && typeof output === 'object') {
    const record = output as Record<string, unknown>
    const candidate = record.url ?? record.image ?? record.output
    return typeof candidate === 'string' ? [candidate] : []
  }
  return []
}

async function generateReplicateImage(
  config: AiProviderConfig,
  request: AiImageGenerationRequest,
  options?: AiImageGenerationOptions,
): Promise<AiGeneratedImage[]> {
  const baseUrl = baseUrlFor('replicate', config, options)
  const model = request.model ?? config.model
  const input: Record<string, unknown> = {
    prompt: request.prompt.trim(),
    num_outputs: request.count ?? 1,
  }
  if (request.aspectRatio) input.aspect_ratio = request.aspectRatio
  if (request.size) input.output_size = request.size
  let prediction = (await fetchJsonWithTimeout(
    `${baseUrl}/predictions`,
    {
      method: 'POST',
      headers: {
        ...headersFor('replicate', config.apiKey),
        Prefer: 'wait=60',
      },
      body: JSON.stringify({ version: model, input }),
    },
    options,
  )) as ReplicatePrediction
  const deadline = Date.now() + (options?.timeoutMs ?? DEFAULT_TIMEOUT_MS)
  while (prediction.status === 'starting' || prediction.status === 'processing') {
    const pollUrl = prediction.urls?.get
    if (!pollUrl || Date.now() >= deadline) throw new Error('Replicate prediction timed out')
    await wait(900, options?.signal)
    prediction = (await fetchJsonWithTimeout(
      pollUrl,
      { headers: headersFor('replicate', config.apiKey) },
      options,
    )) as ReplicatePrediction
  }
  if (prediction.status === 'failed' || prediction.status === 'canceled') {
    throw new Error(`Replicate prediction ${prediction.status}: ${String(prediction.error ?? '')}`)
  }
  return imageResultsFromUrls('replicate', model, replicateOutputUrls(prediction.output))
}

async function generateFalImage(
  config: AiProviderConfig,
  request: AiImageGenerationRequest,
  options?: AiImageGenerationOptions,
): Promise<AiGeneratedImage[]> {
  const baseUrl = baseUrlFor('fal', config, options)
  const model = request.model ?? config.model
  const submitted = (await fetchJsonWithTimeout(
    `${baseUrl}/${model}`,
    {
      method: 'POST',
      headers: headersFor('fal', config.apiKey),
      body: JSON.stringify({
        prompt: request.prompt.trim(),
        num_images: request.count ?? 1,
        ...(request.imageSize ? { image_size: request.imageSize } : {}),
        ...(request.aspectRatio ? { aspect_ratio: request.aspectRatio } : {}),
      }),
    },
    options,
  )) as { status_url?: string; response_url?: string }
  if (!submitted.status_url || !submitted.response_url) {
    throw new Error('fal did not return queue status URLs')
  }
  const deadline = Date.now() + (options?.timeoutMs ?? DEFAULT_TIMEOUT_MS)
  while (Date.now() < deadline) {
    const status = (await fetchJsonWithTimeout(
      submitted.status_url,
      { headers: headersFor('fal', config.apiKey) },
      options,
    )) as { status?: string; error?: string }
    if (status.error) throw new Error(`fal generation failed: ${status.error}`)
    if (status.status === 'COMPLETED') break
    await wait(750, options?.signal)
  }
  if (Date.now() >= deadline) throw new Error('fal generation timed out')
  const result = (await fetchJsonWithTimeout(
    submitted.response_url,
    { headers: headersFor('fal', config.apiKey) },
    options,
  )) as { images?: Array<{ url?: string; content_type?: string }> }
  return imageResultsFromUrls(
    'fal',
    model,
    (result.images ?? []).flatMap((image) => (image.url ? [image.url] : [])),
    result.images?.[0]?.content_type ?? 'image/png',
  )
}

async function generateStabilityImage(
  config: AiProviderConfig,
  request: AiImageGenerationRequest,
  options?: AiImageGenerationOptions,
): Promise<AiGeneratedImage[]> {
  const baseUrl = baseUrlFor('stability', config, options)
  const model = request.model ?? config.model
  const route =
    model === 'stable-image-ultra'
      ? '/stable-image/generate/ultra'
      : model === 'stable-image-core'
        ? '/stable-image/generate/core'
        : '/stable-image/generate/sd3'
  const form = new FormData()
  form.set('prompt', request.prompt.trim())
  form.set('output_format', 'png')
  if (request.aspectRatio) form.set('aspect_ratio', request.aspectRatio)
  if (route.endsWith('/sd3')) form.set('model', model)
  const result = (await fetchJsonWithTimeout(
    `${baseUrl}${route}`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${config.apiKey}`, Accept: 'application/json' },
      body: form,
    },
    options,
  )) as { image?: string }
  if (!result.image) throw new Error('Stability AI returned no image data')
  return [{ provider: 'stability', model, mimeType: 'image/png', base64: result.image }]
}

function parseImageData(provider: AiProviderId, model: string, body: unknown): AiGeneratedImage[] {
  const data = Array.isArray((body as { data?: unknown[] })?.data)
    ? ((body as { data: unknown[] }).data as unknown[])
    : []
  const images = data
    .filter((item): item is Record<string, unknown> => Boolean(item && typeof item === 'object'))
    .map((item) => {
      const base64 = typeof item.b64_json === 'string' ? item.b64_json : undefined
      const url = typeof item.url === 'string' ? item.url : undefined
      const mimeType = typeof item.media_type === 'string' ? item.media_type : 'image/png'
      const revisedPrompt =
        typeof item.revised_prompt === 'string' ? item.revised_prompt : undefined
      return {
        provider,
        model,
        mimeType,
        ...(base64 ? { base64 } : {}),
        ...(url ? { url } : {}),
        ...(revisedPrompt ? { revisedPrompt } : {}),
      }
    })
    .filter((image) => Boolean(image.base64 || image.url))
  if (!images.length) throw new Error('Image provider returned no usable image data')
  return images
}

async function generateGeminiImage(
  config: AiProviderConfig,
  request: AiImageGenerationRequest,
  options?: AiImageGenerationOptions,
): Promise<AiGeneratedImage[]> {
  const baseUrl = baseUrlFor('gemini', config, options)
  const model = request.model ?? config.model
  const generationConfig: Record<string, unknown> = {
    responseModalities: ['TEXT', 'IMAGE'],
  }
  if (request.aspectRatio) generationConfig.imageConfig = { aspectRatio: request.aspectRatio }
  if (request.imageSize) {
    generationConfig.imageConfig = {
      ...(generationConfig.imageConfig as Record<string, unknown> | undefined),
      imageSize: request.imageSize,
    }
  }
  const body = await postJson(
    `${baseUrl}/models/${encodeURIComponent(model)}:generateContent`,
    'gemini',
    config.apiKey,
    { contents: [{ role: 'user', parts: [{ text: request.prompt.trim() }] }], generationConfig },
    options,
  )
  const candidates = Array.isArray((body as { candidates?: unknown[] })?.candidates)
    ? ((body as { candidates: unknown[] }).candidates as unknown[])
    : []
  const images: AiGeneratedImage[] = []
  for (const candidate of candidates) {
    const parts = Array.isArray((candidate as { content?: { parts?: unknown[] } })?.content?.parts)
      ? ((candidate as { content: { parts: unknown[] } }).content.parts as unknown[])
      : []
    for (const part of parts) {
      const inline = ((part as { inlineData?: unknown; inline_data?: unknown }).inlineData ??
        (part as { inline_data?: unknown }).inline_data) as
        { data?: unknown; mimeType?: unknown; mime_type?: unknown } | undefined
      if (!inline || typeof inline.data !== 'string') continue
      images.push({
        provider: 'gemini',
        model,
        mimeType:
          typeof inline.mimeType === 'string'
            ? inline.mimeType
            : typeof inline.mime_type === 'string'
              ? inline.mime_type
              : 'image/png',
        base64: inline.data,
      })
    }
  }
  if (!images.length) throw new Error('Gemini returned no inline image data')
  return images
}

/** Generate images through native Gemini, OpenAI-compatible image APIs, or OpenRouter's
 * dedicated `/api/v1/images` endpoint. URL responses are returned for the caller to download
 * through its own SSRF-guarded main-process fetcher. */
export async function generateImageForProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  request: AiImageGenerationRequest,
  options?: AiImageGenerationOptions,
): Promise<AiGeneratedImage[]> {
  const normalized = validateRequest(request)
  if (provider === 'gemini')
    return generateGeminiImage(config, { ...request, ...normalized }, options)
  const meta = AI_PROVIDERS.find((item) => item.id === provider)
  if (meta?.imageProtocol === 'runware')
    return generateRunwareImage(config, { ...request, ...normalized }, options)
  if (meta?.imageProtocol === 'replicate')
    return generateReplicateImage(config, { ...request, ...normalized }, options)
  if (meta?.imageProtocol === 'fal')
    return generateFalImage(config, { ...request, ...normalized }, options)
  if (meta?.imageProtocol === 'stability')
    return generateStabilityImage(config, { ...request, ...normalized }, options)
  if (meta?.imageProtocol !== 'openai-images') {
    throw new Error(`Provider ${provider} does not advertise image generation support`)
  }
  const baseUrl = baseUrlFor(provider, config, options)
  const model = request.model ?? config.model
  const body: Record<string, unknown> = { model, prompt: normalized.prompt }
  if (request.count !== undefined) body.n = request.count
  if (request.size) body.size = request.size
  if (provider !== 'openrouter') body.response_format = 'b64_json'
  if (request.aspectRatio) body.aspect_ratio = request.aspectRatio
  if (request.imageSize) body.image_size = request.imageSize
  // OpenRouter intentionally uses its unified Image API rather than the older
  // OpenAI-compatible `/images/generations` route.
  const path = provider === 'openrouter' ? '/images' : '/images/generations'
  const response = await postJson(`${baseUrl}${path}`, provider, config.apiKey, body, options)
  return parseImageData(provider, model, response)
}
