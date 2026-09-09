/**
 * Wire implementations for the BYOK media providers. Image generation speaks
 * four shapes — OpenAI Images (OpenAI, Ark/Seedream, Zhipu/CogView, xAI,
 * custom endpoints; the vendors differ in size vocabulary and edit support),
 * Gemini native output / Imagen predict, DashScope multimodal-generation
 * (Qwen-Image) and MiniMax image_generation. Understanding speaks two: OpenAI
 * chat completions (video via a video_url part where the vendor accepts one)
 * and Gemini generateContent (inline or Files API). Genspark is not here — its
 * tools go through the gsk CLI in @genoffice/ai-search.
 */

import { aiFetch } from './fetch'
import { httpBodyDetail } from './http-error'
import {
  DASHSCOPE_BASE_URL,
  GEMINI_MEDIA_BASE_URL,
  MINIMAX_BASE_URL,
  OPENAI_IMAGES_BASE_URL,
  getMediaProviderMeta,
} from './media'
import type { AiMediaProviderConfig, AiMediaProviderId, AiMediaProviderMeta } from './types'

export interface MediaBlob {
  bytes: Uint8Array
  mime: string
  name?: string
}

export interface GenerateImageInput {
  prompt: string
  /** 1:1 | 4:3 | 16:9 | 9:16 | 3:4 | 2:3 | 3:2 | auto (vendors get their nearest supported value) */
  aspectRatio?: string | undefined
  /** images to edit / draw from */
  references?: MediaBlob[] | undefined
}

export interface AnalyzeMediaInput {
  media: MediaBlob[]
  requirements: string
}

export type ByokMediaProviderId = Exclude<AiMediaProviderId, 'genspark'>

/** image generation can take minutes on the large models */
const GENERATE_TIMEOUT_MS = 600_000
const ANALYZE_TIMEOUT_MS = 300_000
const TEST_TIMEOUT_MS = 20_000
/** Gemini caps the whole inline request at 20 MB; bigger media goes through the Files API */
const GEMINI_INLINE_LIMIT_BYTES = 18 * 1024 * 1024
const GEMINI_FILE_POLL_MS = 2_000
const GEMINI_FILE_READY_TIMEOUT_MS = 180_000

function trimSlash(url: string): string {
  return url.replace(/\/+$/, '')
}

function metaOf(provider: ByokMediaProviderId): AiMediaProviderMeta {
  const meta = getMediaProviderMeta(provider)
  if (!meta) throw new Error(`Unknown media provider: ${provider}`)
  return meta
}

/** base URL of the OpenAI-shaped endpoints (images + chat) for a provider */
function openAiBase(provider: ByokMediaProviderId, config: AiMediaProviderConfig): string {
  if (provider === 'qwen') return `${dashscopeRoot(config)}/compatible-mode/v1`
  const meta = metaOf(provider)
  return trimSlash(config.baseUrl || meta.defaultBaseUrl || OPENAI_IMAGES_BASE_URL)
}

/** DashScope root; a pasted compatible-mode or api/v1 URL is reduced to it */
function dashscopeRoot(config: AiMediaProviderConfig): string {
  return trimSlash(config.baseUrl || DASHSCOPE_BASE_URL)
    .replace(/\/compatible-mode\/v1$/, '')
    .replace(/\/api\/v1$/, '')
}

function geminiBase(config: AiMediaProviderConfig): string {
  return trimSlash(config.baseUrl || GEMINI_MEDIA_BASE_URL)
}

function bearer(config: AiMediaProviderConfig): Record<string, string> {
  return config.apiKey ? { Authorization: `Bearer ${config.apiKey}` } : {}
}

export function bytesToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64')
  let binary = ''
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  }
  return btoa(binary)
}

export function base64ToBytes(b64: string): Uint8Array {
  if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(b64, 'base64'))
  const binary = atob(b64)
  const out = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i)
  return out
}

function dataUrl(blob: MediaBlob): string {
  return `data:${blob.mime};base64,${bytesToBase64(blob.bytes)}`
}

/** Image bytes declare their own format; vendors' declared mime is a fallback only */
export function sniffImageMime(bytes: Uint8Array, fallback = 'image/png'): string {
  if (bytes.length > 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e) {
    return 'image/png'
  }
  if (bytes.length > 3 && bytes[0] === 0xff && bytes[1] === 0xd8) return 'image/jpeg'
  if (bytes.length > 6 && bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
    return 'image/gif'
  }
  if (
    bytes.length > 12 &&
    bytes[8] === 0x57 &&
    bytes[9] === 0x45 &&
    bytes[10] === 0x42 &&
    bytes[11] === 0x50
  ) {
    return 'image/webp'
  }
  return fallback
}

/** Blob/fetch bodies want an ArrayBuffer-backed view; a subarray of a shared buffer is copied */
function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
}

async function failFrom(label: string, resp: Response): Promise<never> {
  const body = await resp.text().catch(() => '')
  throw new Error(`${label} ${resp.status}: ${httpBodyDetail(body)}`)
}

function asRecord(v: unknown): Record<string, unknown> {
  return v && typeof v === 'object' ? (v as Record<string, unknown>) : {}
}

function withTimeout(signal: AbortSignal | undefined, ms: number): AbortSignal {
  const timeout = AbortSignal.timeout(ms)
  return signal ? AbortSignal.any([signal, timeout]) : timeout
}

/** vendors that return a URL instead of bytes: download it (the links are short-lived, so right away) */
async function downloadImage(label: string, url: string, signal: AbortSignal): Promise<MediaBlob> {
  const img = await aiFetch(url, { signal })
  if (!img.ok) return failFrom(`${label} download`, img)
  const bytes = new Uint8Array(await img.arrayBuffer())
  return { bytes, mime: sniffImageMime(bytes, img.headers.get('content-type') ?? 'image/png') }
}

function fromBase64(b64: string, declaredMime?: unknown): MediaBlob {
  const bytes = base64ToBytes(b64)
  return {
    bytes,
    mime: sniffImageMime(bytes, typeof declaredMime === 'string' ? declaredMime : 'image/png'),
  }
}

// ── OpenAI Images API family (openai, doubao/Ark, glm/Zhipu, xai, custom) ──

const OPENAI_SIZES: Record<string, string> = {
  '1:1': '1024x1024',
  '16:9': '1536x1024',
  '3:2': '1536x1024',
  '4:3': '1536x1024',
  '9:16': '1024x1536',
  '2:3': '1024x1536',
  '3:4': '1024x1536',
}
/** CogView accepts a fixed grid of WxH values */
const GLM_SIZES: Record<string, string> = {
  '1:1': '1024x1024',
  '16:9': '1344x768',
  '3:2': '1344x768',
  '4:3': '1152x864',
  '9:16': '768x1344',
  '2:3': '768x1344',
  '3:4': '864x1152',
}
/** Seedream takes the aspect ratio itself as `size` (resolution defaults to 2K) */
const ARK_RATIOS = new Set(['1:1', '3:4', '4:3', '9:16', '16:9', '2:3', '3:2', '21:9'])

interface OpenAiImagesStyle {
  size(aspectRatio: string | undefined): string | undefined
  /** vendor-specific fields merged into the generations body */
  bodyExtras: Record<string, unknown>
  /** 'multipart' = POST /images/edits with files; 'inline' = `image` data URLs in the generations body; 'none' = unsupported */
  edits: 'multipart' | 'inline' | 'none'
}

function openAiImagesStyle(provider: ByokMediaProviderId): OpenAiImagesStyle {
  switch (provider) {
    case 'doubao':
      return {
        size: (r) => (r && ARK_RATIOS.has(r) ? r : '2K'),
        bodyExtras: { response_format: 'b64_json', watermark: false },
        edits: 'inline',
      }
    case 'glm':
      return { size: (r) => (r ? GLM_SIZES[r] : undefined), bodyExtras: {}, edits: 'none' }
    case 'xai':
      // xAI rejects size/quality; bytes only on request
      return { size: () => undefined, bodyExtras: { response_format: 'b64_json' }, edits: 'none' }
    default:
      return { size: (r) => (r ? OPENAI_SIZES[r] : undefined), bodyExtras: {}, edits: 'multipart' }
  }
}

async function openAiImageResult(
  label: string,
  resp: Response,
  signal: AbortSignal,
): Promise<MediaBlob> {
  if (!resp.ok) return failFrom(label, resp)
  const json = asRecord(await resp.json())
  const first = asRecord((json.data as unknown[] | undefined)?.[0])
  if (typeof first.b64_json === 'string' && first.b64_json) return fromBase64(first.b64_json)
  if (typeof first.url === 'string' && first.url) return downloadImage(label, first.url, signal)
  throw new Error(`${label} returned no image: ${JSON.stringify(json).slice(0, 200)}`)
}

async function generateImageOpenAi(
  provider: ByokMediaProviderId,
  config: AiMediaProviderConfig,
  model: string,
  input: GenerateImageInput,
  signal: AbortSignal,
): Promise<MediaBlob> {
  const base = openAiBase(provider, config)
  const style = openAiImagesStyle(provider)
  const size = style.size(input.aspectRatio)
  const refs = input.references ?? []
  if (refs.length && style.edits === 'none') {
    throw new Error(
      `${metaOf(provider).label} cannot edit or reference images here; generate from the prompt alone or pick another image provider.`,
    )
  }
  if (refs.length === 0 || style.edits === 'inline') {
    const resp = await aiFetch(`${base}/images/generations`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...bearer(config) },
      body: JSON.stringify({
        model,
        prompt: input.prompt,
        n: 1,
        ...(size ? { size } : {}),
        ...style.bodyExtras,
        ...(refs.length ? { image: refs.map(dataUrl) } : {}),
      }),
      signal,
    })
    return openAiImageResult('Image generation failed:', resp, signal)
  }
  // edits take the source images as multipart files (image[] for several)
  const form = new FormData()
  form.set('model', model)
  form.set('prompt', input.prompt)
  if (size) form.set('size', size)
  refs.forEach((ref, i) => {
    const ext = ref.mime.split('/')[1]?.replace('jpeg', 'jpg') ?? 'png'
    form.append(
      refs.length > 1 ? 'image[]' : 'image',
      new Blob([toArrayBuffer(ref.bytes)], { type: ref.mime }),
      ref.name ?? `ref-${i}.${ext}`,
    )
  })
  const resp = await aiFetch(`${base}/images/edits`, {
    method: 'POST',
    headers: bearer(config),
    body: form,
    signal,
  })
  return openAiImageResult('Image edit failed:', resp, signal)
}

// ── DashScope multimodal-generation (Qwen-Image) ───────────────────

const DASHSCOPE_SIZES: Record<string, string> = {
  '1:1': '1328*1328',
  '16:9': '1664*928',
  '3:2': '1664*928',
  '4:3': '1472*1104',
  '9:16': '928*1664',
  '2:3': '928*1664',
  '3:4': '1104*1472',
}

async function generateImageDashscope(
  config: AiMediaProviderConfig,
  model: string,
  input: GenerateImageInput,
  signal: AbortSignal,
): Promise<MediaBlob> {
  if (input.references?.length) {
    throw new Error(
      'Qwen-Image here generates from the prompt alone; pick another image provider for edits.',
    )
  }
  const size = input.aspectRatio ? DASHSCOPE_SIZES[input.aspectRatio] : undefined
  const resp = await aiFetch(
    `${dashscopeRoot(config)}/api/v1/services/aigc/multimodal-generation/generation`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...bearer(config) },
      body: JSON.stringify({
        model,
        input: { messages: [{ role: 'user', content: [{ text: input.prompt }] }] },
        parameters: { n: 1, watermark: false, ...(size ? { size } : {}) },
      }),
      signal,
    },
  )
  if (!resp.ok) return failFrom('Image generation failed:', resp)
  const json = asRecord(await resp.json())
  if (typeof json.code === 'string' && json.code) {
    throw new Error(`Image generation failed: ${json.code} ${String(json.message ?? '')}`)
  }
  const choice = asRecord((asRecord(json.output).choices as unknown[] | undefined)?.[0])
  const content = asRecord(choice.message).content
  const image = Array.isArray(content)
    ? content.map(asRecord).find((c) => typeof c.image === 'string')?.image
    : undefined
  if (typeof image !== 'string') {
    throw new Error(`Image generation returned no image: ${JSON.stringify(json).slice(0, 200)}`)
  }
  return downloadImage('Image generation failed:', image, signal)
}

// ── MiniMax image_generation ───────────────────────────────────────

const MINIMAX_RATIOS = new Set(['1:1', '16:9', '4:3', '3:2', '2:3', '3:4', '9:16', '21:9'])

async function generateImageMinimax(
  config: AiMediaProviderConfig,
  model: string,
  input: GenerateImageInput,
  signal: AbortSignal,
): Promise<MediaBlob> {
  if (input.references?.length) {
    throw new Error(
      'MiniMax image-01 here generates from the prompt alone; pick another image provider for edits.',
    )
  }
  const base = trimSlash(config.baseUrl || MINIMAX_BASE_URL)
  const ratio =
    input.aspectRatio && MINIMAX_RATIOS.has(input.aspectRatio) ? input.aspectRatio : undefined
  const resp = await aiFetch(`${base}/image_generation`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...bearer(config) },
    body: JSON.stringify({
      model,
      prompt: input.prompt,
      n: 1,
      response_format: 'base64',
      ...(ratio ? { aspect_ratio: ratio } : {}),
    }),
    signal,
  })
  if (!resp.ok) return failFrom('Image generation failed:', resp)
  const json = asRecord(await resp.json())
  const status = asRecord(json.base_resp)
  if (typeof status.status_code === 'number' && status.status_code !== 0) {
    throw new Error(
      `Image generation failed: ${status.status_code} ${String(status.status_msg ?? '')}`,
    )
  }
  const data = asRecord(json.data)
  const b64 = (data.image_base64 as unknown[] | undefined)?.[0]
  if (typeof b64 === 'string' && b64) return fromBase64(b64)
  const url = (data.image_urls as unknown[] | undefined)?.[0]
  if (typeof url === 'string' && url) return downloadImage('Image generation failed:', url, signal)
  throw new Error(`Image generation returned no image: ${JSON.stringify(json).slice(0, 200)}`)
}

// ── OpenAI-compatible chat understanding ───────────────────────────

function openAiContentText(content: unknown): string {
  if (typeof content === 'string') return content
  if (Array.isArray(content)) {
    return content
      .map((part) => {
        const p = asRecord(part)
        return typeof p.text === 'string' ? p.text : ''
      })
      .join('')
  }
  return ''
}

async function analyzeMediaOpenAi(
  provider: ByokMediaProviderId,
  config: AiMediaProviderConfig,
  model: string,
  input: AnalyzeMediaInput,
  signal: AbortSignal,
): Promise<string> {
  const meta = metaOf(provider)
  const parts: unknown[] = [{ type: 'text', text: input.requirements }]
  for (const m of input.media) {
    if (m.mime.startsWith('image/')) {
      parts.push({ type: 'image_url', image_url: { url: dataUrl(m) } })
    } else if (m.mime.startsWith('video/') && meta.videoAnalysis) {
      parts.push({ type: 'video_url', video_url: { url: dataUrl(m) } })
    } else {
      throw new Error(
        `${meta.label} cannot analyze ${m.name ?? m.mime} (${m.mime}) here; ${
          meta.videoAnalysis ? 'audio' : 'video and audio'
        } analysis needs Gemini or Genspark.`,
      )
    }
  }
  const resp = await aiFetch(`${openAiBase(provider, config)}/chat/completions`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...bearer(config) },
    body: JSON.stringify({ model, messages: [{ role: 'user', content: parts }] }),
    signal,
  })
  if (!resp.ok) return failFrom('Media analysis failed:', resp)
  const json = asRecord(await resp.json())
  const choice = asRecord((json.choices as unknown[] | undefined)?.[0])
  const text = openAiContentText(asRecord(choice.message).content).trim()
  if (!text) throw new Error('Media analysis returned an empty answer')
  return text
}

// ── Gemini (native image output, Imagen predict, multimodal understanding) ──

const GEMINI_ASPECT_RATIOS = new Set([
  '1:1',
  '2:3',
  '3:2',
  '3:4',
  '4:3',
  '4:5',
  '5:4',
  '9:16',
  '16:9',
  '21:9',
])

function geminiHeaders(config: AiMediaProviderConfig): Record<string, string> {
  return { 'Content-Type': 'application/json', 'x-goog-api-key': config.apiKey }
}

function geminiParts(candidate: unknown): Record<string, unknown>[] {
  const content = asRecord(asRecord(candidate).content)
  return Array.isArray(content.parts) ? content.parts.map(asRecord) : []
}

function geminiBlockedReason(json: Record<string, unknown>): string {
  const feedback = asRecord(json.promptFeedback)
  const candidate = asRecord((json.candidates as unknown[] | undefined)?.[0])
  const reason = feedback.blockReason ?? candidate.finishReason
  return typeof reason === 'string' && reason !== 'STOP' ? ` (${reason})` : ''
}

async function generateImageGemini(
  config: AiMediaProviderConfig,
  model: string,
  input: GenerateImageInput,
  signal: AbortSignal,
): Promise<MediaBlob> {
  const base = geminiBase(config)
  const aspectRatio =
    input.aspectRatio && GEMINI_ASPECT_RATIOS.has(input.aspectRatio) ? input.aspectRatio : undefined
  if (model.startsWith('imagen-')) {
    const resp = await aiFetch(`${base}/models/${model}:predict`, {
      method: 'POST',
      headers: geminiHeaders(config),
      body: JSON.stringify({
        instances: [{ prompt: input.prompt }],
        parameters: { sampleCount: 1, ...(aspectRatio ? { aspectRatio } : {}) },
      }),
      signal,
    })
    if (!resp.ok) return failFrom('Image generation failed:', resp)
    const json = asRecord(await resp.json())
    const first = asRecord((json.predictions as unknown[] | undefined)?.[0])
    if (typeof first.bytesBase64Encoded !== 'string') {
      throw new Error(`Image generation returned no image: ${JSON.stringify(json).slice(0, 200)}`)
    }
    return fromBase64(first.bytesBase64Encoded, first.mimeType)
  }
  const resp = await aiFetch(`${base}/models/${model}:generateContent`, {
    method: 'POST',
    headers: geminiHeaders(config),
    body: JSON.stringify({
      contents: [
        {
          role: 'user',
          parts: [
            { text: input.prompt },
            ...(input.references ?? []).map((ref) => ({
              inline_data: { mime_type: ref.mime, data: bytesToBase64(ref.bytes) },
            })),
          ],
        },
      ],
      generationConfig: {
        responseModalities: ['TEXT', 'IMAGE'],
        ...(aspectRatio ? { imageConfig: { aspectRatio } } : {}),
      },
    }),
    signal,
  })
  if (!resp.ok) return failFrom('Image generation failed:', resp)
  const json = asRecord(await resp.json())
  const parts = geminiParts((json.candidates as unknown[] | undefined)?.[0])
  const image = parts
    .map((p) => asRecord(p.inlineData ?? p.inline_data))
    .find((d) => typeof d.data === 'string')
  if (!image) {
    const text = parts
      .map((p) => (typeof p.text === 'string' ? p.text : ''))
      .join(' ')
      .trim()
    throw new Error(
      `Image generation returned no image${geminiBlockedReason(json)}${text ? `: ${text.slice(0, 200)}` : ''}`,
    )
  }
  return fromBase64(image.data as string, image.mimeType ?? image.mime_type)
}

/** Files API resumable upload (start → upload+finalize → wait for ACTIVE); returns the file_data part */
async function geminiUploadFile(
  config: AiMediaProviderConfig,
  blob: MediaBlob,
  signal: AbortSignal,
): Promise<{ file_data: { mime_type: string; file_uri: string } }> {
  const base = geminiBase(config)
  const root = base.replace(/\/v1(beta)?$/, '')
  const start = await aiFetch(`${root}/upload/v1beta/files`, {
    method: 'POST',
    headers: {
      ...geminiHeaders(config),
      'X-Goog-Upload-Protocol': 'resumable',
      'X-Goog-Upload-Command': 'start',
      'X-Goog-Upload-Header-Content-Length': String(blob.bytes.byteLength),
      'X-Goog-Upload-Header-Content-Type': blob.mime,
    },
    body: JSON.stringify({ file: { display_name: blob.name ?? 'media' } }),
    signal,
  })
  if (!start.ok) return failFrom('Media upload failed:', start)
  const uploadUrl = start.headers.get('x-goog-upload-url')
  if (!uploadUrl) throw new Error('Media upload failed: no upload URL returned')
  const upload = await aiFetch(uploadUrl, {
    method: 'POST',
    headers: {
      'x-goog-api-key': config.apiKey,
      'Content-Length': String(blob.bytes.byteLength),
      'X-Goog-Upload-Offset': '0',
      'X-Goog-Upload-Command': 'upload, finalize',
    },
    body: toArrayBuffer(blob.bytes),
    signal,
  })
  if (!upload.ok) return failFrom('Media upload failed:', upload)
  let file = asRecord(asRecord(await upload.json()).file)
  const deadline = Date.now() + GEMINI_FILE_READY_TIMEOUT_MS
  // videos are transcoded server-side before they can be referenced
  while (file.state === 'PROCESSING') {
    if (Date.now() > deadline) {
      throw new Error('Media upload timed out while the file was processing')
    }
    await new Promise((r) => setTimeout(r, GEMINI_FILE_POLL_MS))
    const poll = await aiFetch(`${base}/${String(file.name)}`, {
      headers: { 'x-goog-api-key': config.apiKey },
      signal,
    })
    if (!poll.ok) return failFrom('Media upload failed:', poll)
    file = asRecord(await poll.json())
  }
  if (file.state !== 'ACTIVE' || typeof file.uri !== 'string') {
    throw new Error(`Media upload failed: file state ${String(file.state ?? 'unknown')}`)
  }
  return { file_data: { mime_type: String(file.mimeType ?? blob.mime), file_uri: file.uri } }
}

async function analyzeMediaGemini(
  config: AiMediaProviderConfig,
  model: string,
  input: AnalyzeMediaInput,
  signal: AbortSignal,
): Promise<string> {
  const total = input.media.reduce((n, m) => n + m.bytes.byteLength, 0)
  const inline = total <= GEMINI_INLINE_LIMIT_BYTES
  const mediaParts: unknown[] = []
  for (const blob of input.media) {
    mediaParts.push(
      inline
        ? { inline_data: { mime_type: blob.mime, data: bytesToBase64(blob.bytes) } }
        : await geminiUploadFile(config, blob, signal),
    )
  }
  const resp = await aiFetch(`${geminiBase(config)}/models/${model}:generateContent`, {
    method: 'POST',
    headers: geminiHeaders(config),
    body: JSON.stringify({
      contents: [{ role: 'user', parts: [...mediaParts, { text: input.requirements }] }],
    }),
    signal,
  })
  if (!resp.ok) return failFrom('Media analysis failed:', resp)
  const json = asRecord(await resp.json())
  const text = geminiParts((json.candidates as unknown[] | undefined)?.[0])
    .map((p) => (typeof p.text === 'string' ? p.text : ''))
    .join('')
    .trim()
  if (!text) throw new Error(`Media analysis returned an empty answer${geminiBlockedReason(json)}`)
  return text
}

// ── dispatch ────────────────────────────────────────────────────────

function modelOf(
  meta: AiMediaProviderMeta,
  config: AiMediaProviderConfig,
  field: 'imageModel' | 'analysisModel',
): string {
  const model =
    config[field] || (field === 'imageModel' ? meta.defaultImageModel : meta.defaultAnalysisModel)
  if (!model) {
    throw new Error(
      `No ${field === 'imageModel' ? 'image' : 'analysis'} model configured for the ${meta.label} media provider`,
    )
  }
  return model
}

function requireBaseUrl(meta: AiMediaProviderMeta, config: AiMediaProviderConfig): void {
  if (meta.needsBaseUrl && !config.baseUrl) {
    throw new Error(`The ${meta.label} media provider requires a Base URL`)
  }
}

export async function generateImageWithProvider(
  provider: ByokMediaProviderId,
  config: AiMediaProviderConfig,
  input: GenerateImageInput,
  signal?: AbortSignal,
): Promise<MediaBlob> {
  const meta = metaOf(provider)
  requireBaseUrl(meta, config)
  if (!meta.imageProtocol) throw new Error(`${meta.label} does not generate images`)
  const model = modelOf(meta, config, 'imageModel')
  const guard = withTimeout(signal, GENERATE_TIMEOUT_MS)
  switch (meta.imageProtocol) {
    case 'gemini':
      return generateImageGemini(config, model, input, guard)
    case 'dashscope':
      return generateImageDashscope(config, model, input, guard)
    case 'minimax':
      return generateImageMinimax(config, model, input, guard)
    case 'openai-images':
      return generateImageOpenAi(provider, config, model, input, guard)
  }
}

export async function analyzeMediaWithProvider(
  provider: ByokMediaProviderId,
  config: AiMediaProviderConfig,
  input: AnalyzeMediaInput,
  signal?: AbortSignal,
): Promise<string> {
  const meta = metaOf(provider)
  requireBaseUrl(meta, config)
  if (!meta.analysisProtocol) throw new Error(`${meta.label} does not analyze media`)
  const model = modelOf(meta, config, 'analysisModel')
  const guard = withTimeout(signal, ANALYZE_TIMEOUT_MS)
  return meta.analysisProtocol === 'gemini'
    ? analyzeMediaGemini(config, model, input, guard)
    : analyzeMediaOpenAi(provider, config, model, input, guard)
}

/**
 * Cheap credential check — the settings-UI connection test; no image is
 * billed. A model listing is the closest thing every vendor has; vendors
 * without one answer 404/405 to a valid key, so only 401/403 count as failure.
 */
export async function testMediaProvider(
  provider: ByokMediaProviderId,
  config: AiMediaProviderConfig,
  signal?: AbortSignal,
): Promise<{ ok: boolean; error?: string }> {
  try {
    const meta = metaOf(provider)
    requireBaseUrl(meta, config)
    const guard = withTimeout(signal, TEST_TIMEOUT_MS)
    const resp =
      provider === 'gemini'
        ? await aiFetch(`${geminiBase(config)}/models?pageSize=1`, {
            headers: { 'x-goog-api-key': config.apiKey },
            signal: guard,
          })
        : await aiFetch(`${openAiBase(provider, config)}/models`, {
            headers: bearer(config),
            signal: guard,
          })
    if (resp.ok) return { ok: true }
    const body = await resp.text().catch(() => '')
    if (resp.status === 401 || resp.status === 403 || provider === 'gemini') {
      return { ok: false, error: `HTTP ${resp.status}: ${httpBodyDetail(body)}` }
    }
    return { ok: true }
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : String(e) }
  }
}
