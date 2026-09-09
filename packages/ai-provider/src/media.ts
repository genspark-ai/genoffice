import type {
  AiMediaProviderConfig,
  AiMediaProviderId,
  AiMediaProviderMeta,
  AiMediaSettings,
  AiSettings,
} from './types'

export const OPENAI_IMAGES_BASE_URL = 'https://api.openai.com/v1'
export const GEMINI_MEDIA_BASE_URL = 'https://generativelanguage.googleapis.com/v1beta'
export const ARK_BASE_URL = 'https://ark.cn-beijing.volces.com/api/v3'
export const ZHIPU_BASE_URL = 'https://open.bigmodel.cn/api/paas/v4'
export const XAI_BASE_URL = 'https://api.x.ai/v1'
/** DashScope root: images ride /api/v1/services/aigc/..., understanding rides /compatible-mode/v1 */
export const DASHSCOPE_BASE_URL = 'https://dashscope.aliyuncs.com'
export const MINIMAX_BASE_URL = 'https://api.minimax.io/v1'

// Model ids verified against vendor docs 2026-09; keep chat-capable analysis
// models in step with the chat catalog in providers.ts.
export const AI_MEDIA_PROVIDERS: AiMediaProviderMeta[] = [
  {
    id: 'genspark',
    label: 'Genspark',
    description: 'Image generation, media analysis and search through your Genspark sign-in',
    keyPlaceholder: 'Not required - sign in to Genspark',
    defaultBaseUrl: '',
    imageProtocol: 'openai-images',
    imageModels: [],
    defaultImageModel: '',
    analysisProtocol: 'openai-chat',
    analysisModels: [],
    defaultAnalysisModel: '',
    videoAnalysis: true,
  },
  {
    id: 'openai',
    label: 'OpenAI',
    description: 'GPT Image for generation and editing; GPT chat models for image analysis',
    keyPlaceholder: 'sk-...',
    defaultBaseUrl: OPENAI_IMAGES_BASE_URL,
    imageProtocol: 'openai-images',
    imageModels: ['gpt-image-2', 'gpt-image-1.5', 'gpt-image-1', 'gpt-image-1-mini'],
    defaultImageModel: 'gpt-image-2',
    analysisProtocol: 'openai-chat',
    analysisModels: ['gpt-5.6-sol', 'gpt-5.6-terra', 'gpt-5.6-luna', 'gpt-5.5', 'gpt-5.4-mini'],
    defaultAnalysisModel: 'gpt-5.6-luna',
    videoAnalysis: false,
  },
  {
    id: 'gemini',
    label: 'Gemini',
    description:
      'Gemini native image output (Nano Banana) and Imagen; Gemini reads images, video and audio',
    keyPlaceholder: 'AIza...',
    defaultBaseUrl: GEMINI_MEDIA_BASE_URL,
    imageProtocol: 'gemini',
    imageModels: [
      'gemini-3.1-flash-image',
      'gemini-3-pro-image',
      'gemini-3.1-flash-lite-image',
      'gemini-2.5-flash-image',
    ],
    defaultImageModel: 'gemini-3.1-flash-image',
    analysisProtocol: 'gemini',
    analysisModels: ['gemini-3.7-flash', 'gemini-3.1-pro-preview', 'gemini-3.6-flash'],
    defaultAnalysisModel: 'gemini-3.7-flash',
    videoAnalysis: true,
  },
  {
    id: 'doubao',
    label: 'Doubao (Volcengine Ark)',
    description: 'Seedream image generation and editing; Doubao Seed reads images and video',
    keyPlaceholder: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
    defaultBaseUrl: ARK_BASE_URL,
    imageProtocol: 'openai-images',
    imageModels: [
      'doubao-seedream-5-0-260128',
      'doubao-seedream-4-5-251128',
      'doubao-seedream-4-0-250828',
    ],
    defaultImageModel: 'doubao-seedream-5-0-260128',
    analysisProtocol: 'openai-chat',
    analysisModels: ['doubao-seed-2-1-pro-260628', 'doubao-seed-2-1-turbo-260628'],
    defaultAnalysisModel: 'doubao-seed-2-1-pro-260628',
    videoAnalysis: true,
  },
  {
    id: 'glm',
    label: 'GLM (Zhipu)',
    description: 'CogView image generation; GLM-V reads images and video',
    keyPlaceholder: 'xxxxxxxx.xxxxxxxx',
    defaultBaseUrl: ZHIPU_BASE_URL,
    imageProtocol: 'openai-images',
    imageModels: ['cogview-4-250304', 'cogview-3-flash'],
    defaultImageModel: 'cogview-4-250304',
    analysisProtocol: 'openai-chat',
    analysisModels: ['glm-4.6v', 'glm-4.6v-flash'],
    defaultAnalysisModel: 'glm-4.6v',
    videoAnalysis: true,
  },
  {
    id: 'xai',
    label: 'Grok (xAI)',
    description: 'Grok Imagine image generation; Grok reads images',
    keyPlaceholder: 'xai-...',
    defaultBaseUrl: XAI_BASE_URL,
    imageProtocol: 'openai-images',
    imageModels: ['grok-imagine-image-2.0', 'grok-2-image-1212'],
    defaultImageModel: 'grok-imagine-image-2.0',
    analysisProtocol: 'openai-chat',
    analysisModels: ['grok-4.6', 'grok-4.5'],
    defaultAnalysisModel: 'grok-4.6',
    videoAnalysis: false,
  },
  {
    id: 'qwen',
    label: 'Qwen (Alibaba Cloud)',
    description: 'Qwen-Image generation; Qwen-VL reads images and video (DashScope)',
    keyPlaceholder: 'sk-...',
    defaultBaseUrl: DASHSCOPE_BASE_URL,
    imageProtocol: 'dashscope',
    imageModels: ['qwen-image-plus', 'qwen-image'],
    defaultImageModel: 'qwen-image-plus',
    analysisProtocol: 'openai-chat',
    analysisModels: ['qwen3-vl-plus', 'qwen3-vl-flash', 'qwen3.8-max', 'qwen3.7-plus'],
    defaultAnalysisModel: 'qwen3-vl-plus',
    videoAnalysis: true,
  },
  {
    id: 'minimax',
    label: 'MiniMax',
    description:
      'image-01 image generation (api.minimax.io; use api.minimaxi.com/v1 for the CN region)',
    keyPlaceholder: 'eyJ...',
    defaultBaseUrl: MINIMAX_BASE_URL,
    imageProtocol: 'minimax',
    imageModels: ['image-01'],
    defaultImageModel: 'image-01',
    analysisModels: [],
    defaultAnalysisModel: '',
    videoAnalysis: false,
  },
  {
    id: 'custom',
    label: 'Custom',
    description: 'Any OpenAI-compatible endpoint: /images/generations and /chat/completions',
    keyPlaceholder: 'API Key',
    needsBaseUrl: true,
    defaultBaseUrl: '',
    imageProtocol: 'openai-images',
    imageModels: [],
    defaultImageModel: '',
    analysisProtocol: 'openai-chat',
    analysisModels: [],
    defaultAnalysisModel: '',
    videoAnalysis: false,
  },
]

export type MediaCapability = 'image' | 'analysis' | 'video'

export function getMediaProviderMeta(id: AiMediaProviderId): AiMediaProviderMeta | undefined {
  return AI_MEDIA_PROVIDERS.find((m) => m.id === id)
}

export function providerHasCapability(
  meta: AiMediaProviderMeta,
  capability: MediaCapability,
): boolean {
  if (capability === 'image') return !!meta.imageProtocol
  if (capability === 'video') return !!meta.analysisProtocol && meta.videoAnalysis
  return !!meta.analysisProtocol
}

export function defaultAiMediaSettings(): AiMediaSettings {
  const providers = {} as AiMediaSettings['providers']
  for (const meta of AI_MEDIA_PROVIDERS) {
    providers[meta.id] = {
      apiKey: '',
      imageModel: meta.defaultImageModel,
      analysisModel: meta.defaultAnalysisModel,
      baseUrl: meta.needsBaseUrl ? '' : undefined,
    }
  }
  return {
    imageProvider: 'genspark',
    analysisProvider: 'genspark',
    videoAnalysisProvider: 'genspark',
    providers,
  }
}

/**
 * Merge a stored media block over defaults, trimming pasted whitespace. The
 * pre-catalog `provider` field (one choice for both tools) seeds both
 * per-capability choices. Unknown provider ids are kept and simply never activate.
 */
export function resolveAiMediaSettings(
  stored: Partial<AiMediaSettings> | undefined,
): AiMediaSettings {
  const defaults = defaultAiMediaSettings()
  if (!stored) return defaults
  const providers = { ...defaults.providers }
  for (const [id, config] of Object.entries(stored.providers ?? {})) {
    if (!config || typeof config !== 'object') continue
    const base = providers[id as AiMediaProviderId]
    providers[id as AiMediaProviderId] = {
      apiKey: (config.apiKey ?? base?.apiKey ?? '').trim(),
      imageModel: (config.imageModel ?? base?.imageModel ?? '').trim(),
      analysisModel: (config.analysisModel ?? base?.analysisModel ?? '').trim(),
      ...(config.baseUrl !== undefined
        ? { baseUrl: config.baseUrl.trim() }
        : base?.baseUrl !== undefined
          ? { baseUrl: base.baseUrl }
          : {}),
    }
  }
  const legacy = stored.provider
  const analysisProvider = stored.analysisProvider ?? legacy ?? defaults.analysisProvider
  return {
    imageProvider: stored.imageProvider ?? legacy ?? defaults.imageProvider,
    analysisProvider,
    // a pre-split file used one vendor for all media analysis
    videoAnalysisProvider: stored.videoAnalysisProvider ?? analysisProvider,
    providers,
  }
}

/** Key (or base URL for custom) present — the minimum for a BYOK media provider to be honored */
export function mediaConfigUsable(
  meta: AiMediaProviderMeta,
  config: AiMediaProviderConfig | undefined,
): boolean {
  if (!config) return false
  if (meta.needsBaseUrl) return !!config.baseUrl
  return !!config.apiKey
}

/**
 * The stored provider for one capability, honored only when it exists, has
 * that capability and is usable; anything else falls back to genspark so a
 * half-filled setup degrades to the signed-in default.
 */
export function activeMediaProvider(
  settings: Pick<AiSettings, 'media'>,
  capability: MediaCapability,
): AiMediaProviderId {
  const media = settings.media
  if (!media) return 'genspark'
  const id =
    capability === 'image'
      ? media.imageProvider
      : capability === 'video'
        ? media.videoAnalysisProvider
        : media.analysisProvider
  if (!id || id === 'genspark') return 'genspark'
  const meta = getMediaProviderMeta(id)
  if (!meta || !providerHasCapability(meta, capability)) return 'genspark'
  if (!mediaConfigUsable(meta, media.providers?.[id])) return 'genspark'
  return id
}

/** the active BYOK config for one capability, or null when it runs through Genspark */
export function activeMediaConfig(
  settings: Pick<AiSettings, 'media'>,
  capability: MediaCapability,
): { provider: Exclude<AiMediaProviderId, 'genspark'>; config: AiMediaProviderConfig } | null {
  const provider = activeMediaProvider(settings, capability)
  if (provider === 'genspark') return null
  return { provider, config: settings.media!.providers[provider] }
}

function byokModel(
  settings: Pick<AiSettings, 'media'>,
  capability: MediaCapability,
): string | null {
  const active = activeMediaConfig(settings, capability)
  if (!active) return null
  const meta = getMediaProviderMeta(active.provider)!
  return capability === 'image'
    ? active.config.imageModel || meta.defaultImageModel
    : active.config.analysisModel || meta.defaultAnalysisModel
}

function capabilityAvailable(
  settings: Pick<AiSettings, 'media' | 'gskToolsEnabled'> | null | undefined,
  gskLoggedIn: boolean,
  capability: MediaCapability,
): boolean {
  if (!settings) return gskLoggedIn
  const model = byokModel(settings, capability)
  if (model !== null) return model !== ''
  return gskLoggedIn && settings.gskToolsEnabled !== false
}

/** live predicate for the generate_image tool: BYOK image model configured, or gsk login + cloud tools on */
export function imageGenerationAvailable(
  settings: Pick<AiSettings, 'media' | 'gskToolsEnabled'> | null | undefined,
  gskLoggedIn: boolean,
): boolean {
  return capabilityAvailable(settings, gskLoggedIn, 'image')
}

/** live predicate for the analyze_media tool: image analysis or video analysis reachable */
export function mediaAnalysisAvailable(
  settings: Pick<AiSettings, 'media' | 'gskToolsEnabled'> | null | undefined,
  gskLoggedIn: boolean,
): boolean {
  return (
    capabilityAvailable(settings, gskLoggedIn, 'analysis') ||
    capabilityAvailable(settings, gskLoggedIn, 'video')
  )
}

export function videoAnalysisAvailable(
  settings: Pick<AiSettings, 'media' | 'gskToolsEnabled'> | null | undefined,
  gskLoggedIn: boolean,
): boolean {
  return capabilityAvailable(settings, gskLoggedIn, 'video')
}
