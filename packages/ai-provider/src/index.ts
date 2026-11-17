export type {
  AiCustomCompat,
  AiCustomEndpoint,
  AiCustomModelProfile,
  AiCustomProfile,
  AiCustomProtocol,
  AiChatRequest,
  AiChatResponse,
  AiMediaProviderConfig,
  AiMediaProviderId,
  AiMediaProviderMeta,
  AiMediaSettings,
  AiAnalysisProtocol,
  AiImageProtocol,
  AiSearchProviderId,
  AiSearchProviderMeta,
  AiSearchSettings,
  CodexModelCatalog,
  AiProviderConfig,
  AiProviderId,
  AiProviderMeta,
  AiSettings,
  AiStreamChunk,
  AiTokenUsage,
  AiStreamRequest,
  GenSparkAccountStatus,
  LegacyAiSettings,
} from './types'
export {
  AI_PROVIDERS,
  DEFAULT_MAX_OUTPUT_TOKENS,
  GENSPARK_LLM_BASE_URLS,
  MAX_MAX_OUTPUT_TOKENS,
  MIN_MAX_OUTPUT_TOKENS,
  activeProvider,
  clampMaxOutputTokens,
  cloudToolsEnabled,
  defaultAiSettings,
  maxOutputTokensOf,
  resolveAiSettings,
} from './providers'
export {
  activeCustomEndpoint,
  customEndpointLabel,
  newCustomEndpointId,
  removeCustomEndpoint,
  resolveCustomEndpoints,
  selectCustomEndpoint,
  upsertCustomEndpoint,
} from './custom-endpoints'
export {
  AI_CUSTOM_PROTOCOLS,
  declaredModelProfile,
  normalizeCustomProfile,
  resolveCustomProfile,
} from './custom-profile'
export type { ResolvedCustomProfile } from './custom-profile'
export { resolveProfileCredential } from './credential-env'
export type { EnvLookup } from './credential-env'
export {
  AI_MEDIA_PROVIDERS,
  GEMINI_MEDIA_BASE_URL,
  OPENAI_IMAGES_BASE_URL,
  activeMediaConfig,
  activeMediaProvider,
  defaultAiMediaSettings,
  getMediaProviderMeta,
  imageGenerationAvailable,
  mediaAnalysisAvailable,
  mediaConfigUsable,
  providerHasCapability,
  resolveAiMediaSettings,
  videoAnalysisAvailable,
} from './media'
export type { MediaCapability } from './media'
export {
  AI_SEARCH_PROVIDERS,
  activeSearchProvider,
  defaultAiSearchSettings,
  resolveAiSearchSettings,
} from './search-settings'
export {
  analyzeMediaWithProvider,
  generateImageWithProvider,
  sniffImageMime,
  testMediaProvider,
} from './media-protocols'
export type {
  AnalyzeMediaInput,
  ByokMediaProviderId,
  GenerateImageInput,
  MediaBlob,
} from './media-protocols'
export {
  AI_PROVIDER_ADAPTERS,
  getProviderAdapter,
  modelLacksVision,
  normalizeBaseUrl,
} from './registry'
export type {
  AiProtocol,
  ProviderAdapter,
  ProviderCapabilities,
  ResolvedEndpoint,
} from './registry'
export {
  applySettingsWrite,
  decideSettingsWrite,
  sanitizeAiSettings,
  sanitizeCliPath,
  validCliPath,
  type SettingsWriteDecision,
} from './ai-settings-guard'
export { chatForProvider } from './chat'
export { canAttemptChat } from './readiness'
export { listGeminiModels } from './gemini-models'
export type { GeminiModelInfo } from './gemini-models'
export { setAiUserAgent, setRescueFetch } from './fetch'
export { isAiNetworkError } from './network-error'
export { isAiOverloadedError } from './overload-error'
export { isAiRegionBlockedError } from './region-error'
export { parseOutputCapRejection } from './output-cap'
export { AiCreditsError, sseLines, streamForProvider } from './stream'
export type { StreamCallbacks } from './stream'
export {
  AI_CHAT_RESPONSE_TIMEOUT_MS,
  AI_CONNECT_TIMEOUT_MS,
  AI_IDLE_TIMEOUT_MS,
  AiTimeoutError,
  createStreamWatchdog,
} from './watchdog'
export type { StreamWatchdog } from './watchdog'

export { setAiErrorLogger } from './stream'
export type { AiErrorDiagnostic } from './stream'
export { setSystemAddendum, withSystemAddendum } from './system-addendum'
export {
  agyDefaultsUsable,
  agyUsabilityKnown,
  onAgyUsableChange,
  setAgyUsable,
} from './agy-default'
export { agyUsableForDefaults, probeAgyUsable } from './agy-detect'
export { classifyAiStreamError } from './stream-error'
export type { AiStreamErrorFields } from './stream-error'
export { AgyError, agyErrorKindOf, isAgyError } from './agy-errors'
export type { AgyErrorKind } from './agy-errors'
export { agyChoice } from './agy-choice'
export type { AgyChoice } from './agy-choice'
export { prewarmAgy, shutdownAgyResidentServers, retireAgyResidentScope } from './agy-resident'
export { isKeylessProvider } from './agy-meta'
export {
  DEFAULT_LOCAL_VLM_SETTINGS,
  LOCAL_VLM_DEFAULT_MODEL,
  LOCAL_VLM_HF_REPO,
  LOCAL_VLM_MODEL_FILENAME,
  LOCAL_VLM_MMPROJ_FILENAME,
  LOCAL_VLM_PROVIDER_ID,
  LOCAL_VLM_SAMPLING_DEFAULTS,
  LOCAL_VLM_SETTINGS_KEY,
  isLocalVlmProvider,
  localVlmDefaultsUsable,
  localVlmSettingsFrom,
  mergeLocalVlmSettings,
  setLocalVlmUsable,
} from './local-vlm-meta'
export type { LocalVlmDownloadProgress, LocalVlmSettings, LocalVlmStatus } from './local-vlm-meta'
export {
  LocalVlmManager,
  getLocalVlmRuntime,
  isValidGgufFile,
  registerLocalVlmRuntime,
  selectLlamaReleaseAsset,
} from './local-vlm-manager'
export type {
  LlamaReleaseAsset,
  LocalVlmManagerDeps,
  LocalVlmRuntimeBridge,
} from './local-vlm-manager'
export {
  ensureLocalVlmEndpoint,
  localVlmRequestConfig,
  shouldFallbackFromAgyToLocalVlm,
} from './local-vlm-fallback'
