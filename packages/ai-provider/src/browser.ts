/** Browser-safe settings surface. Keep Node-backed transports out of renderer bundles. */
export type {
  AiCustomCompat,
  AiCustomEndpoint,
  AiCustomModelProfile,
  AiCustomProfile,
  AiCustomProtocol,
  AiProviderConfig,
  AiProviderId,
  AiProviderMeta,
  AiSettings,
  CodexModelCatalog,
} from './types'
export type { GeminiModelInfo } from './gemini-models'
export {
  createGeminiRouter,
  GEMINI_DEFAULT_ORDER,
  GEMINI_ROUTING_EVENT,
  readGeminiChoice,
  readGeminiModels,
  readGeminiUsage,
  saveGeminiChoice,
  saveGeminiModels,
} from './gemini-routing'
export type { GeminiModelChoice } from './gemini-routing'
export {
  AI_PROVIDERS,
  activeProvider,
  cloudToolsEnabled,
  DEFAULT_MAX_OUTPUT_TOKENS,
  MAX_MAX_OUTPUT_TOKENS,
  MIN_MAX_OUTPUT_TOKENS,
  clampMaxOutputTokens,
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
export { endpointSupportsVision, getProviderAdapter, modelLacksVision } from './registry'
export { AI_MEDIA_PROVIDERS, imageGenerationAvailable, mediaAnalysisAvailable } from './media'
export { AI_SEARCH_PROVIDERS } from './search-settings'
export { agyErrorText, formatAgyResetTime } from './agy-error-text'
export type { AgyErrorLang, AgyErrorTextCode } from './agy-error-text'
export { canAttemptChat } from './readiness'
export { isKeylessProvider } from './agy-meta'
export {
  DEFAULT_LOCAL_VLM_SETTINGS,
  LOCAL_VLM_DEFAULT_MODEL,
  LOCAL_VLM_HF_REPO,
  LOCAL_VLM_MODEL_FILENAME,
  LOCAL_VLM_MMPROJ_FILENAME,
  LOCAL_VLM_PROVIDER_ID,
  LOCAL_VLM_SETTINGS_KEY,
  isLocalVlmProvider,
  localVlmDefaultsUsable,
  localVlmSettingsFrom,
  mergeLocalVlmSettings,
} from './local-vlm-meta'
export type { LocalVlmDownloadProgress, LocalVlmSettings, LocalVlmStatus } from './local-vlm-meta'
