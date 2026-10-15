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
