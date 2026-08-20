export type {
  AiChatRequest,
  AiChatResponse,
  AiProviderConfig,
  AiProviderId,
  AiProviderMeta,
  AiSettings,
  AiStreamChunk,
  AiStreamRequest,
  GenSparkAccountStatus,
  LegacyAiSettings,
} from './types'
export {
  AI_PROVIDERS,
  GENSPARK_LLM_BASE_URLS,
  OLLAMA_DEFAULT_BASE_URL,
  defaultAiSettings,
  isProviderConfigured,
  providerRequiresApiKey,
  resolveAiSettings,
} from './providers'
export { listOllamaModels, ollamaListStatus, type OllamaModelsResult } from './ollama'
export {
  chunkText,
  cosineSimilarity,
  embedWithOllama,
  ollamaApiRoot,
  pickEmbeddingModel,
  searchChunks,
  type WorkspaceIndexHit,
  type WorkspaceIndexResult,
  type WorkspaceSearchResult,
} from './embed'
export { testProviderConnection } from './connection'
export { createWorkspaceSkill, type WorkspaceSearchFn } from './workspace-skill'
export type {
  AiConnectionStatus,
  AiConnectionTestInput,
  AiConnectionTestResult,
} from './connection'
export { chatForProvider } from './chat'
export { setRescueFetch } from './fetch'
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
