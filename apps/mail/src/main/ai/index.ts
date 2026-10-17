/**
 * Mail AI provider transport (Lane W owns). Account-scoped envelope,
 * retrieval core over the Lane-T store, redaction-backed prompt hygiene,
 * the single mail:ai-retrieve IPC registration, and the stubbed-provider
 * fixture for tests and stubbed e2e. Imports store/sync APIs, never edits
 * them (the spec-blessed redact.ts extension point excepted).
 */

export type { AiScope } from './scope'
export { AI_SCOPE_MAX_ACCOUNTS, isScopeStale, normalizeAiScope } from './scope'
export type {
  AiBodyEntry,
  AiCitation,
  AiRetrieveDeps,
  AiRetrieveOp,
  AiRetrieveRequest,
  AiRetrieveResponse,
  AiRetrieveStorePort,
} from './retrieve'
export {
  AI_RETRIEVE_DEFAULT_LIMIT,
  AI_RETRIEVE_MAX_LIMIT,
  AI_RETRIEVE_MAX_REFS,
  aiRetrieve,
} from './retrieve'
export {
  AI_RETRIEVE_RATE_BURST,
  AI_RETRIEVE_RATE_SUSTAINED_PER_MINUTE,
  createAiRetrieveRateLimiter,
  resetSharedAiRetrieveRateLimiter,
  sharedAiRetrieveRateLimiter,
} from './rate-limit'
export type { AiRateLimiter } from './rate-limit'
export { defaultAiRetrieveDeps } from './stores'
export type { AiIpcBus } from './ipc'
export { registerAiRetrieveHandlers, resetAiRetrieveHandlers } from './ipc'
export type { StubSeedAccount, StubSeedMessage } from './stub'
export { collectStubStream, createStubAiDeps, stubRef, twoAccountStubSeeds } from './stub'
export { sameUidTwoAccountSeeds, stubRefHinted } from './stub'
