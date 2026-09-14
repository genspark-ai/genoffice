/**
 * `@genoffice/chat-runtime` — unified Chat model for GenOffice.
 *
 * Mounts a `ChatRuntime` against any app's existing `AgentSkill` +
 * transport, providing Session / Run / ChangePlan semantics and an
 * `AIError` classifier that the four renderer panels share.
 */

export * from './types.js'
export * from './errors.js'
export { createSession, loadOrCreateSession, appendMessage, replaceMessages } from './session.js'
// `JsonlChatPersistence` is intentionally NOT re-exported here:
// `persistence.ts` uses `node:fs`/`node:path` and breaks browser bundles.
// Import directly from `@genoffice/chat-runtime/persistence` (Node-only)
// or use `MemoryChatPersistence` for in-memory tests.
export { MemoryChatPersistence } from './persistence.js'
export type { JsonlPersistenceOptions } from './persistence.js'
export { normalizeChangePlan, summarizeChangePlan } from './change-plan.js'
export type { NormalizeChangePlanInput } from './change-plan.js'
export { ChatRuntime } from './runtime.js'
export { startRun } from './run.js'
export type { RunHandle, RunBridge, StartRunOptions } from './run.js'
