/**
 * Sync + outbox replay engine entry point (Y lane owns this directory).
 * Drivers run in the main process only; the renderer talks to the engine
 * through the typed+zod-validated sync/outbox IPC channels.
 */
export { capsForKind, checkOutboxCaps, countOutboxRecipients, estimateOutboxBytes } from './caps'
export type { AccountKind, CapCheck, CapVerdict, SendCapSet } from './caps'
export { replayIntents } from './conflicts'
export type { IntentDriver, IntentOutcome, IntentStore, ReplayReport } from './conflicts'
export { SyncEngine } from './engine'
export type {
  OfflineStatusPayload,
  EngineEmit,
  ResyncReason,
  SyncDriverPort,
  SyncEngineOptions,
  SyncFolderResult,
  SyncStorePort,
} from './engine'
export { registerSyncHandlers } from './ipc'
export type { SyncIpcBus, SyncIpcDeps, SyncIpcHandle } from './ipc'
export { OutboxCapError, OutboxQueue } from './outbox'
export type { OutboxRecord, OutboxState, OutboxStatus, OutboxQueueOptions } from './outbox'
export {
  OUTBOX_MAX_ATTEMPTS,
  OUTBOX_RETRY_DELAYS_MS,
  POLL_BASE_MS,
  POLL_CEILING_MS,
  nextPollDelay,
  outboxRetryDelay,
} from './poll'
export type { PollSchedule } from './poll'
export { containsPii, sanitizeError } from './redact'
