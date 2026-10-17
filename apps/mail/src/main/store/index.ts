/**
 * Public surface of the mail local store (P4 lane T).
 *
 * Y lane imports from here: account stores, the full record/parameter
 * types, and the outbox retry/error helpers. Only `sqlite.ts` in this
 * directory touches `node:sqlite`.
 */
export {
  SqliteMailStore,
  openMailStore,
  closeMailStore,
  closeAllMailStores,
  deleteAccountStore,
  accountDbPath,
  defaultMailStoreDir,
  nextOutboxRetryDelay,
  sanitizeErrorText,
  INLINE_BODY_CAP_BYTES,
  SNIPPET_CHARS,
  LAST_ERROR_CHARS,
  SENT_RETENTION_DAYS,
  OUTBOX_LIST_LIMIT,
  FTS_DEFAULT_LIMIT,
  OUTBOX_RETRY_DELAYS_MS,
  OUTBOX_MAX_ATTEMPTS,
} from './store'
export type {
  AccountRecord,
  AttachmentRecord,
  AttachmentState,
  FolderInput,
  IntentState,
  MessageBody,
  MessageInput,
  OpenStoreOptions,
  OutboxRecord,
  OutboxState,
  StoredIntent,
} from './store'
export { MIGRATIONS, LATEST_STORE_VERSION } from './migrations'
export type { StoreMigration } from './migrations'
export {
  openDatabase,
  closeDatabase,
  runMigrations,
  withTransaction,
  getUserVersion,
} from './sqlite'
