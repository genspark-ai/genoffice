/**
 * SQLite migrations for the GenOffice Mail local store (P4 lane T).
 *
 * The DDL below follows apps/mail/docs/sync-spec.md section 1.2 verbatim
 * with two lane-local adjustments, both reported in the lane handoff:
 * - The FTS5 tokenizer is plain `unicode61`. The spec text names
 *   `unicode61 "remove_diacritics=2"`, but the Node-bundled SQLite build
 *   (verified on 3.53.4) rejects that option with a parse error, so the
 *   verbatim form cannot run. Diacritic folding level is the only delta;
 *   CJK retrieval is pinned by tests. Needs a SPEC-DELTA ratification.
 * - Migration 2 adds `messages.bodyTruncated`, the T-local addition the
 *   spec explicitly allows, so over-cap inline bodies stay detectable.
 *
 * Migrations are forward-only; each runs inside one transaction and bumps
 * PRAGMA user_version to its own version on success.
 */

export interface StoreMigration {
  version: number
  name: string
  sql: string
}

const MIGRATION_1_SCHEMA = `
CREATE TABLE accounts (
  accountId   TEXT PRIMARY KEY,
  label       TEXT NOT NULL,
  kind        TEXT NOT NULL CHECK (kind IN ('imap','gmail','graph')),
  sessionOnly INTEGER NOT NULL DEFAULT 0,
  credRef     TEXT,
  createdAt   TEXT NOT NULL
) STRICT;

CREATE TABLE folders (
  accountId TEXT NOT NULL REFERENCES accounts(accountId) ON DELETE CASCADE,
  folderId  TEXT NOT NULL,
  name      TEXT NOT NULL,
  role      TEXT NOT NULL CHECK (role IN
              ('inbox','sent','drafts','trash','spam','archive','custom')),
  unread    INTEGER NOT NULL DEFAULT 0,
  PRIMARY KEY (accountId, folderId)
) STRICT;

CREATE TABLE messages (
  accountId TEXT NOT NULL,
  folderId  TEXT NOT NULL,
  uid       TEXT NOT NULL,
  msgId     TEXT,
  threadId  TEXT NOT NULL,
  fromJson  TEXT NOT NULL,
  toJson    TEXT NOT NULL,
  ccJson    TEXT NOT NULL DEFAULT '[]',
  subject   TEXT NOT NULL DEFAULT '',
  dateIso   TEXT NOT NULL,
  seen      INTEGER NOT NULL DEFAULT 0,
  flagged   INTEGER NOT NULL DEFAULT 0,
  answered  INTEGER NOT NULL DEFAULT 0,
  draft     INTEGER NOT NULL DEFAULT 0,
  labelsJson TEXT NOT NULL DEFAULT '[]',
  snippet   TEXT NOT NULL DEFAULT '',
  bodyText  TEXT NOT NULL DEFAULT '',
  bodyHtmlRef TEXT,
  hasAttach INTEGER NOT NULL DEFAULT 0,
  modseq    TEXT,
  updatedAt TEXT NOT NULL,
  PRIMARY KEY (accountId, folderId, uid),
  FOREIGN KEY (accountId, folderId) REFERENCES folders(accountId, folderId) ON DELETE CASCADE
) STRICT;

CREATE VIRTUAL TABLE messages_fts USING fts5(
  subject, fromText, snippet, bodyText,
  content='messages', content_rowid='rowid',
  tokenize='unicode61'
);

CREATE TABLE attachments (
  accountId  TEXT NOT NULL,
  msgUid     TEXT NOT NULL,
  contentId  TEXT,
  filename   TEXT NOT NULL,
  mime       TEXT NOT NULL,
  size       INTEGER NOT NULL,
  blobPath   TEXT NOT NULL,
  state      TEXT NOT NULL DEFAULT 'stored'
             CHECK (state IN ('staged','quarantined','stored')),
  PRIMARY KEY (accountId, msgUid, filename)
) STRICT;

CREATE TABLE sync_state (
  accountId     TEXT NOT NULL,
  folderId      TEXT NOT NULL,
  uidValidity   INTEGER,
  highestModseq TEXT,
  historyId     TEXT,
  deltaToken    TEXT,
  lastSyncAt    TEXT NOT NULL,
  PRIMARY KEY (accountId, folderId)
) STRICT;

CREATE TABLE outbox (
  id          TEXT PRIMARY KEY,
  accountId   TEXT NOT NULL REFERENCES accounts(accountId) ON DELETE CASCADE,
  reqJson     TEXT NOT NULL,
  state       TEXT NOT NULL DEFAULT 'queued'
              CHECK (state IN ('queued','sending','failed','sent','cancelled')),
  attempts    INTEGER NOT NULL DEFAULT 0,
  nextRetryAt TEXT,
  lastError   TEXT,
  queuedAt    TEXT NOT NULL,
  sentAt      TEXT
) STRICT;

CREATE TABLE drafts (
  accountId TEXT NOT NULL,
  threadKey TEXT NOT NULL,
  text      TEXT NOT NULL DEFAULT '',
  updatedAt TEXT NOT NULL,
  PRIMARY KEY (accountId, threadKey)
) STRICT;

CREATE TABLE intents (
  id             TEXT PRIMARY KEY,
  accountId      TEXT NOT NULL,
  kind           TEXT NOT NULL CHECK (kind IN ('move','flag','delete','label')),
  targetUid      TEXT NOT NULL,
  folderId       TEXT NOT NULL,
  expectedModseq TEXT,
  payloadJson    TEXT NOT NULL,
  state          TEXT NOT NULL DEFAULT 'queued'
                 CHECK (state IN ('queued','acked','failed')),
  attempts       INTEGER NOT NULL DEFAULT 0,
  createdAt      TEXT NOT NULL
) STRICT;

CREATE INDEX idx_messages_thread ON messages(threadId);
CREATE INDEX idx_messages_account_date ON messages(accountId, dateIso DESC);
CREATE INDEX idx_outbox_account_state ON outbox(accountId, state);
CREATE INDEX idx_intents_account_state ON intents(accountId, state);

CREATE TRIGGER messages_ai AFTER INSERT ON messages BEGIN
  INSERT INTO messages_fts(rowid, subject, fromText, snippet, bodyText)
  VALUES (new.rowid, new.subject, new.fromJson, new.snippet, new.bodyText);
END;

CREATE TRIGGER messages_ad AFTER DELETE ON messages BEGIN
  INSERT INTO messages_fts(messages_fts, rowid, subject, fromText, snippet, bodyText)
  VALUES ('delete', old.rowid, old.subject, old.fromJson, old.snippet, old.bodyText);
END;

CREATE TRIGGER messages_au AFTER UPDATE ON messages BEGIN
  INSERT INTO messages_fts(messages_fts, rowid, subject, fromText, snippet, bodyText)
  VALUES ('delete', old.rowid, old.subject, old.fromJson, old.snippet, old.bodyText);
  INSERT INTO messages_fts(rowid, subject, fromText, snippet, bodyText)
  VALUES (new.rowid, new.subject, new.fromJson, new.snippet, new.bodyText);
END;
`

const MIGRATION_2_BODY_TRUNCATED = `
ALTER TABLE messages ADD COLUMN bodyTruncated INTEGER NOT NULL DEFAULT 0;
`

export const MIGRATIONS: StoreMigration[] = [
  { version: 1, name: 'p4-spec-schema', sql: MIGRATION_1_SCHEMA },
  { version: 2, name: 'messages-body-truncated', sql: MIGRATION_2_BODY_TRUNCATED },
]

export const LATEST_STORE_VERSION = MIGRATIONS[MIGRATIONS.length - 1].version
