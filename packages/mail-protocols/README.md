# Lane B: ImapSmtpDriver (IMAP + SMTP)

Phase-2 P1-protocols Lane B. Implements the `MailProvider` contract from the
amended arch spec (section 3.2) for generic IMAP/SMTP accounts with
app-password or injected credentials.

## Files

- `src/types.ts` — canonical contract (spec section 3.2/4.2 verbatim, Lane A).
  Lane B's temporary `src/mail-provider.ts` copy was removed at lane merge.
- `src/imap-smtp-driver.ts` — the driver. Pure TypeScript: no Electron, no
  safeStorage, no sockets. Connection objects, credentials, and stores are
  injected by the main-process caller (`apps/mail/src/main`, locked D6).
- `src/imap-transport.ts` — production adapters: imapflow 1.5.0 (MIT) for
  IMAP, nodemailer 10.1.1 (MIT-0) for SMTP with path-streamed attachments.
- `src/imap-fakes.ts` — in-memory IMAP/SMTP doubles for tests and staging.
- `tests/imap-smtp-driver.test.ts` — round-trip matrix row, UIDVALIDITY
  resync, offline intents, caps, IDLE/poll watch paths.

## Sync behavior

- IDLE where advertised, else 30-60 s poll (default 45 s) with doubling
  backoff to a 5 min ceiling.
- CONDSTORE/MODSEQ where advertised, else UID-range fetch.
- UIDVALIDITY change triggers full folder re-fetch plus a `resync` event.
  Resync never deletes queued outbox sends, drafts, or intents: the driver
  owns none of those tables (P3 sync engine), pinned by fixture.
- Offline move/flag/delete queue as idempotent intents
  (`targetUid + expectedModseq`). Stale modseq re-fetches then retries once,
  else marks the intent failed and emits `conflict`.

## Send caps

Defaults: warning above 25 MiB total, hard reject above 50 MiB total,
25 MiB per attachment, 100 recipients. All overridable via
`ImapSmtpSendCaps` in the account config.

## Security notes (for the Alya eyes lane)

- This driver never imports `safeStorage`, `shell.openExternal`, or any
  refresh-token/password store. Credentials arrive injected from
  `credential-vault.ts` (single safeStorage caller, R2).
- Test credentials in this package are hard-coded fakes
  (`test-user`/`test-pass`, `.example.local` hosts). No real secrets.

## Commands

- `npm run test -w @genoffice/mail-protocols`
- `npm run typecheck -w @genoffice/mail-protocols`
- `npm run licenses` (root, allowlist gate)
- `npm run notices` (root, third-party notices)
