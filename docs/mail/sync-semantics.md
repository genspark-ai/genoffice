# Mail sync and offline behavior

User-facing description of how GenOffice Mail stays in sync, what happens
offline, and what wins when both sides change the same message. It mirrors
the locked engineering contract in `apps/mail/docs/sync-spec.md` exactly;
if this page ever disagrees with that file, the spec wins and the drift
should be reported back as a DOCS-QUERY.

> H10 verification (t_86121039): re-checked against the H7 hardening
> integration tip `77278127`. The retry schedule
> (`OUTBOX_RETRY_DELAYS_MS = [5s, 30s, 2min, 10min]`), the per-driver caps
> (IMAP 50/25 MB, 100 recipients; Gmail 35 MB, 500 recipients; Graph 4/3 MB,
> 500 recipients), the server-wins conflict rule, and the security notes
> below all still match the code. No drift.

## The short version

- Mail works offline. Reading, composing, moving, flagging, and deleting
  keep working with no connection; sends and changes are queued on your
  machine and applied when you reconnect.
- When you and the server disagree, the server wins for flags, labels, and
  moves. Your queued sends and your draft text always win.
- Every account syncs on its own, within that provider's size caps.

## Offline behavior

- Going offline never loses work. Composing while offline puts the message
  in the outbox in `queued` state; the app shows an offline indicator.
- Moves, flags, deletes, and label changes made offline are queued as
  intents and replayed first-in-first-out when the connection returns.
- On reconnect, queued outbox sends go first, then queued intents.
- A full re-sync (see below) never deletes your queued outbox rows, your
  drafts, or your queued intents.

## How syncing works per account

Each account syncs independently in the background:

- IMAP: incremental sync where the server supports it, otherwise a UID-range
  fetch. If the server's folder identity (`UIDVALIDITY`) changes, the folder
  is fully re-fetched and you see a "resynced from the server" notice.
  Live updates where advertised, otherwise a 30-second poll that backs off
  (doubling, up to 10 minutes) on errors and resets on success.
- Gmail: incremental history sync from the stored position, polled every
  30-60 seconds. If the stored position expires, a full re-list runs and a
  fresh position is saved.
- Outlook/Graph: incremental delta sync, polled every 30-60 seconds. If the
  stored delta link expires, a full re-list runs and a new link is saved.

## Outbox retry

- Sends go one at a time per account; different accounts send concurrently.
- A failed send retries automatically on a 5 s, 30 s, 2 min, 10 min
  schedule, up to 5 attempts, then it is marked `failed`.
- You can retry a failed send manually (attempts reset, back at the head of
  the queue) or cancel a queued send. A cancelled send is never sent, even
  mid-retry.
- Safety against double delivery: every queued send carries a stable id that
  doubles as the idempotency key, and the Sent folder is reconciled on the
  next sync so a crash between send and bookkeeping still delivers at most
  once. Sent confirmations are kept 7 days; cancelled rows stay until you
  dismiss them. The outbox list shows the newest 200 entries.
- Failures surface as a failed outbox entry plus a sync-state notice; error
  text is a sanitized provider summary, never message content or secrets.

## The conflict rule (plain language)

- Server wins for flags, labels, and moves. If you flagged a message offline
  and someone (or another device) changed it on the server first, the
  server's version is what you see, plus a "change conflicted" notice.
  A failed change never merges silently.
- Client wins ONLY for two things: messages you queued to send, and the text
  of your local drafts.
- Technically, each offline change is replayed once; if the message changed
  on the server meanwhile, the app re-fetches it and retries once, and if it
  is still stale the change is marked failed with a conflict notice.

## Per-account send caps

Caps are enforced twice: when you queue the send (fast reject) and again at
send time in case caps changed while you were offline.

| Account type  | Total cap | Per-attachment cap        | Recipient cap |
| ------------- | --------- | ------------------------- | ------------- |
| IMAP/SMTP     | 50 MB     | 25 MB                     | 100           |
| Gmail         | 35 MB     | (total only)              | 500           |
| Outlook/Graph | 4 MB      | 3 MB (warn at 3 MB total) | 500           |

Search covers subjects, senders, snippets, and plaintext; full HTML bodies
and attachment contents are never searched. Bodies over 1 MB keep the first
1 MB inline and are marked truncated.

## Security notes (locked, documented here — not decided here)

- Links: external links open only through the `safeExternalUrl` gate, which
  stays http/https-only everywhere except the compose/mail open-link
  call-sites, where a per-call `mailto:` allowlist applies (Alya P1 policy).
- Credentials: refresh tokens and app passwords live only in
  `safeStorage`-encrypted blobs referenced by opaque `credRef` handles.
  Access tokens stay in memory only.
- Linux without a keyring: accounts are session-only with an explicit
  warning shown; secrets are never written to disk as a silent fallback.
- No-plaintext rule: secrets never appear in SQLite, JSON stores, logs,
  telemetry, or crash reports; error text carries references and counts only.
