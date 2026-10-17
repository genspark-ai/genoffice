# Contributor notes: fixture-driven resync tests

How to run and extend the mail sync test suite. All tests use scripted fakes;
no real accounts, no real credentials, local-dev only.

> H10 note (t_86121039): procedure unchanged at the H7 hardening integration
> tip `77278127`. Case 10 (PII redaction) is additionally covered by the H3
> hostile-fixture suite (`ai-redaction-hostile.test.ts`, 13 tests).

## Running the suite

```bash
npx vitest run apps/mail/tests/sync-*.test.ts   # Y lane: driver sync semantics
npx vitest run apps/mail/tests/store-*.test.ts  # T lane: SQLite store + FTS
```

O-lane UI tests use the same fakes, never the drivers.

## The 10 named cases (spec section 7)

1. `uidvalidity-change-resync` — fake IMAP reports new `UIDVALIDITY`;
   assert full re-fetch, folder rows replaced, `resync` event with
   `reason:'uidvalidity-changed'`, outbox/drafts/intents intact.
2. `history-expiry-resync` — fake Gmail answers 404 to `history.list`;
   assert full re-list, fresh `historyId` persisted, `resync` event emitted.
3. `delta-expiry-resync` — fake Graph rejects the delta link; assert full
   re-list and new `deltaToken` persisted.
4. `offline-compose-reconnect-send` — go offline, compose-send (queued),
   assert `queued` state plus offline status; reconnect, assert FIFO send,
   `sent` state, Sent reconcile marks at-most-once.
5. `flag-move-conflict-server-wins` — server flag changed while an offline
   flag intent was queued; assert intent ends `failed` with a `conflict`
   event and visible state equals server state.
6. `idempotent-replay-double-delivery-safe` — crash-simulate between send
   and mark (fake transport records sends); assert exactly one server-side
   message after reconnect plus reconcile.
7. `poll-backoff-schedule` — fake timers assert 30 s base, doubling backoff
   to the 10 min ceiling, reset on success, for all three drivers.
8. `outbox-cap-enforcement` — over-cap send rejected at enqueue AND at send
   with per-driver messages (IMAP 50 MB, Gmail 35 MB, Graph 4 MB).
9. `fts-mirror-and-rebuild` — insert/update/delete keeps the FTS index in
   sync; full resync rebuilds it; a CJK subject query retrieves.
10. `pii-redaction` — sync-error and outbox-failure paths assert no
    body/snippet/secret bytes in `lastError`, events, or logs (Alya signs
    the assertion list).

## Adding a fixture case

- Put driver cases in `apps/mail/tests/sync-*.test.ts` (Y owns) and store
  cases in `apps/mail/tests/store-*.test.ts` (T owns); never cross lanes.
- Reuse the pure helpers under test (`decideSync`, `nextIntent`,
  `resolveIntentAttempt`, `intentKey`); do not rewrite them in the test.
- Resync cases must always assert outbox/drafts/intents survive the resync.
- Name the test exactly after the case above so lane reports match the spec.
