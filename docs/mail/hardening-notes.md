# Mail hardening notes (H1 spec → what shipped)

User-visible summary of the GenOffice Mail hardening phase plus the
contributor pins each lane left behind. Acceptance spec:
`apps/mail/docs/hardening-spec.md` (H1, t_0672f5e2, `dc39e131`).
Integrated at the H7 tip `77278127` (H2–H6 merged cleanly, zero conflicts).

> H10 status (t_86121039): H7 QA verification (t_f909e70e) was still running
> when this page was written. Lane SHAs and counts below come from the
> lanes' own handoffs; re-check the H7 verdict card before release.

## Credentials (H2, `570b7fea`, verdict GO)

- Refresh tokens and app passwords live only in `safeStorage`-encrypted
  blobs referenced by opaque handles. Access tokens stay in memory only.
- Least-privilege OAuth scopes: Gmail asks `gmail.modify` only; Outlook asks
  delegated `Mail.ReadWrite` + `offline_access` + `User.Read` — never
  full-mailbox or application permissions.
- Linux without a keyring: accounts are session-only with an explicit banner
  (`mail.noKeyringBanner`); secrets are never written to disk as a silent
  fallback. Rotating credentials wipes and re-authenticates.
- Contributor pin: 31 C-1..C-11 tests (`credential-store.test.ts`); mail
  suite 303/303, protocols 107/107 at the lane HEAD.

## PII redaction on AI paths (H3, `58c14d9b`)

- The PII matrix in `ai-features.md` holds; hostile/prompt-injection
  fixtures pin it: message content is untrusted data in every skill prompt,
  and forged or cross-account citations are rejected.
- Contributor pin: 13 hostile-fixture tests (`ai-redaction-hostile.test.ts`);
  mail suite 285/285 at the lane HEAD. Alya sec re-eyes required before
  merge (redaction paths changed).

## Transport, scope, rate (H4, `c31e7c96`)

- Citation references resolve exactly in the account they came from
  (D-H1-C3); a stale account hint fails loud, never cross-account.
  `cite-resolve` returns matches from every in-scope account holding the
  reference, so ambiguity is visible, not silent.
- AI retrieval is paced per account (token bucket, H1 P-9); the empty-bucket
  error is PII-free (`mail: rate-limited, retry in <n>s`).
- Contributor pins: two-account same-uid fixtures (hint-hits, stale-fails,
  no-hint-deterministic, resolve-returns-both) + P-6..P-9 budget pins; mail
  suite 302/302 at the lane HEAD. Measured on the shared box: 100k FTS p95
  256–464 ms vs the 200 ms budget (LIKE fallback stays under 1 s) — recorded,
  not gated. Alya sec re-eyes required before merge (scope paths changed).

## Attachment quarantine (H5, `9ae6d323`)

- Every attachment travels staged-temp → quarantine-scan → store with a
  verdict: **clean** (preview allowed), **quarantined** (held for review —
  save and open only), **blocked** (dropped with a reason).
- Path traversal and hostile file names are rejected; executables are never
  previewed (confirm-gated open only); previews reuse the file-parse
  extractor, which never runs on executables. Per-driver send caps resolve
  in main, so over-cap verdicts match the compose-send errors exactly.
- Contributor pins: 34 pipeline tests (`mail-quarantine-pipeline.test.tsx`);
  mail suite 306/306, protocols 107/107, zero new external deps at the lane
  HEAD. Alya sec re-eyes required before merge.

## 100k list performance (H6, `71069454`)

- Folder lists stay windowed over 100k messages (at most ~15 mounted rows),
  search input debounced at 200 ms with supersede + stale-discard.
- Contributor pins (CI-enforced): scroll handler p95 0.003 ms vs 8 ms
  budget; ≤ 30 mounted nodes; debounce within [150, 250] ms; 1 MiB message
  sanitize ~8–9 ms in real Chromium vs the 200 ms budget (jsdom figures are
  harness overhead, guarded only against hangs). Mail suite 284/284,
  protocols 107/107 at the lane HEAD, additive-only list files.

## Ship order and boundaries (unchanged)

- IMAP/SMTP with app passwords first (no verification needed); Gmail/Graph
  API drivers stay behind the OAuth verification landing — see
  `oauth-verification.md`.
- Zero new production dependencies; English-only; synthetic mail plus a
  stubbed provider in tests; local-dev only; no upstream PRs (standing
  instruction t_3169df70 — hardening green plus owner nod gates any push).
