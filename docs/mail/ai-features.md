# Mail AI features

What the AI panel can do inside GenOffice Mail, what data each feature may
see, and where the boundaries are. Engineering contract:
`apps/mail/docs/ai-spec.md` (spec wins on any disagreement); redaction
audit: H3 (t_2d86bb14); scope/rate implementation: H4 (t_0d786056).

> H10 pass (t_86121039) against the H7 hardening integration tip `77278127`.
> Local dev and tests use a stubbed provider and synthetic mail only — no
> real accounts, no real credentials.

## The four features

- **Triage** — sorts one folder into digest buckets (priority, needs-reply,
  newsletter, noise). It sees headers (from/to/subject/date), folder,
  snippet, and labels/flags ONLY. Bodies are never attached and cannot be
  fetched from this skill.
- **Summarize / catch-up** — summarizes the thread or folder you selected.
  Bodies are fetched ONLY when you explicitly ask, and only for the cited
  messages. Every summary bullet carries a message-id citation.
- **Compose-assist** — drafts from your bullets, rewrites tone, or replies
  from the thread you opened. Output is text returned for insertion: the AI
  proposes, you send. It never sends, never touches the outbox, and never
  harvests addresses beyond recipients you typed.
- **Ask-mail** — searches your mail (full-text, headers-first) and answers
  with a per-claim message-id citation payload. Factual claims without a
  citation are rejected and retried, so an un-cited answer never stands.

## Account scoping (the picker is the boundary)

- Every AI call carries the account set selected in the account picker
  (non-empty, at most 8). Results outside that set are refused; results from
  a stale picker selection are discarded.
- Citations round-trip their account, so a reference resolves in the account
  it came from. If that account hint is stale (removed or unknown), the call
  fails loudly with an error — it never silently shows another account's
  mail. Without a hint, the first account in picker order wins,
  deterministically.
- Asking across accounts in one question is refused: every tool call echoes
  its scope, and out-of-scope results are dropped.

## What the AI may see (and never sees)

| Feature              | May see                                                  | Never sees                                     |
| -------------------- | -------------------------------------------------------- | ---------------------------------------------- |
| Triage               | Headers, folder, snippet, labels/flags                   | Bodies, attachments, other accounts' mail      |
| Summarize            | Bodies of the thread/folder you selected                 | Mail outside the picker scope                  |
| Compose-assist       | Your draft + the thread you opened                       | Address-book harvesting beyond your recipients |
| Ask-mail             | Retrieved messages with citations                        | Cross-account joins; un-cited claims as fact   |
| Logs / crash reports | Allowlisted fields only (account hash, outcome, latency) | Bodies, snippets, subjects, addresses, tokens  |

Skill traces and checked-in fixtures use synthetic mail only. Reply
addressing is preserved verbatim — the AI never redacts recipient addresses.
Mail bodies in hostile messages cannot steer the AI: tool instructions treat
message content as untrusted data, and forged or cross-account citations are
rejected even when mixed with valid ones.

## Provider and pacing

- AI calls go only to your configured provider: your own key (BYOK) or the
  signed-in Genspark proxy. No training use, no secondary recipients. Local
  full-text search stays on your machine.
- Rapid repeated AI requests are paced per account (token bucket on
  `mail:ai-retrieve`). When the bucket is empty the panel reports
  `mail: rate-limited, retry in <n>s` — timing only, never content.
