# Mail help-langs gate checklist

Gate checklist for every PR that adds or changes user-visible Mail strings:
the mail domain strings table must stay in parity before merge. Background
procedure (single-file `strings.ts`, `zh` defines the key set) is the P1 i18n
rule; this page is the per-PR gate plus the completeness inventory.

> H10 verification (t_86121039): inventory below re-counted against the H7
> hardening integration tip `77278127`. `mail-strings.test.ts` green is the
> binding gate; counts are informational and drift only when a lane adds keys.

## Parity rule

- Mail strings live in exactly one file:
  `apps/mail/src/renderer/i18n/strings.ts`.
- `zh` defines the key set; every other locale carries exactly the same keys
  with no empty values (pinned by `apps/mail/tests/mail-strings.test.ts`,
  which also pins a frozen inventory of representative keys).
- Non-zh/en locales share the `en` table until their prefix is filled.

## Completeness inventory (H7 tip `77278127`)

- `zh` key set: **1243 keys** across 20 locales (`zh`, `en`, `ja`, `ko`,
  `fr`, `de`, `es`, `th`, `id`, `ru`, `ar`, `pt`, `it`, `pl`, `cs`, `nl`,
  `ms`, `he`, `hi`, `zh-TW`).
- Keys per domain (`mail.<domain>.*` in `zh`):

| Domain                                                 | Keys      | Owns (lane)                                                                                                         |
| ------------------------------------------------------ | --------- | ------------------------------------------------------------------------------------------------------------------- |
| `compose`                                              | 287       | Lane C + H5 quarantine-compose suffix                                                                               |
| `reader`                                               | 266       | Lane R                                                                                                              |
| `ai`                                                   | 247       | P5 AI + H3/H4                                                                                                       |
| `folders`                                              | 152       | Lane R                                                                                                              |
| `list`                                                 | 95        | Lane R + H6                                                                                                         |
| `accounts`                                             | 95        | H2 credential UI                                                                                                    |
| `search`                                               | 57        | Lane R                                                                                                              |
| `outbox`                                               | 16        | P4 outbox UI                                                                                                        |
| `quarantine`                                           | 9         | H5 (`scanning`, `clean`, `quarantined`, `blocked`, `execNote`, `preview`, `saveOpen`, `confirmTitle`, `confirmMsg`) |
| `errors`                                               | 4         | sync/conflict/resync notices                                                                                        |
| `sync`                                                 | 3         | resync wording                                                                                                      |
| `close` / `btn` / `offline`                            | 3 / 3 / 2 | tab-close guard, buttons, offline indicator                                                                         |
| `tabTitle`, `openMessage`, `noKeyringBanner`, `newTab` | 1 each    | shell + H2 session-only banner                                                                                      |

- Hardening additions since P4 (all `zh`-first with identical `en` in the
  same lane PR): `mail.quarantine.*` (9 keys, H5),
  `mail.noKeyringBanner` + `mail.accounts.sessionOnly` (H2),
  `mail.ai.*` growth to 247 (P5 + H3 guardrail copy + H4).
- Deliberately NOT user strings: the H4 per-account AI rate limit surfaces
  the PII-free error `mail: rate-limited, retry in <n>s` from main — no
  strings.ts key by design (error text carries timing only, never content).

## Pre-merge gate

1. `mail-strings.test.ts` green: every locale has exactly the `zh` key set,
   no empty values, frozen inventory keys present.
2. No new key without a `zh` value plus identical `en` value in the same PR.
3. Reused existing keys where the meaning matches (cap, outbox, error keys)
   instead of minting near-duplicates.
4. English-only comments and strings in the diff.
5. The lane reports the final key list in its lane metadata; docs updates the
   inventory table above if keys were added.
