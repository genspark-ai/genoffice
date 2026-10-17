# Mail OAuth app-verification paperwork (prepared text)

Checklist source: `apps/mail/docs/hardening-spec.md` section 5 (V-1..V-12,
H1 t_0672f5e2). Boundary: **no agent files submissions or touches
production consoles** — agents prepare the paper below, the owner clicks
submit. Ship order is unchanged: IMAP/SMTP with app passwords first (no
verification needed); Gmail/Graph API drivers stay behind the verification
landing until it is granted or a waiver is documented (V-12).

## Google (Gmail API, restricted scope `gmail.modify`)

- V-1 Brand pack: verified-domain homepage URL; privacy-policy URL carrying
  the disclosure draft below (mail-data access/use/storage/sharing,
  retention, deletion path); the in-product notice text below.
- V-2 Consent screen: app name, logo, homepage + privacy links; plan
  `External` user-type while unverified and acknowledge the ≤ 100 test-user
  cap.
- V-3 Scope justification (`gmail.modify`, restricted — why not read-only,
  why not full mail): "GenOffice Mail reads, sends, and manages the user's
  own mailbox (read, send, label/flag changes). Read-only scopes cannot send
  or update labels/flags, which is core function; full `mail` scope exceeds
  least privilege. `gmail.modify` is the minimum that covers read + send +
  label/flag write."
- V-4 Limited-use statement: "Mail data is used only to provide the mail
  features the user invokes. No advertising use, no sale, no human reading
  except consented support or security review, no model-training use. AI
  features call only the user's configured provider (own key or signed-in
  proxy)."
- V-5 CASA: confirm the assessment tier for restricted scopes, engage an
  assessor, and calendar the annual re-verification.
- V-6 Domain/redirects: domain ownership verified in Search Console;
  authorized domains and redirect URIs registered exactly as coded.

## Microsoft (Graph delegated `Mail.ReadWrite` + `offline_access` + `User.Read`)

- V-7 App registration: multitenant, delegated permissions ONLY — no
  `*.ReadWrite.All` application permissions. Redirect URIs
  (`http://localhost` loopback and/or `ms-appx-web`) match the code exactly.
- V-8 Publisher verification: MPN ID plus verified domain (removes the
  "unverified publisher" consent warning).
- V-9 Admin-consent path: designed as NORMAL — a tenant admin may still
  require approval; the app treats that as a handled state, not an error.
- V-10 Refresh-token story: `offline_access` yields the refresh token;
  rotation and wipe rules follow the credential-store policy (memory-only
  access tokens, wipe-and-reauth rotation).

## Shared

- V-11 Fork/rebrand rule: OAuth registrations belong to the DISTRIBUTION —
  a fork ships its own client IDs; upstream PRs keep the GenOffice marks.
  Client IDs are public identifiers, not secrets. Trademark boundary: the
  GenOffice and Genspark names and logos are trademarks of Mainfunc, Inc.;
  the Apache-2.0 license grants no right to use them (README License
  section) — forks rebrand, contributions keep the marks.
- V-12 Status tracking: verification status tracked per release; release
  builds require landing or a documented waiver.

## Prepared disclosure draft (privacy policy + in-product notice)

> GenOffice Mail accesses your mail accounts only to provide the mail
> features you invoke: reading, syncing, searching, composing, and sending,
> including the optional AI triage, summaries, compose assistance, and
> mail search. Message content is processed on your device; AI features
> send only the excerpt needed for your request to your configured AI
> provider (your own key, or the signed-in proxy). Credentials (refresh
> tokens, app passwords) are stored encrypted on your device; access tokens
> are kept in memory only and never written to disk. We do not sell mail
> data, use it for advertising, allow human reading except consented support
> or security review, or use it to train models. Retention: synced mail stays
> in your local store until you remove the account (which deletes its local
> cache) — deletion path: remove the account in Mail, then revoke the grant
> in your Google/Microsoft account settings.

In-product notice (shown before OAuth connect): "GenOffice Mail will access
your mailbox to sync, search, and send mail, including optional AI features
using your configured provider. Credentials stay encrypted on this device.
You can remove the account (and its local cache) at any time."
