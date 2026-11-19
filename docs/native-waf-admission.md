# Optional native admission

The auth module's explicit-cookie API-key create/revoke requests use cookie-v1
admission by default (gen-spark#72015). `GENOFFICE_NATIVE_WAF=0` is the rollback:
no issuer request, wait, timer or transport change. Device-code polling and the
Bearer session exchange stay unchanged, as do all provider/agent/subprocess requests.

The helper preserves legacy headers, binds one exact HTTPS origin and session
cookie, coalesces requests and fences logout/reload by generation. Issuer
reads are bounded to 5s/4096 bytes, with memory-only 240s positive (bounded by token
expiry) and 15s negative cache. Caller cancellation still aborts the wait. Failures
send the original business request without a pass; no new 403 replay is added.
Opt-in cookie traffic uses explicit Cookie with credentials omitted and manual
redirects, retaining a 302 response without leaking credentials across origins.

## Genspark provider requests (gsk key)

The provider lane is on by default in Node host processes (`GENOFFICE_NATIVE_WAF=0`
turns it off): the AI provider's Genspark LLM proxy requests use token admission (native-tok-v1, family `llm`) at the shared `aiFetch` seam. Only
HTTPS requests to an admitted host (the default proxy host, the `GSK_BASE_URL` host and any
hosts in `GENOFFICE_NATIVE_WAF_HOSTS`) on `/api/llm_proxy/` or `/api/anthropic` that carry
exactly one gsk credential (`Authorization: Bearer gsk-…` or `x-api-key: gsk-…`)
qualify. The issuer call binds that exact header; bootstrap is bounded to 2s and
4096 bytes with no redirects; the memory cache is at most 240s and never past
the signed lifetime, with a 15s negative cache. Requests carrying a pass refuse
redirects. BYOK keys, custom endpoints and other hosts are never sent to the
issuer, and issuer failure sends the request unchanged. Browsers and Electron
renderers never take part. The `gsk` CLI subprocess tools follow the CLI's
own `GSK_NATIVE_WAF` switch (CLI ≥ 1.15.1), not this one.

The prod edge token rules are gen-spark#77242; without them the pass is ignored
at the edge and legacy rules still apply.

Admitted hosts come from configuration, not source: the `GSK_BASE_URL` host
(default `www.genspark.ai`) plus any comma-separated hostnames in
`GENOFFICE_NATIVE_WAF_HOSTS`. Requests to any other origin are sent unchanged.
