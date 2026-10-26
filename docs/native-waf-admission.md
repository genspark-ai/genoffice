# Optional native admission

`GENOFFICE_NATIVE_WAF=1` opts the auth module's explicit-cookie API-key create/revoke
requests into cookie-v1 admission. Unset is the shipped default: no issuer request,
wait, timer or transport change. Device-code polling and the Bearer session exchange
stay unchanged, as do all provider/agent/subprocess requests.

The helper preserves legacy headers, binds one exact HTTPS origin and session
cookie, coalesces requests and fences logout/reload by generation. Issuer
reads are bounded to 5s/4096 bytes, with memory-only 240s positive (bounded by token
expiry) and 15s negative cache. Caller cancellation still aborts the wait. Failures
send the original business request without a pass; no new 403 replay is added.
Opt-in cookie traffic uses explicit Cookie with credentials omitted and manual
redirects, retaining a 302 response without leaking credentials across origins.

## Genspark provider requests (gsk key)

The same switch also opts the AI provider's Genspark LLM proxy requests into
token admission (native-tok-v1, family `llm`) at the shared `aiFetch` seam. Only
HTTPS requests to an admitted host (the default proxy host, the `GSK_BASE_URL` host and any
hosts in `GENOFFICE_NATIVE_WAF_HOSTS`) on `/api/llm_proxy/` or `/api/anthropic` that carry
exactly one gsk credential (`Authorization: Bearer gsk-…` or `x-api-key: gsk-…`)
qualify. The issuer call binds that exact header; bootstrap is bounded to 2s and
4096 bytes with no redirects; the memory cache is at most 240s and never past
the signed lifetime, with a 15s negative cache. Requests carrying a pass refuse
redirects. BYOK keys, custom endpoints and other hosts are never sent to the
issuer, and issuer failure sends the request unchanged. Renderers have no
process env and never opt in. The `gsk` CLI subprocess tools follow the CLI's
own `GSK_NATIVE_WAF` switch (CLI ≥ 1.15.1), not this one.

No edge activation or rollout is included.

Admitted hosts come from configuration, not source: the `GSK_BASE_URL` host
(default `www.genspark.ai`) plus any comma-separated hostnames in
`GENOFFICE_NATIVE_WAF_HOSTS`. Requests to any other origin are sent unchanged.
