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

This is **partial coverage**, not a migration of the AI provider. GSK/BYOK traffic
requires a separate token-subject issuer and matching edge contract;
arbitrary providers must never receive this pass. No edge activation or
rollout is included.

Admitted hosts come from configuration, not source: the `GSK_BASE_URL` host
(default `www.genspark.ai`) plus any comma-separated hostnames in
`GENOFFICE_NATIVE_WAF_HOSTS`. Requests to any other origin are sent unchanged.
