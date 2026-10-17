/**
 * Mail redaction helper for sync/outbox log and error paths (P4 spec
 * section 9). The plan named `agent-core/redact-core.ts` as the extension
 * point, but no such module exists in the tree (verified: no redact-core
 * anywhere outside the CLI argv helper), so the Y lane carries this local
 * helper in sync scope. Every new log/error call-site in sync/* passes
 * through sanitizeError; bodies, snippets, addresses, and secrets never
 * reach lastError, MailEvents, or logs.
 */

const EMAIL_RE = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g
const TOKEN_RE = /(bearer\s+[A-Za-z0-9\-._~+/]+|ya29\.[A-Za-z0-9\-_]+|xox[a-z]-[A-Za-z0-9-]+)/gi
const LONG_B64_RE = /[A-Za-z0-9+/=_-]{64,}/g
const IPV4_RE = /\b(?:\d{1,3}\.){3}\d{1,3}\b/g
const IPV6_RE = /(?:\b(?:[0-9A-Fa-f]{1,4}:){3,}[0-9A-Fa-f:.]*[0-9A-Fa-f]\b|::[0-9A-Fa-f:.]{1,})/g
const KEY_MATERIAL_RE =
  /-----BEGIN [A-Z0-9 ]*?(?:PRIVATE KEY|KEY|CERTIFICATE)-----[\s\S]*?-----END [A-Z0-9 ]*?(?:PRIVATE KEY|KEY|CERTIFICATE)-----/g
const PASSWORD_RE = /(password|passwd|pwd)\s*[:=]\s*("[^"]*"|'[^']*'|[^\s",;}]+)/gi
const RECEIVED_LINE_RE = /^received:.*$/gim
const AUTH_RESULTS_LINE_RE = /^authentication-results:.*$/gim
const DKIM_SIG_LINE_RE = /^dkim-signature:.*$/gim
const RETURN_PATH_LINE_RE = /^return-path:.*$/gim
const X_RECEIVED_LINE_RE = /^x-received:.*$/gim

/** Cap for prompt-bound text: retrieval context never grows a prompt without bound. */
export const PROMPT_TEXT_CAP_CHARS = 200_000

/** Provider error text reduced to a PII-free one-line summary. */
export function sanitizeError(err: unknown): string {
  const raw = err instanceof Error ? err.message : String(err ?? 'sync failed')
  const oneLine = raw.split('\n')[0].slice(0, 280).trim()
  return (
    oneLine
      .replace(EMAIL_RE, '[address]')
      .replace(TOKEN_RE, '[secret]')
      .replace(LONG_B64_RE, '[blob]')
      .replace(KEY_MATERIAL_RE, '[redacted-key]')
      .replace(PASSWORD_RE, '$1 [redacted]')
      .replace(IPV4_RE, '[ip]')
      .replace(IPV6_RE, '[ip]')
      .trim() || 'sync failed'
  )
}

/**
 * Drop the transport header chain before any header block enters a prompt
 * (threat T9: metadata leakage). Removes Received, Authentication-Results,
 * DKIM-Signature, Return-Path, and X-Received lines plus their folded
 * continuation lines; every other header and the body pass through.
 */
export function stripTransportHeaders(raw: string): string {
  const dropStart = /^(received|authentication-results|dkim-signature|return-path|x-received):/i
  const kept: string[] = []
  let dropping = false
  for (const line of raw.split('\n')) {
    if (/^[ \t]/.test(line) && dropping) continue
    dropping = dropStart.test(line)
    if (!dropping) kept.push(line)
  }
  return kept.join('\n')
}

/**
 * Redact free text before it crosses into a log line or a model prompt.
 * Applies the sanitizeError rules plus transport-header removal, so
 * tokens, passwords, Received/Authentication-Results internals, IPs, and
 * key material never reach logs, telemetry, or provider calls. Output is
 * capped so retrieval context cannot grow a prompt without bound.
 */
export function redactForPrompt(text: string): string {
  const stripped = stripTransportHeaders(text)
    .replace(RECEIVED_LINE_RE, '[redacted]')
    .replace(AUTH_RESULTS_LINE_RE, '[redacted]')
    .replace(DKIM_SIG_LINE_RE, '[redacted]')
    .replace(RETURN_PATH_LINE_RE, '[redacted]')
    .replace(X_RECEIVED_LINE_RE, '[redacted]')
    .replace(KEY_MATERIAL_RE, '[redacted-key]')
    .replace(PASSWORD_RE, '$1 [redacted]')
    .replace(EMAIL_RE, '[address]')
    .replace(TOKEN_RE, '[secret]')
    .replace(IPV4_RE, '[ip]')
    .replace(IPV6_RE, '[ip]')
    .replace(LONG_B64_RE, '[blob]')
  if (stripped.length <= PROMPT_TEXT_CAP_CHARS) return stripped
  return `${stripped.slice(0, PROMPT_TEXT_CAP_CHARS)}[truncated]`
}

/**
 * True when text looks free of PII/secret bytes (test assertion helper).
 * Flags addresses, token shapes, key material, IP literals, and long
 * base64 runs that could carry body or credential bytes.
 */
export function containsPii(text: string): boolean {
  // Global regexes keep lastIndex state across .test calls; reset first
  // so repeated calls give stable answers.
  EMAIL_RE.lastIndex = 0
  TOKEN_RE.lastIndex = 0
  LONG_B64_RE.lastIndex = 0
  IPV4_RE.lastIndex = 0
  IPV6_RE.lastIndex = 0
  KEY_MATERIAL_RE.lastIndex = 0
  PASSWORD_RE.lastIndex = 0
  return (
    EMAIL_RE.test(text) ||
    TOKEN_RE.test(text) ||
    LONG_B64_RE.test(text) ||
    IPV4_RE.test(text) ||
    IPV6_RE.test(text) ||
    KEY_MATERIAL_RE.test(text) ||
    PASSWORD_RE.test(text)
  )
}
