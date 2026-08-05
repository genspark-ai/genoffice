/**
 * Convert provider failures into diagnostics safe for renderer IPC. Never
 * return a raw response body, URL, or error message: those may contain account
 * or request data. A provider-sanitized 400 code/detail is the sole exception.
 */
export function safeCodexStreamError(error: unknown): string {
  const message = error instanceof Error ? error.message : ''
  const record = error && typeof error === 'object' ? (error as Record<string, unknown>) : undefined
  const status =
    typeof record?.status === 'number'
      ? String(record.status)
      : /^Codex HTTP (\d{3})(?:\D|$)/.exec(message)?.[1]
  if (status) {
    switch (Number(status)) {
      case 401:
        return 'ChatGPT authorization was rejected (HTTP 401). Sign in again.'
      case 403:
        return 'ChatGPT access was denied (HTTP 403). Check your ChatGPT access.'
      case 404:
        return 'ChatGPT Codex endpoint or selected model is unavailable (HTTP 404).'
      case 429:
        return 'ChatGPT Codex rate limit reached (HTTP 429). Try again later.'
      default:
        if (Number(status) === 400) {
          const code = safeDiagnosticCode(record?.code)
          const detail = safeDiagnosticDetail(record?.detail)
          const diagnostic = [code, detail].filter(Boolean).join(': ')
          return diagnostic
            ? `ChatGPT Codex rejected the request (HTTP 400; ${diagnostic}).`
            : 'ChatGPT Codex rejected the request (HTTP 400).'
        }
        if (Number(status) >= 500) {
          return `ChatGPT Codex service error (HTTP ${status}). Try again later.`
        }
        return `ChatGPT Codex request failed (HTTP ${status}). Try again.`
    }
  }
  if (/timed? out|timeout/i.test(message)) {
    return 'ChatGPT Codex request timed out. Try again.'
  }
  if (/fetch failed|network|enotfound|econnrefused|econnreset/i.test(message)) {
    return 'ChatGPT Codex network error. Check your connection and try again.'
  }
  if (/stream malformed|invalid stream/i.test(message)) {
    return 'ChatGPT Codex returned an invalid stream response. Try again.'
  }
  if (/model.*(?:not found|not available|unsupported)|unsupported.*model/i.test(message)) {
    return 'ChatGPT Codex selected model is unavailable. Choose another model and try again.'
  }
  if (/tool arguments malformed/i.test(message)) {
    return 'ChatGPT Codex returned an invalid tool call. Try again.'
  }
  return 'ChatGPT Codex request failed. Try again.'
}

function safeDiagnosticCode(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined
  const code = value.trim()
  return /^[a-z][a-z0-9_.-]{0,63}$/i.test(code) ? code : undefined
}

function safeDiagnosticDetail(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined
  const detail = value.replace(/\s+/g, ' ').trim()
  if (
    !detail ||
    /https?:\/\/|\b(?:authorization|bearer|token|api[_ -]?key|password|secret|credential)\b|\b(?:sk-|sess-|eyJ)[a-z0-9_-]{8,}/i.test(
      detail,
    )
  ) {
    return undefined
  }
  return detail.slice(0, 160)
}
