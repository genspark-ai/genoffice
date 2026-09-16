/** Process exit codes; the JSON error payload carries the same number. */
export const EXIT = {
  ok: 0,
  usage: 1,
  file: 2,
  conversion: 3,
  app: 4,
} as const

export type ExitCode = (typeof EXIT)[keyof typeof EXIT]

/**
 * Machine-readable reason, finer than the exit code: a program branches on
 * `error`, a person reads `message`. Every reason is also documented in the
 * skill, so add sparingly and keep the names stable.
 */
export type ErrorReason =
  | 'unknown_command'
  | 'unknown_option'
  | 'missing_argument'
  | 'invalid_argument'
  | 'invalid_json'
  | 'unsupported'
  | 'unknown_op'
  | 'op_rejected'
  | 'target_not_found'
  | 'out_of_range'
  | 'sheet_not_found'
  | 'file_not_found'
  | 'resource_limit'
  | 'output_exists'
  | 'file_open_in_gui'
  | 'file_not_open_in_gui'
  | 'outside_allowed_roots'
  | 'conversion_failed'
  | 'app_unavailable'
  | 'invalid_usage'

export interface ErrorHints {
  reason?: ErrorReason
  /** one sentence on what to do next, imperative, no trailing period */
  suggestion?: string
}

const DEFAULT_REASON: Record<ExitCode, ErrorReason> = {
  [EXIT.ok]: 'invalid_usage',
  [EXIT.usage]: 'invalid_usage',
  [EXIT.file]: 'file_not_found',
  [EXIT.conversion]: 'conversion_failed',
  [EXIT.app]: 'app_unavailable',
}

export class CliError extends Error {
  readonly reason: ErrorReason
  readonly suggestion?: string

  constructor(
    readonly code: ExitCode,
    message: string,
    readonly detail?: Record<string, unknown>,
    hints: ErrorHints = {},
  ) {
    super(message)
    this.name = 'CliError'
    this.reason = hints.reason ?? DEFAULT_REASON[code]
    if (hints.suggestion) this.suggestion = hints.suggestion
  }
}

/** Something a program should know about a result that still succeeded. */
export interface Warning {
  code: string
  message: string
  suggestion?: string
}

export interface CommandResult {
  /** `partial`: the file was written but a batch lost ops (see detail.batch / detail.failures) */
  status?: 'ok' | 'partial'
  summary: string
  outputPath?: string
  detail?: Record<string, unknown>
  warnings?: Warning[]
}

export interface JsonOk {
  status: 'ok' | 'partial'
  command: string
  summary: string
  output_path?: string
  warnings?: Warning[]
  detail?: Record<string, unknown>
}

export interface JsonError {
  status: 'error'
  command: string | null
  code: ExitCode
  error: ErrorReason
  message: string
  suggestion?: string
  detail?: Record<string, unknown>
}

export function toJsonOk(command: string, r: CommandResult): JsonOk {
  return {
    status: r.status ?? 'ok',
    command,
    summary: r.summary,
    ...(r.outputPath ? { output_path: r.outputPath } : {}),
    ...(r.warnings?.length ? { warnings: r.warnings } : {}),
    ...(r.detail ? { detail: r.detail } : {}),
  }
}

export function toJsonError(command: string | null, err: CliError): JsonError {
  return {
    status: 'error',
    command,
    code: err.code,
    error: err.reason,
    message: err.message,
    ...(err.suggestion ? { suggestion: err.suggestion } : {}),
    ...(err.detail ? { detail: err.detail } : {}),
  }
}

export function formatHuman(r: CommandResult): string {
  const lines = [r.summary]
  if (r.outputPath) lines.push(`  output: ${r.outputPath}`)
  for (const w of r.warnings ?? []) {
    lines.push(`  warning: ${w.message}${w.suggestion ? ` (${w.suggestion})` : ''}`)
  }
  if (r.detail) {
    for (const [key, value] of Object.entries(r.detail)) {
      if (key === 'batch') {
        lines.push(batchLine(value))
        continue
      }
      const text =
        value !== null && typeof value === 'object' ? JSON.stringify(value) : String(value)
      lines.push(`  ${key}: ${text}`)
    }
  }
  return lines.join('\n')
}

export function formatHumanError(err: CliError): string {
  const lines = [`genoffice: ${err.message}`]
  if (err.suggestion) lines.push(`  hint: ${err.suggestion}`)
  if (err.detail) {
    const { batch, ...rest } = err.detail
    if (batch !== undefined) lines.push(batchLine(batch))
    if (Object.keys(rest).length) lines.push(`  ${JSON.stringify(rest)}`)
  }
  return lines.join('\n')
}

function batchLine(value: unknown): string {
  const b = value as { total: number; applied: number; failed: number; skipped: number }
  return `  batch: ${b.applied} of ${b.total} ops applied, ${b.failed} failed, ${b.skipped} skipped`
}
