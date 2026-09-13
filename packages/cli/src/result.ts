/** Process exit codes; the JSON error payload carries the same number. */
export const EXIT = {
  ok: 0,
  usage: 1,
  file: 2,
  conversion: 3,
  app: 4,
} as const

export type ExitCode = (typeof EXIT)[keyof typeof EXIT]

export class CliError extends Error {
  constructor(
    readonly code: ExitCode,
    message: string,
    readonly detail?: Record<string, unknown>,
  ) {
    super(message)
    this.name = 'CliError'
  }
}

export interface CommandResult {
  summary: string
  outputPath?: string
  detail?: Record<string, unknown>
}

export interface JsonOk {
  status: 'ok'
  command: string
  summary: string
  output_path?: string
  detail?: Record<string, unknown>
}

export interface JsonError {
  status: 'error'
  command: string | null
  code: ExitCode
  message: string
  detail?: Record<string, unknown>
}

export function toJsonOk(command: string, r: CommandResult): JsonOk {
  return {
    status: 'ok',
    command,
    summary: r.summary,
    ...(r.outputPath ? { output_path: r.outputPath } : {}),
    ...(r.detail ? { detail: r.detail } : {}),
  }
}

export function toJsonError(command: string | null, err: CliError): JsonError {
  return {
    status: 'error',
    command,
    code: err.code,
    message: err.message,
    ...(err.detail ? { detail: err.detail } : {}),
  }
}

export function formatHuman(r: CommandResult): string {
  const lines = [r.summary]
  if (r.outputPath) lines.push(`  output: ${r.outputPath}`)
  if (r.detail) {
    for (const [key, value] of Object.entries(r.detail)) {
      const text =
        value !== null && typeof value === 'object' ? JSON.stringify(value) : String(value)
      lines.push(`  ${key}: ${text}`)
    }
  }
  return lines.join('\n')
}
