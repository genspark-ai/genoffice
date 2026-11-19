import { appendFileSync, mkdirSync, renameSync, statSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, join } from 'node:path'

/**
 * One JSON line per executed command, so the person who let an agent loose
 * with genoffice can see afterwards what it touched. Default location
 * ~/.genoffice/cli-audit.jsonl; GENOFFICE_AUDIT_LOG=<path> redirects, `off`
 * disables. Never throws: a failed audit write must not fail the command.
 */
export interface AuditRecord {
  command: string
  argv: readonly string[]
  status: 'ok' | 'error'
  code: number
  output_path?: string
  ms: number
  cwd: string
}

const MAX_BYTES = 2_000_000

export function auditLogPath(env: NodeJS.ProcessEnv): string | null {
  const raw = env.GENOFFICE_AUDIT_LOG?.trim()
  if (raw?.toLowerCase() === 'off') return null
  if (raw) return raw
  return join(env.GENOFFICE_AUTH_DIR || join(homedir(), '.genoffice'), 'cli-audit.jsonl')
}

export function appendAudit(env: NodeJS.ProcessEnv, record: AuditRecord): void {
  const path = auditLogPath(env)
  if (!path) return
  try {
    mkdirSync(dirname(path), { recursive: true, mode: 0o700 })
    try {
      if (statSync(path).size > MAX_BYTES) renameSync(path, `${path}.1`)
    } catch {}
    const line = JSON.stringify({ ts: new Date().toISOString(), pid: process.pid, ...record })
    appendFileSync(path, line + '\n', { mode: 0o600 })
  } catch {}
}
