import { existsSync, readFileSync } from 'node:fs'
import { connect } from 'node:net'
import { join } from 'node:path'
import {
  CONTROL_FILE,
  CONTROL_MAX_REQUEST_BYTES,
  type ControlEndpoint,
  type ControlError,
  type ControlReply,
  type ControlRequest,
} from './control-protocol'
import { genofficeUserDataDir } from './gui'
import { CliError, EXIT, type ErrorReason } from './result'

/** The running shell's control endpoint, or null when no live shell published one. */
export function controlEndpoint(env: NodeJS.ProcessEnv): ControlEndpoint | null {
  const dirs = env.GENOFFICE_USER_DATA
    ? [env.GENOFFICE_USER_DATA]
    : [genofficeUserDataDir(env), `${genofficeUserDataDir(env)} Dev`]
  for (const dir of dirs) {
    const file = join(dir, CONTROL_FILE)
    if (!existsSync(file)) continue
    try {
      const raw = JSON.parse(readFileSync(file, 'utf8')) as Partial<ControlEndpoint>
      if (
        typeof raw.pid !== 'number' ||
        typeof raw.endpoint !== 'string' ||
        typeof raw.token !== 'string'
      )
        continue
      if (!processAlive(raw.pid)) continue
      return raw as ControlEndpoint
    } catch {
      continue
    }
  }
  return null
}

function processAlive(pid: number): boolean {
  try {
    process.kill(pid, 0)
    return true
  } catch (err) {
    return (err as NodeJS.ErrnoException).code === 'EPERM'
  }
}

export async function waitForControlEndpoint(
  env: NodeJS.ProcessEnv,
  timeoutMs: number,
): Promise<ControlEndpoint | null> {
  const deadline = Date.now() + timeoutMs
  for (;;) {
    const found = controlEndpoint(env)
    if (found || Date.now() > deadline) return found
    await new Promise((r) => setTimeout(r, 250))
  }
}

/** One request over a fresh connection; the reply resolves or a CliError explains why not. */
export function controlRequest(
  endpoint: ControlEndpoint,
  request: ControlRequest,
  timeoutMs = 30_000,
): Promise<Record<string, unknown>> {
  return new Promise((resolve, reject) => {
    const socket = connect(endpoint.endpoint)
    let buffer = ''
    let settled = false
    const settle = (fn: () => void) => {
      if (settled) return
      settled = true
      socket.destroy()
      fn()
    }
    const unavailable = (why: string) =>
      settle(() =>
        reject(
          new CliError(EXIT.app, `GenOffice did not answer: ${why}`, undefined, {
            reason: 'app_unavailable',
            suggestion: 'start GenOffice (or `genoffice open <file>`) and retry',
          }),
        ),
      )
    socket.setEncoding('utf8')
    socket.setTimeout(timeoutMs, () => unavailable('timed out'))
    socket.on('error', (err) => unavailable(err.message))
    socket.on('connect', () => {
      socket.write(JSON.stringify({ token: endpoint.token, request }) + '\n')
    })
    socket.on('data', (chunk: string) => {
      buffer += chunk
      if (buffer.length > CONTROL_MAX_REQUEST_BYTES * 4) return unavailable('reply too large')
      const nl = buffer.indexOf('\n')
      if (nl < 0) return
      let reply: ControlReply
      try {
        reply = JSON.parse(buffer.slice(0, nl)) as ControlReply
      } catch {
        return unavailable('malformed reply')
      }
      settle(() => (reply.ok ? resolve(reply.result) : reject(controlError(reply.error))))
    })
    socket.on('close', () => unavailable('connection closed'))
  })
}

const EXIT_FOR_REASON: Partial<Record<ErrorReason, (typeof EXIT)[keyof typeof EXIT]>> = {
  file_not_found: EXIT.file,
  file_not_open_in_gui: EXIT.file,
  app_unavailable: EXIT.app,
}

function controlError(error: ControlError): CliError {
  const reason = error.reason as ErrorReason
  const { suggestion, ...detail } = error.detail ?? {}
  return new CliError(
    EXIT_FOR_REASON[reason] ?? EXIT.usage,
    error.message,
    Object.keys(detail).length ? detail : undefined,
    {
      reason,
      suggestion: typeof suggestion === 'string' ? suggestion : defaultSuggestion(reason, detail),
    },
  )
}

function defaultSuggestion(reason: ErrorReason, detail: Record<string, unknown>): string {
  switch (reason) {
    case 'out_of_range':
      return detail.valid_range ? `use a value in ${detail.valid_range}` : 'read the file first'
    case 'target_not_found':
      return 'run `slides read --json` and pick an id from `available`'
    case 'sheet_not_found':
      return 'use one of `detail.sheets` verbatim'
    case 'file_not_open_in_gui':
      return 'run `genoffice open <file>` first'
    default:
      return 'check the target and retry'
  }
}
