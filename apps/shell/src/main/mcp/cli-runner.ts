import { spawn } from 'node:child_process'

/**
 * Shell-main half of the MCP → `genoffice` CLI delegation.
 *
 * The app already ships the CLI (`@genoffice/cli`) and runs it on its own Node
 * runtime, so the headless MCP tools do not reimplement document engines: they
 * spawn the same CLI a shell user or another agent would, with `--json`, and
 * map its result contract (`{ status, command, summary, output_path, detail }`,
 * exit codes 0-4 — see packages/cli/src/result.ts) back to a tool result.
 *
 * Spawning (not importing) is deliberate: the CLI entry bundles jsdom/jszip/
 * undici and the whole engine surface, none of which belongs inside the Electron
 * main bundle. The runner is injectable so unit tests use a fake instead of a
 * process.
 */

/** the CLI's `--json` success payload */
export interface CliJsonOk {
  status: 'ok'
  command: string
  summary: string
  output_path?: string
  detail?: Record<string, unknown>
}

/** the CLI's `--json` error payload; `code` mirrors the process exit code */
export interface CliJsonError {
  status: 'error'
  command: string | null
  code: number
  message: string
  detail?: Record<string, unknown>
}

export interface CliRunOutcome {
  ok: boolean
  /** process exit code, or -1 when the process never started/produced output */
  code: number
  json?: CliJsonOk | CliJsonError
  stdout: string
  stderr: string
}

export interface CliRunOptions {
  /** text piped to stdin (for commands that read `-`, e.g. `create --ops -`) */
  input?: string
  /** abort the child after this long; default 120s (the CLI lazily loads jsdom) */
  timeoutMs?: number
}

export interface CliRunner {
  /** run `genoffice <args> --json`; never rejects — failures come back as `ok:false` */
  run(args: string[], options?: CliRunOptions): Promise<CliRunOutcome>
}

export interface CliRunnerPaths {
  /** node/electron executable that runs the CLI entry */
  executable: string
  /** absolute path to the bundled CLI entry (dist/genoffice.cjs) */
  entry: string
  /** extra environment (the runner adds ELECTRON_RUN_AS_NODE when needed) */
  env?: NodeJS.ProcessEnv
}

function parseJson(stdout: string): CliJsonOk | CliJsonError | undefined {
  const text = stdout.trim()
  if (!text) return undefined
  // the CLI prints exactly one JSON object with --json; take the last line so a
  // stray warning on stdout cannot mask it
  const lastLine = text.split('\n').filter(Boolean).at(-1) ?? text
  try {
    const parsed = JSON.parse(lastLine) as CliJsonOk | CliJsonError
    return parsed && typeof parsed === 'object' && 'status' in parsed ? parsed : undefined
  } catch {
    return undefined
  }
}

/** Child output budget: past this the buffers are truncated and the child killed. */
export const MAX_CLI_OUTPUT_BYTES = 64 * 1024 * 1024

export function createCliRunner(paths: CliRunnerPaths): CliRunner {
  return {
    run(args, options = {}) {
      // timeoutMs arrives from tool callers: a NaN/Infinity value would break
      // the watchdog (NaN fires immediately, Infinity never), so normalize it.
      const timeoutMs = Number.isFinite(options.timeoutMs)
        ? Math.min(Math.max(1000, Math.floor(options.timeoutMs as number)), 600_000)
        : 120_000
      return new Promise<CliRunOutcome>((resolve) => {
        const child = spawn(paths.executable, [paths.entry, ...args, '--json'], {
          env: { ...process.env, ELECTRON_RUN_AS_NODE: '1', ...paths.env },
          windowsHide: true,
        })
        let stdout = ''
        let stderr = ''
        let settled = false
        let truncated = false
        const finish = (outcome: CliRunOutcome): void => {
          if (settled) return
          settled = true
          clearTimeout(timer)
          resolve(outcome)
        }
        // Unbounded child output would OOM the shell main process: cap the
        // buffers and kill the child past the budget.
        const onChunk = (chunk: Buffer, stream: 'out' | 'err'): void => {
          if (stream === 'out') stdout += chunk.toString()
          else stderr += chunk.toString()
          if (stdout.length + stderr.length > MAX_CLI_OUTPUT_BYTES && !truncated) {
            truncated = true
            child.kill()
            stdout = stdout.slice(0, MAX_CLI_OUTPUT_BYTES)
            stderr = `${stderr.slice(0, 1000)}\n…output truncated past ${MAX_CLI_OUTPUT_BYTES} bytes`
            finish({ ok: false, code: -1, stdout, stderr })
          }
        }
        const timer = setTimeout(() => {
          child.kill()
          finish({ ok: false, code: -1, stdout, stderr: `${stderr}\ngenoffice timed out` })
        }, timeoutMs)
        child.stdout?.on('data', (chunk: Buffer) => onChunk(chunk, 'out'))
        child.stderr?.on('data', (chunk: Buffer) => onChunk(chunk, 'err'))
        child.on('error', (err) => {
          finish({ ok: false, code: -1, stdout, stderr: `${stderr}\n${String(err)}` })
        })
        child.on('close', (code) => {
          const json = parseJson(stdout)
          finish({
            ok: code === 0 && json?.status === 'ok',
            code: code ?? -1,
            json,
            stdout,
            stderr,
          })
        })
        if (options.input !== undefined) child.stdin?.end(options.input)
        else child.stdin?.end()
      })
    },
  }
}

/** turn a failed CLI run into a tool error message an agent can act on */
export function cliErrorMessage(outcome: CliRunOutcome): string {
  if (outcome.json && outcome.json.status === 'error') return outcome.json.message
  const detail = outcome.stderr.trim() || outcome.stdout.trim()
  return detail ? `genoffice failed: ${detail}` : `genoffice exited with code ${outcome.code}`
}
