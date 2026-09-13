import { spawn as nodeSpawn, type ChildProcess } from 'node:child_process'
import { existsSync } from 'node:fs'
import { appLaunch } from '../resources'
import { CliError, EXIT } from '../result'

/**
 * Conversions that need an app renderer (Word, Excel and PowerPoint layout,
 * HTML and Markdown print, Word ↔ HTML) run in the GenOffice binary itself through its
 * `--headless-export` entry: Dock hidden, no window, one export, exit. genoffice
 * just spawns it and reads the JSON envelope it prints. The app skips the
 * single-instance lock in that mode, so a running GUI does not interfere.
 */
export type AppExportTarget = 'pdf' | 'docx' | 'html'

export interface AppExportOptions {
  env?: NodeJS.ProcessEnv
  log?: (message: string) => void
  timeoutMs?: number
  /** how long a timed-out export gets to quit on SIGTERM before SIGKILL */
  killGraceMs?: number
  /** test seam */
  spawn?: typeof nodeSpawn
}

export interface AppExportResult {
  outputPath: string
  summary: string
}

const DEFAULT_TIMEOUT_MS = 180_000
const STDIO_DRAIN_MS = 500

export async function exportViaApp(
  input: string,
  target: AppExportTarget,
  outputPath: string,
  opts: AppExportOptions = {},
): Promise<AppExportResult> {
  const env = opts.env ?? process.env
  const launch = appLaunch(env)
  if (!launch) {
    throw new CliError(EXIT.app, 'GenOffice app not found (needed for this conversion)', {
      hint: 'install GenOffice, or set GENOFFICE_APP_BIN to its executable',
    })
  }
  const args = [
    ...launch.args,
    '--headless-export',
    input,
    '--to',
    target,
    '--out',
    outputPath,
    '--json',
  ]
  const childEnv = { ...env }
  delete childEnv.ELECTRON_RUN_AS_NODE
  opts.log?.(`starting GenOffice for ${target} export`)
  const spawn = opts.spawn ?? nodeSpawn
  const child = spawn(launch.command, args, { env: childEnv, stdio: ['ignore', 'pipe', 'pipe'] })
  const { code, stdout, stderr, timedOut } = await waitFor(
    child,
    opts.timeoutMs ?? DEFAULT_TIMEOUT_MS,
    opts.killGraceMs ?? 2000,
  )
  const envelope = parseEnvelope(stdout)
  const exported = envelope?.status === 'ok' && existsSync(outputPath)
  if (exported && (code === 0 || timedOut)) {
    if (timedOut) opts.log?.('export finished but GenOffice had to be terminated on quit')
    return { outputPath, summary: envelope.summary ?? `exported to ${outputPath}` }
  }
  if (timedOut) {
    throw new CliError(
      EXIT.conversion,
      `GenOffice did not finish the export within ${Math.round((opts.timeoutMs ?? DEFAULT_TIMEOUT_MS) / 1000)}s`,
      {
        app: launch.command,
      },
    )
  }
  const tail = stderr.trim().split('\n').filter(Boolean).slice(-3).join(' ')
  const message =
    envelope?.error ??
    envelope?.summary ??
    (tail ||
      (code === 0
        ? 'GenOffice exited without writing the file; this GenOffice version may not support --headless-export'
        : `GenOffice exited with code ${code}`))
  throw new CliError(exitCodeFor(code), message, {
    app: launch.command,
    exit_code: code,
  })
}

/** The app's HEADLESS_EXIT codes: 1 bad args, 2 input error, 3 conversion failure. */
function exitCodeFor(
  code: number | null,
): typeof EXIT.file | typeof EXIT.conversion | typeof EXIT.app {
  if (code === 2) return EXIT.file
  if (code === 3 || code === 1) return EXIT.conversion
  return EXIT.app
}

export interface HeadlessEnvelope {
  status: 'ok' | 'error'
  summary?: string
  output_path?: string
  error?: string
}

/** Last JSON line on stdout; the app may log other lines before it. */
export function parseEnvelope(stdout: string): HeadlessEnvelope | null {
  const lines = stdout
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean)
  for (let i = lines.length - 1; i >= 0; i--) {
    const line = lines[i]!
    if (!line.startsWith('{')) continue
    try {
      const parsed = JSON.parse(line) as HeadlessEnvelope
      if (parsed && (parsed.status === 'ok' || parsed.status === 'error')) return parsed
    } catch {}
  }
  return null
}

function waitFor(
  child: ChildProcess,
  timeoutMs: number,
  killGraceMs: number,
): Promise<{ code: number | null; stdout: string; stderr: string; timedOut: boolean }> {
  return new Promise((resolve) => {
    let stdout = ''
    let stderr = ''
    let settled = false
    let timedOut = false
    child.stdout?.on('data', (d) => (stdout += String(d)))
    child.stderr?.on('data', (d) => (stderr += String(d)))
    const finish = (code: number | null) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      child.stdout?.destroy()
      child.stderr?.destroy()
      resolve({ code, stdout, stderr, timedOut })
    }
    // On timeout: ask Electron to quit (it tears its children down), give it
    // a grace period, then kill; only report once the process is gone, so the
    // CLI never exits with a hung export still running.
    const timer = setTimeout(() => {
      timedOut = true
      child.kill('SIGTERM')
      setTimeout(() => {
        if (settled) return
        child.kill('SIGKILL')
        setTimeout(() => finish(null), 1000)
      }, killGraceMs)
    }, timeoutMs)
    child.once('error', (err) => {
      stderr += `\n${err.message}`
      finish(null)
    })
    child.once('close', (code) => finish(code))
    // `close` waits for every stdio handle to reach EOF, and Electron's renderer
    // and GPU helpers can hold the inherited pipes open after the main process
    // has printed the envelope and exited; settle shortly after `exit` instead.
    child.once('exit', (code) => setTimeout(() => finish(code), STDIO_DRAIN_MS))
  })
}
