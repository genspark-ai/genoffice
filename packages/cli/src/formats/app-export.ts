import { spawn as nodeSpawn, spawnSync, type ChildProcess } from 'node:child_process'
import { existsSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { appLaunch } from '../resources'
import { CliError, EXIT } from '../result'

/**
 * Conversions that need an app renderer (Word, Excel and PowerPoint layout,
 * HTML and Markdown print, Word ↔ HTML) run in the GenOffice binary itself through its
 * `--headless-export` entry: Dock hidden, no window, one export, exit. genoffice
 * just spawns it and reads the JSON envelope it prints. The app skips the
 * single-instance lock in that mode, so a running GUI does not interfere.
 *
 * On macOS a crash before any envelope is retried once through LaunchServices
 * (`open`): when genoffice itself runs under a command sandbox (Codex and
 * friends wrap every command in a deny-default Seatbelt profile), the spawned
 * app inherits it and AppKit aborts in _RegisterApplication during
 * `+[NSApplication sharedApplication]`, before any JS can run. A process
 * launchd starts is outside the caller's sandbox, so the relaunch works where
 * the direct spawn cannot. `open` has no stdout pipe back, so the retry asks
 * the app for a `--json-file` envelope instead.
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
  /** test seam: process probes (pgrep/pkill) for the LaunchServices relaunch */
  spawnSync?: typeof spawnSync
  /** test seam */
  platform?: NodeJS.Platform
  /** test seam: how often the LaunchServices relaunch polls for its envelope file */
  lsPollMs?: number
  /** test seam: how long the relaunched app gets to start before a silent death is reported */
  lsGraceMs?: number
}

export interface AppExportResult {
  outputPath: string
  summary: string
}

const DEFAULT_TIMEOUT_MS = 180_000
const STDIO_DRAIN_MS = 500
const LS_OPEN_TIMEOUT_MS = 15_000

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
  const { code, signal, stdout, stderr, timedOut } = await waitFor(
    child,
    opts.timeoutMs ?? DEFAULT_TIMEOUT_MS,
    opts.killGraceMs ?? 2000,
  )
  const envelope = parseEnvelope(stdout)
  // The envelope is printed after the file is written and every teardown the
  // export owns is done; a crash or kill during Electron's own quit must not
  // turn a finished export into a failure.
  if (envelope?.status === 'ok' && existsSync(outputPath)) {
    if (timedOut) opts.log?.('export finished but GenOffice had to be terminated on quit')
    else if (code !== 0)
      opts.log?.(`export finished but GenOffice ${describeExit(code, signal)} while quitting`)
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
  // A startup crash leaves no envelope to explain it; on macOS the usual cause
  // is the caller's own command sandbox, which a LaunchServices relaunch escapes.
  if (signal && !envelope && (opts.platform ?? process.platform) === 'darwin') {
    const relaunched = await exportViaLaunchServices(launch.command, args, outputPath, opts)
    if (relaunched.kind === 'success') return { outputPath, summary: relaunched.summary }
    if (relaunched.kind === 'error') throw relaunched.error
    // skipped / failed: nothing better to report than the original crash
  }
  const tail = stderr.trim().split('\n').filter(Boolean).slice(-3).join(' ')
  if (!envelope && (signal || code === null)) {
    throw new CliError(
      EXIT.app,
      `GenOffice ${describeExit(code, signal)} while exporting ${input}${tail ? `: ${tail}` : ''}`,
      { app: launch.command, exit_code: code, signal },
      {
        reason: 'app_crashed',
        suggestion:
          'retry once; if it crashes again, report it with the document and the crash log (macOS: ~/Library/Logs/DiagnosticReports)',
      },
    )
  }
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

type LaunchServicesRelaunch =
  | { kind: 'success'; summary: string }
  | { kind: 'error'; error: CliError }
  | { kind: 'skipped' }
  | { kind: 'failed' }

/**
 * Re-runs the export through `open`, which asks LaunchServices to start the
 * app outside this process's sandbox, and polls for the envelope file the app
 * writes on its way out. Apps older than `--json-file` never write one, so a
 * finished output file with the app gone is accepted as success too.
 */
async function exportViaLaunchServices(
  command: string,
  args: string[],
  outputPath: string,
  opts: AppExportOptions,
): Promise<LaunchServicesRelaunch> {
  const bundle = appBundleOf(command)
  if (!bundle) return { kind: 'skipped' }
  const spawn = opts.spawn ?? nodeSpawn
  const probe = opts.spawnSync ?? spawnSync
  const pollMs = opts.lsPollMs ?? 250
  const graceMs = opts.lsGraceMs ?? 3000
  const jsonFile = join(tmpdir(), `genoffice-export-${process.pid}-${Date.now()}.json`)
  opts.log?.('direct launch crashed; retrying through LaunchServices')
  const opened = await waitFor(
    spawn('open', ['-n', '-a', bundle, '--args', ...args, '--json-file', jsonFile], {
      stdio: ['ignore', 'pipe', 'pipe'],
    }),
    LS_OPEN_TIMEOUT_MS,
    2000,
  )
  if (opened.code !== 0) {
    const reason = opened.stderr.trim().split('\n').filter(Boolean).pop() ?? `open exited`
    opts.log?.(`LaunchServices relaunch unavailable (${reason})`)
    return { kind: 'failed' }
  }
  const startedAt = Date.now()
  const deadline = startedAt + (opts.timeoutMs ?? DEFAULT_TIMEOUT_MS)
  for (;;) {
    await new Promise((resolve) => setTimeout(resolve, pollMs))
    if (existsSync(jsonFile)) {
      const envelope = parseEnvelope(readOrEmpty(jsonFile))
      discard(jsonFile)
      if (envelope?.status === 'ok' && existsSync(outputPath)) {
        return { kind: 'success', summary: envelope.summary ?? `exported to ${outputPath}` }
      }
      if (envelope?.status === 'error') {
        return {
          kind: 'error',
          error: new CliError(
            exitCodeFor(envelope.exit_code ?? null),
            envelope.error ?? envelope.summary ?? 'export failed',
            { app: bundle, exit_code: envelope.exit_code },
          ),
        }
      }
      return { kind: 'failed' }
    }
    if (Date.now() >= deadline) {
      probe('pkill', ['-f', jsonFile])
      discard(jsonFile)
      return {
        kind: 'error',
        error: new CliError(
          EXIT.conversion,
          `GenOffice did not finish the export within ${Math.round((opts.timeoutMs ?? DEFAULT_TIMEOUT_MS) / 1000)}s`,
          { app: bundle },
        ),
      }
    }
    // the relaunched app vanished without an envelope: either it crashed again
    // or it predates --json-file and left only the output file
    if (Date.now() - startedAt > graceMs && !appRunning(jsonFile, probe)) {
      if (existsSync(outputPath)) return { kind: 'success', summary: `exported to ${outputPath}` }
      return { kind: 'failed' }
    }
  }
}

/** The .app bundle a binary lives in, or null when it is not a bundle app. */
function appBundleOf(command: string): string | null {
  const at = command.lastIndexOf('/Contents/MacOS/')
  if (at === -1) return null
  const bundle = command.slice(0, at)
  return bundle.endsWith('.app') ? bundle : null
}

function readOrEmpty(path: string): string {
  try {
    return readFileSync(path, 'utf8')
  } catch {
    return ''
  }
}

function discard(path: string): void {
  try {
    rmSync(path, { force: true })
  } catch {
    // a leftover temp envelope is harmless
  }
}

/** `pgrep -f` on the unique envelope path; a failed probe counts as running. */
function appRunning(marker: string, probe: typeof spawnSync): boolean {
  try {
    return probe('pgrep', ['-f', marker]).status === 0
  } catch {
    return true
  }
}

function describeExit(code: number | null, signal: NodeJS.Signals | null): string {
  return signal ? `crashed (${signal})` : `exited with code ${code}`
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
  /** present in --json-file envelopes, where no process exit can be read */
  exit_code?: number
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
): Promise<{
  code: number | null
  signal: NodeJS.Signals | null
  stdout: string
  stderr: string
  timedOut: boolean
}> {
  return new Promise((resolve) => {
    let stdout = ''
    let stderr = ''
    let settled = false
    let timedOut = false
    child.stdout?.on('data', (d) => (stdout += String(d)))
    child.stderr?.on('data', (d) => (stderr += String(d)))
    const finish = (code: number | null, signal: NodeJS.Signals | null = null) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      child.stdout?.destroy()
      child.stderr?.destroy()
      resolve({ code, signal, stdout, stderr, timedOut })
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
    child.once('close', (code, signal) => finish(code, signal))
    // `close` waits for every stdio handle to reach EOF, and Electron's renderer
    // and GPU helpers can hold the inherited pipes open after the main process
    // has printed the envelope and exited; settle shortly after `exit` instead.
    child.once('exit', (code, signal) => setTimeout(() => finish(code, signal), STDIO_DRAIN_MS))
  })
}
