import { spawnSync } from 'node:child_process'
import {
  accessSync,
  constants,
  existsSync,
  lstatSync,
  mkdirSync,
  readlinkSync,
  realpathSync,
  symlinkSync,
  unlinkSync,
} from 'node:fs'
import { basename, dirname, isAbsolute, join, resolve, win32 } from 'node:path'

/**
 * Making `genoffice` reachable from a terminal. The launcher ships inside the app
 * (Resources/cli/genoffice, or genoffice.cmd on Windows); nothing at install time puts
 * it on PATH for the dmg, so the app tries this on every launch until it
 * succeeds and `genoffice install-cli` repeats it on demand. Best effort
 * everywhere: no prompts, no elevation, never an error on the app's startup
 * path. `inspectCliLink` reports the same states without writing anything.
 */
export interface InstallOptions {
  /** absolute path of the launcher to expose */
  launcher: string
  platform?: NodeJS.Platform
  /** directories to try, first writable wins (defaults per platform) */
  candidateDirs?: string[]
  env?: NodeJS.ProcessEnv
  /** test seam for the Windows registry edit */
  runPowerShell?: (script: string) => { ok: boolean; stdout: string }
}

export type InstallStatus =
  'linked' | 'present' | 'missing' | 'unwritable' | 'occupied' | 'unsupported'

export interface InstallOutcome {
  status: InstallStatus
  /** where the link or PATH entry lives (or would live) */
  location?: string
  /** what a person can run to finish the job when the app could not */
  manual?: string
  /** the link landed in a directory the current shell does not search; this line adds it */
  pathHint?: string
}

/**
 * `/usr/local/bin` is on every login shell's PATH and on the PATH GUI apps
 * inherit, so it is the target even when it does not exist yet: a missing
 * directory is reported as unwritable (creating it needs root), not skipped.
 * `/opt/homebrew/bin` is only a fallback because shells see it solely through
 * `brew shellenv`. When neither system dir is writable the user's own bin dir
 * takes the link, so a non-admin install still gets a `genoffice` command.
 */
export function defaultCandidateDirs(
  platform: NodeJS.Platform,
  env: NodeJS.ProcessEnv = process.env,
): string[] {
  if (platform === 'darwin') return ['/usr/local/bin', '/opt/homebrew/bin', ...userBinDirs(env)]
  if (platform === 'linux') return ['/usr/local/bin', ...userBinDirs(env)]
  return []
}

/** `$XDG_BIN_HOME`, then `~/.local/bin`: created on demand because a fresh account rarely has them. */
export function userBinDirs(env: NodeJS.ProcessEnv): string[] {
  const dirs: string[] = []
  if (env.XDG_BIN_HOME && isAbsolute(env.XDG_BIN_HOME)) dirs.push(env.XDG_BIN_HOME)
  if (env.HOME) dirs.push(join(env.HOME, '.local', 'bin'))
  return [...new Set(dirs)]
}

/** Shell line that puts `dir` on PATH, or undefined when the current environment already searches it. */
export function pathHint(dir: string, env: NodeJS.ProcessEnv): string | undefined {
  const entries = (env.PATH ?? '').split(':').map((p) => p.replace(/\/+$/, ''))
  if (entries.includes(dir.replace(/\/+$/, ''))) return undefined
  const home = env.HOME?.replace(/\/+$/, '')
  const shown = home && dir.startsWith(home + '/') ? '$HOME' + dir.slice(home.length) : dir
  return `export PATH="${shown}:$PATH"`
}

export function installCliLink(opts: InstallOptions): InstallOutcome {
  const platform = opts.platform ?? process.platform
  if (platform === 'win32') return installWindowsPath(opts)
  if (platform !== 'darwin' && platform !== 'linux') return { status: 'unsupported' }
  const env = opts.env ?? process.env
  const dirs = opts.candidateDirs ?? defaultCandidateDirs(platform, env)
  const userDirs = new Set(userBinDirs(env))
  const manual = manualCommand(opts.launcher)
  let occupied: string | undefined
  for (const dir of dirs) {
    const link = join(dir, 'genoffice')
    const state = linkState(link, opts.launcher)
    if (state === 'ours' && readlinkSync(link) === opts.launcher) {
      return withPathHint({ status: 'present', location: link }, dir, userDirs, env)
    }
    if (state === 'file' || state === 'foreign') {
      // somebody else's genoffice (a file, or npm's symlink): never clobber it
      occupied = link
      continue
    }
    if (userDirs.has(dir) && !existsSync(dir)) {
      try {
        mkdirSync(dir, { recursive: true })
      } catch {
        continue
      }
    }
    if (!writable(dir)) continue
    try {
      // ours from another install dir, or a dead link nobody can run
      if (state !== 'missing') unlinkSync(link)
      symlinkSync(opts.launcher, link)
      return withPathHint({ status: 'linked', location: link }, dir, userDirs, env)
    } catch {
      continue
    }
  }
  if (occupied) return { status: 'occupied', location: occupied, manual }
  return { status: 'unwritable', location: join(dirs[0] ?? '/usr/local/bin', 'genoffice'), manual }
}

/** Read-only twin of `installCliLink`: what a fresh terminal would find, without changing anything. */
export function inspectCliLink(opts: InstallOptions): InstallOutcome {
  const platform = opts.platform ?? process.platform
  if (platform === 'win32') return inspectWindowsPath(opts)
  if (platform !== 'darwin' && platform !== 'linux') return { status: 'unsupported' }
  const env = opts.env ?? process.env
  const dirs = opts.candidateDirs ?? defaultCandidateDirs(platform, env)
  const userDirs = new Set(userBinDirs(env))
  const manual = manualCommand(opts.launcher)
  let occupied: string | undefined
  // same walk installCliLink does: an occupied name is skipped, the first free writable dir wins
  for (const dir of dirs) {
    const link = join(dir, 'genoffice')
    const state = linkState(link, opts.launcher)
    if (state === 'ours' && readlinkSync(link) === opts.launcher) {
      return withPathHint({ status: 'present', location: link }, dir, userDirs, env)
    }
    if (state === 'file' || state === 'foreign') {
      occupied ??= link
      continue
    }
    if (writable(dir) || (userDirs.has(dir) && creatable(dir))) {
      return withPathHint({ status: 'missing', location: link, manual }, dir, userDirs, env)
    }
  }
  if (occupied) return { status: 'occupied', location: occupied, manual }
  return { status: 'unwritable', location: join(dirs[0] ?? '/usr/local/bin', 'genoffice'), manual }
}

function manualCommand(launcher: string): string {
  return `sudo mkdir -p /usr/local/bin && sudo ln -sf "${launcher}" /usr/local/bin/genoffice`
}

function linkState(
  path: string,
  launcher: string,
): 'missing' | 'ours' | 'dangling' | 'file' | 'foreign' {
  try {
    const st = lstatSync(path)
    if (!st.isSymbolicLink()) return 'file'
    const target = readlinkSync(path)
    if (target === launcher) return 'ours'
    const resolved = isAbsolute(target) ? target : resolve(dirname(path), target)
    if (!existsSync(resolved)) return 'dangling'
    return isOurLauncher(resolved, launcher) ? 'ours' : 'foreign'
  } catch {
    return 'missing'
  }
}

/**
 * A link is ours when it resolves to the launcher we are installing, or to a
 * `genoffice` launcher shipped by another copy of the app (an earlier version,
 * a second install dir): one that has our CLI bundle or the Windows twin beside
 * it. A path that merely ends in `/cli/genoffice` belongs to whoever put it there.
 */
export function isOurLauncher(target: string, launcher: string): boolean {
  let real: string
  try {
    real = realpathSync(target)
  } catch {
    return false
  }
  try {
    if (real === realpathSync(launcher)) return true
  } catch {}
  if (basename(real) !== 'genoffice') return false
  const dir = dirname(real)
  return existsSync(join(dir, 'genoffice.cjs')) || existsSync(join(dir, 'genoffice.cmd'))
}

function writable(dir: string): boolean {
  try {
    accessSync(dir, constants.W_OK)
    return true
  } catch {
    return false
  }
}

/** A missing dir that `mkdir -p` could create: its nearest existing ancestor is writable. */
function creatable(dir: string): boolean {
  let probe = dir
  while (!existsSync(probe)) {
    const parent = dirname(probe)
    if (parent === probe) return false
    probe = parent
  }
  return writable(probe)
}

/** System dirs are on PATH by construction; only a per-user dir may need the shell told about it. */
function withPathHint(
  outcome: InstallOutcome,
  dir: string,
  userDirs: Set<string>,
  env: NodeJS.ProcessEnv,
): InstallOutcome {
  const hint = userDirs.has(dir) ? pathHint(dir, env) : undefined
  return hint ? { ...outcome, pathHint: hint } : outcome
}

/**
 * Windows keeps the launcher's directory on the user PATH. The registry value
 * is read and written unexpanded (REG_EXPAND_SZ) so entries like
 * `%USERPROFILE%\bin` survive, unlike `[Environment]::SetEnvironmentVariable`,
 * which writes back the expanded text. Running processes keep their copy of
 * the environment; the WM_SETTINGCHANGE broadcast makes Explorer reload so
 * terminals opened afterwards see the new PATH.
 */
const WIN_PATH_READ =
  "$key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Environment', $true); " +
  "$p = [string]$key.GetValue('Path', '', 'DoNotExpandEnvironmentNames'); "

// best effort: a locked-down PowerShell may refuse Add-Type, the PATH edit still stands
const WIN_BROADCAST =
  'try { ' +
  'Add-Type -Namespace GenOfficePath -Name Native -MemberDefinition \'[DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)] public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);\'; ' +
  '$r = [UIntPtr]::Zero; ' +
  "[void][GenOfficePath.Native]::SendMessageTimeout([IntPtr]0xffff, 0x1A, [UIntPtr]::Zero, 'Environment', 2, 5000, [ref]$r) " +
  '} catch {}; '

function psQuote(text: string): string {
  return `'${text.replace(/'/g, "''")}'`
}

function installWindowsPath(opts: InstallOptions): InstallOutcome {
  const dir = win32.dirname(opts.launcher)
  const manual = `[Environment]::SetEnvironmentVariable('Path', [Environment]::GetEnvironmentVariable('Path','User') + ';${dir}', 'User')`
  const script =
    `$dir = ${psQuote(dir)}; ` +
    WIN_PATH_READ +
    "if (($p -split ';') -contains $dir) { 'present'; exit 0 }; " +
    "$next = (@($p, $dir) | Where-Object { $_ }) -join ';'; " +
    "$key.SetValue('Path', $next, [Microsoft.Win32.RegistryValueKind]::ExpandString); " +
    WIN_BROADCAST +
    "'linked'"
  const run = opts.runPowerShell ?? runPowerShell
  const r = run(script)
  if (!r.ok) return { status: 'unwritable', location: dir, manual }
  return { status: r.stdout.trim() === 'present' ? 'present' : 'linked', location: dir }
}

function inspectWindowsPath(opts: InstallOptions): InstallOutcome {
  const dir = win32.dirname(opts.launcher)
  const manual = `[Environment]::SetEnvironmentVariable('Path', [Environment]::GetEnvironmentVariable('Path','User') + ';${dir}', 'User')`
  const script =
    `$dir = ${psQuote(dir)}; ` +
    "$key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Environment'); " +
    "$p = [string]$key.GetValue('Path', '', 'DoNotExpandEnvironmentNames'); " +
    "if (($p -split ';') -contains $dir) { 'present' } else { 'missing' }"
  const run = opts.runPowerShell ?? runPowerShell
  const r = run(script)
  if (!r.ok) return { status: 'unwritable', location: dir, manual }
  return { status: r.stdout.trim() === 'present' ? 'present' : 'missing', location: dir, manual }
}

function runPowerShell(script: string): { ok: boolean; stdout: string } {
  const r = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], {
    encoding: 'utf-8',
    windowsHide: true,
    timeout: 15_000,
  })
  return { ok: r.status === 0, stdout: r.stdout ?? '' }
}
