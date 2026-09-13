import { spawnSync } from 'node:child_process'
import { accessSync, constants, lstatSync, readlinkSync, symlinkSync, unlinkSync } from 'node:fs'
import { join, win32 } from 'node:path'

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
}

/**
 * `/usr/local/bin` is on every login shell's PATH and on the PATH GUI apps
 * inherit, so it is the target even when it does not exist yet: a missing
 * directory is reported as unwritable (creating it needs root), not skipped.
 * `/opt/homebrew/bin` is only a fallback because shells see it solely through
 * `brew shellenv`.
 */
export function defaultCandidateDirs(platform: NodeJS.Platform): string[] {
  if (platform === 'darwin') return ['/usr/local/bin', '/opt/homebrew/bin']
  if (platform === 'linux') return ['/usr/local/bin']
  return []
}

export function installCliLink(opts: InstallOptions): InstallOutcome {
  const platform = opts.platform ?? process.platform
  if (platform === 'win32') return installWindowsPath(opts)
  if (platform !== 'darwin' && platform !== 'linux') return { status: 'unsupported' }
  const dirs = opts.candidateDirs ?? defaultCandidateDirs(platform)
  const manual = manualCommand(opts.launcher)
  let occupied: string | undefined
  for (const dir of dirs) {
    const link = join(dir, 'genoffice')
    const state = linkState(link, opts.launcher)
    if (state === 'ours' && readlinkSync(link) === opts.launcher) {
      return { status: 'present', location: link }
    }
    if (state === 'file' || state === 'foreign') {
      // somebody else's genoffice (a file, or npm's symlink): never clobber it
      occupied = link
      continue
    }
    if (!writable(dir)) continue
    try {
      if (state !== 'missing') unlinkSync(link)
      symlinkSync(opts.launcher, link)
      return { status: 'linked', location: link }
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
  const dirs = opts.candidateDirs ?? defaultCandidateDirs(platform)
  const manual = manualCommand(opts.launcher)
  let occupied: string | undefined
  // same walk installCliLink does: an occupied name is skipped, the first free writable dir wins
  for (const dir of dirs) {
    const link = join(dir, 'genoffice')
    const state = linkState(link, opts.launcher)
    if (state === 'ours' && readlinkSync(link) === opts.launcher) {
      return { status: 'present', location: link }
    }
    if (state === 'file' || state === 'foreign') {
      occupied ??= link
      continue
    }
    if (writable(dir)) return { status: 'missing', location: link, manual }
  }
  if (occupied) return { status: 'occupied', location: occupied, manual }
  return { status: 'unwritable', location: join(dirs[0] ?? '/usr/local/bin', 'genoffice'), manual }
}

function manualCommand(launcher: string): string {
  return `sudo mkdir -p /usr/local/bin && sudo ln -sf "${launcher}" /usr/local/bin/genoffice`
}

function linkState(path: string, launcher: string): 'missing' | 'ours' | 'file' | 'foreign' {
  try {
    const st = lstatSync(path)
    if (!st.isSymbolicLink()) return 'file'
    const target = readlinkSync(path)
    return target === launcher || isOurLauncher(target) ? 'ours' : 'foreign'
  } catch {
    return 'missing'
  }
}

/** Only launchers we shipped (<app resources>/cli/genoffice, any version or install dir) may be replaced. */
function isOurLauncher(target: string): boolean {
  return /[\\/]cli[\\/]genoffice$/.test(target)
}

function writable(dir: string): boolean {
  try {
    accessSync(dir, constants.W_OK)
    return true
  } catch {
    return false
  }
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
