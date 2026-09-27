/**
 * How an MCP client starts `genoffice mcp`. Shared by the app's Settings snippet
 * and `genoffice mcp install` so the two cannot drift. Clients spawn without a
 * shell, so on Windows neither genoffice.cmd nor `cmd /c` is safe (a path with
 * a space splits); the entry does what genoffice.cmd does instead: the app
 * binary as Node on the bundled CLI. Pure string code: the renderer imports it.
 */
export interface McpLaunch {
  command: string
  args: string[]
  env?: Record<string, string>
}

const isWindowsPath = (p: string) => p.includes('\\')

/** The app's snippet: the bare name once genoffice is on the PATH, else the launcher itself. */
export function mcpLaunch(cli: { status: string; launcherDir: string }): McpLaunch {
  const dir = cli.launcherDir
  if (isWindowsPath(dir)) return windowsAppLaunch(dir)
  return { command: cli.status === 'present' ? 'genoffice' : `${dir}/genoffice`, args: ['mcp'] }
}

function windowsAppLaunch(dir: string): McpLaunch {
  return {
    command: `${dir}\\..\\..\\GenOffice.exe`,
    args: [`${dir}\\genoffice.cjs`, 'mcp'],
    env: { ELECTRON_RUN_AS_NODE: '1' },
  }
}

export interface LauncherLaunchOptions {
  platform?: string
  /** test seam for the packaged-app probe */
  exists?: (path: string) => boolean
}

/**
 * `mcp install`: always the absolute launcher, never the bare name. On Windows
 * the packaged app is run as Node (same entry as the app's snippet); a checkout
 * has no GenOffice.exe beside it and runs the bundle on the system node, as the
 * bin/genoffice script does.
 */
export function mcpLaunchFromLauncher(
  launcher: string,
  opts: LauncherLaunchOptions = {},
): McpLaunch {
  const win = isWindowsPath(launcher) || opts.platform === 'win32'
  if (!win) return { command: launcher, args: ['mcp'] }
  const sep = isWindowsPath(launcher) ? '\\' : '/'
  const dir = launcher.slice(0, Math.max(launcher.lastIndexOf('\\'), launcher.lastIndexOf('/')))
  const exists = opts.exists ?? (() => false)
  if (exists(`${dir}${sep}..${sep}..${sep}GenOffice.exe`)) return windowsAppLaunch(dir)
  const bundle = exists(`${dir}${sep}genoffice.cjs`)
    ? `${dir}${sep}genoffice.cjs`
    : `${dir}${sep}..${sep}dist${sep}genoffice.cjs`
  return { command: 'node', args: [bundle, 'mcp'] }
}
