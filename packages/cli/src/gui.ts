import { existsSync, readFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { realizedPath } from './fs'
import { CliError, EXIT } from './result'

/** The shell's Electron userData directory, located without Electron (GENOFFICE_USER_DATA overrides). */
export function genofficeUserDataDir(env: NodeJS.ProcessEnv): string {
  if (env.GENOFFICE_USER_DATA) return env.GENOFFICE_USER_DATA
  const base =
    process.platform === 'darwin'
      ? join(homedir(), 'Library', 'Application Support')
      : process.platform === 'win32'
        ? env.APPDATA || join(homedir(), 'AppData', 'Roaming')
        : env.XDG_CONFIG_HOME || join(homedir(), '.config')
  return join(base, 'GenOffice')
}

export interface GuiOpenDocuments {
  pid: number
  paths: string[]
}

/**
 * Files the running GenOffice shell has open, from the registry it publishes
 * on every tab change (apps/shell/src/main/open-documents.ts). Null when no
 * shell is running: a registry whose pid is gone is a crash leftover.
 */
export function guiOpenDocuments(env: NodeJS.ProcessEnv): GuiOpenDocuments | null {
  // a checkout's shell keeps its userData in "<name> Dev" beside the packaged one
  const dirs = env.GENOFFICE_USER_DATA
    ? [env.GENOFFICE_USER_DATA]
    : [genofficeUserDataDir(env), `${genofficeUserDataDir(env)} Dev`]
  for (const dir of dirs) {
    const path = join(dir, 'open-documents.json')
    if (!existsSync(path)) continue
    try {
      const raw = JSON.parse(readFileSync(path, 'utf8')) as Partial<GuiOpenDocuments>
      if (typeof raw.pid !== 'number' || !Array.isArray(raw.paths)) continue
      if (!processAlive(raw.pid)) continue
      return { pid: raw.pid, paths: raw.paths.filter((p): p is string => typeof p === 'string') }
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

/** Refuses to write a document the editor is showing; `--force` skips this. */
export function assertNotOpenInGui(abs: string, env: NodeJS.ProcessEnv): void {
  const open = guiOpenDocuments(env)
  if (!open || open.paths.length === 0) return
  const target = realizedPath(abs)
  if (!open.paths.some((p) => realizedPath(p) === target)) return
  throw new CliError(
    EXIT.file,
    `GenOffice has this file open: ${abs}`,
    { gui_pid: open.pid },
    {
      reason: 'file_open_in_gui',
      suggestion:
        'close the tab in GenOffice first, or pass --force to write anyway (the editor may overwrite your change on its next save)',
    },
  )
}
