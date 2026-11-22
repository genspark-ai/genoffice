import { existsSync } from 'node:fs'
import { createConnection } from 'node:net'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { ZOTERO_INTEGRATION_HOST, ZOTERO_INTEGRATION_PORT } from './zotero-wire'

/** where a Zotero install or its profile lives; a profile outlasts a portable or moved app */
export function zoteroInstallPaths(
  platform: NodeJS.Platform,
  env: NodeJS.ProcessEnv,
  home: string,
): string[] {
  if (platform === 'darwin') {
    return [
      '/Applications/Zotero.app',
      join(home, 'Applications', 'Zotero.app'),
      join(home, 'Library', 'Application Support', 'Zotero'),
    ]
  }
  if (platform === 'win32') {
    return [
      [env.ProgramFiles, 'Zotero', 'zotero.exe'],
      [env['ProgramFiles(x86)'], 'Zotero', 'zotero.exe'],
      [env.LOCALAPPDATA, 'Zotero', 'zotero.exe'],
      [env.APPDATA, 'Zotero', 'Zotero'],
    ].flatMap(([root, ...rest]) => (root ? [join(root, ...(rest as string[]))] : []))
  }
  return [join(home, '.zotero', 'zotero'), '/usr/lib/zotero', '/opt/zotero']
}

function zoteroListening(host: string, port: number): Promise<boolean> {
  return new Promise((resolve) => {
    const socket = createConnection({ host, port })
    const done = (up: boolean) => {
      socket.destroy()
      resolve(up)
    }
    socket.setTimeout(500, () => done(false))
    socket.once('connect', () => done(true))
    socket.once('error', () => done(false))
  })
}

export async function zoteroAvailable(
  options: {
    env?: NodeJS.ProcessEnv
    exists?: (path: string) => boolean
    host?: string
    port?: number
  } = {},
): Promise<boolean> {
  const env = options.env ?? process.env
  // tests pin the answer: the ribbon must not depend on what the machine has installed
  if (env.GENOFFICE_ZOTERO === '1') return true
  if (env.GENOFFICE_ZOTERO === '0') return false
  const exists = options.exists ?? existsSync
  if (zoteroInstallPaths(process.platform, env, homedir()).some(exists)) return true
  return zoteroListening(
    options.host ?? ZOTERO_INTEGRATION_HOST,
    options.port ?? ZOTERO_INTEGRATION_PORT,
  )
}
