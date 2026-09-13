import { join } from 'node:path'
import { installCliLink } from '../install'
import type { CommandDef } from '../registry'
import { packagedResourcesDir, repoRoot } from '../resources'
import { CliError, EXIT } from '../result'

/** The launcher the terminal should run: the one shipped beside this bundle, or the checkout's. */
export function launcherPath(): string | null {
  const packaged = packagedResourcesDir()
  if (packaged)
    return join(packaged, 'cli', process.platform === 'win32' ? 'genoffice.cmd' : 'genoffice')
  const root = repoRoot()
  return root ? join(root, 'packages', 'cli', 'bin', 'genoffice') : null
}

export const installCommand: CommandDef = {
  name: 'install-cli',
  summary: 'Put genoffice on the PATH (symlink into /usr/local/bin, or the user PATH on Windows).',
  usage: 'install-cli',
  async run() {
    const launcher = launcherPath()
    if (!launcher) throw new CliError(EXIT.app, 'cannot locate the genoffice launcher')
    const r = installCliLink({ launcher })
    const detail = { status: r.status, launcher, ...(r.location ? { location: r.location } : {}) }
    switch (r.status) {
      case 'linked':
        return { summary: `genoffice is now on the PATH (${r.location})`, detail }
      case 'present':
        return { summary: `genoffice is already on the PATH (${r.location})`, detail }
      case 'occupied':
        throw new CliError(EXIT.app, `${r.location} is another program; not replacing it`, {
          ...detail,
          manual: r.manual,
        })
      case 'unwritable':
        throw new CliError(EXIT.app, `${r.location} is not writable; run the manual command`, {
          ...detail,
          manual: r.manual,
        })
      default:
        throw new CliError(EXIT.app, `PATH setup is not supported on ${process.platform}`, detail)
    }
  },
}
