import { spawn } from 'node:child_process'
import { resolveInput } from '../fs'
import type { CommandDef } from '../registry'
import { appLaunch } from '../resources'
import { CliError, EXIT } from '../result'

/** The shell's second-instance handler forwards the path when the app is already running. */
export const openCommand: CommandDef = {
  name: 'open',
  summary: 'Open a document in the GenOffice app (starts the app if needed).',
  usage: 'open <file>',
  async run(args, ctx) {
    const path = resolveInput(args.positionals[0], ctx)
    const launch = appLaunch(ctx.env)
    if (!launch) {
      throw new CliError(EXIT.app, 'GenOffice app not found', { hint: 'set GENOFFICE_APP_BIN' })
    }
    const env = { ...ctx.env }
    delete env.ELECTRON_RUN_AS_NODE
    await new Promise<void>((resolve, reject) => {
      const child = spawn(launch.command, [...launch.args, path], {
        detached: true,
        stdio: 'ignore',
        env,
      })
      child.once('error', (err) =>
        reject(new CliError(EXIT.app, `failed to start GenOffice: ${err.message}`)),
      )
      child.once('spawn', () => {
        child.unref()
        resolve()
      })
    })
    return { summary: `opening ${path} in GenOffice`, detail: { app: launch.command } }
  },
}
