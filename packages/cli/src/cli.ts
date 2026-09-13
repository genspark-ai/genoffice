import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { flagBool, parseArgs } from './args'
import { appendAudit } from './audit'
import { capabilitiesCommand } from './commands/capabilities'
import { convertCommand } from './commands/convert'
import { docsCommand } from './commands/docs'
import { createCommand } from './commands/create'
import { guideCommand } from './commands/guide'
import { imageCommand } from './commands/image'
import { mediaCommand } from './commands/media'
import { searchCommand } from './commands/search'
import { infoCommand } from './commands/info'
import { installCommand } from './commands/install'
import { openCommand } from './commands/open'
import { renderCommand } from './commands/render'
import { sheetCommand } from './commands/sheet'
import { slidesCommand } from './commands/slides'
import { CommandRegistry, commandHelp, type CommandContext } from './registry'
import { CliError, EXIT, formatHuman, toJsonError, toJsonOk, type ExitCode } from './result'

declare const __GENOFFICE_VERSION__: string | undefined

/** The @genoffice/cli package version: inlined by build.mjs, read from disk when running from source. */
export const VERSION: string =
  typeof __GENOFFICE_VERSION__ === 'string' ? __GENOFFICE_VERSION__ : devVersion()

function devVersion(): string {
  try {
    const dir = typeof __dirname === 'string' ? __dirname : dirname(fileURLToPath(import.meta.url))
    return `${JSON.parse(readFileSync(join(dir, '..', 'package.json'), 'utf-8')).version}-dev`
  } catch {
    return '0.0.0-dev'
  }
}

export function defaultRegistry(): CommandRegistry {
  return new CommandRegistry()
    .register(infoCommand)
    .register(convertCommand)
    .register(createCommand)
    .register(slidesCommand)
    .register(sheetCommand)
    .register(docsCommand)
    .register(renderCommand)
    .register(guideCommand)
    .register(searchCommand)
    .register(imageCommand)
    .register(mediaCommand)
    .register(openCommand)
    .register(capabilitiesCommand)
    .register(installCommand)
}

export interface RunIo {
  stdout: (text: string) => void
  stderr: (text: string) => void
}

export interface RunOptions {
  cwd?: string
  env?: NodeJS.ProcessEnv
  registry?: CommandRegistry
  io?: RunIo
}

/** Runs one invocation; returns the exit code instead of exiting so tests and hosts can embed it. */
export async function runCli(argv: readonly string[], opts: RunOptions = {}): Promise<ExitCode> {
  const registry = opts.registry ?? defaultRegistry()
  const io = opts.io ?? {
    stdout: (t) => process.stdout.write(t + '\n'),
    stderr: (t) => process.stderr.write(t + '\n'),
  }
  const args = parseArgs(argv, booleanFlags(registry))
  const json = flagBool(args, 'json')
  const name = args.positionals.shift() ?? null

  const fail = (err: CliError): ExitCode => {
    if (json) io.stdout(JSON.stringify(toJsonError(name, err)))
    else
      io.stderr(
        `genoffice: ${err.message}` + (err.detail ? `\n  ${JSON.stringify(err.detail)}` : ''),
      )
    return err.code
  }

  if (flagBool(args, 'version')) {
    io.stdout(json ? JSON.stringify({ status: 'ok', version: VERSION }) : `genoffice ${VERSION}`)
    return EXIT.ok
  }
  if (name === null || (name === 'help' && args.positionals.length === 0)) {
    io.stdout(globalHelp(registry))
    return name === null && !flagBool(args, 'help') ? EXIT.usage : EXIT.ok
  }
  const def = registry.get(name === 'help' ? args.positionals[0] : name)
  if (!def)
    return fail(
      new CliError(EXIT.usage, `unknown command: ${name}`, {
        commands: registry.list().map((d) => d.name),
      }),
    )
  if (name === 'help' || flagBool(args, 'help')) {
    io.stdout(commandHelp(def))
    return EXIT.ok
  }
  const known = new Set(['json', 'help', 'version', ...(def.options ?? []).map((o) => o.name)])
  const unknown = Object.keys(args.flags).filter((f) => !known.has(f))
  if (unknown.length) {
    return fail(
      new CliError(
        EXIT.usage,
        `unknown option${unknown.length > 1 ? 's' : ''} for ${def.name}: ${unknown.map((f) => `--${f}`).join(', ')}`,
        {
          options: [...known].map((f) => `--${f}`),
        },
      ),
    )
  }

  const ctx: CommandContext = {
    cwd: opts.cwd ?? process.cwd(),
    env: opts.env ?? process.env,
    log: (m) => io.stderr(m),
  }
  const started = Date.now()
  const audit = (status: 'ok' | 'error', code: ExitCode, outputPath?: string) =>
    appendAudit(ctx.env, {
      command: def.name,
      argv: redactArgv(argv),
      status,
      code,
      ...(outputPath ? { output_path: outputPath } : {}),
      ms: Date.now() - started,
      cwd: ctx.cwd,
    })
  try {
    const result = await def.run(args, ctx)
    audit('ok', EXIT.ok, result.outputPath)
    io.stdout(json ? JSON.stringify(toJsonOk(def.name, result)) : formatHuman(result))
    return EXIT.ok
  } catch (err) {
    const error =
      err instanceof CliError
        ? err
        : new CliError(EXIT.conversion, err instanceof Error ? err.message : String(err))
    audit('error', error.code)
    return fail(error)
  }
}

const SECRET_FLAG = /^--(password|api[-_]?key|token|secret)$/i

/** Secrets never reach the audit log; the flag stays so the call is still recognizable. */
export function redactArgv(argv: readonly string[]): string[] {
  const out: string[] = []
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]!
    const eq = arg.indexOf('=')
    if (eq !== -1 && SECRET_FLAG.test(arg.slice(0, eq))) {
      out.push(`${arg.slice(0, eq)}=***`)
      continue
    }
    out.push(arg)
    if (SECRET_FLAG.test(arg) && i + 1 < argv.length) {
      out.push('***')
      i++
    }
  }
  return out
}

function booleanFlags(registry: CommandRegistry): Set<string> {
  const names = new Set(['json', 'help', 'version'])
  for (const def of registry.list()) {
    for (const o of def.options ?? []) if (!o.value) names.add(o.name)
  }
  return names
}

function globalHelp(registry: CommandRegistry): string {
  const defs = registry.list()
  const width = Math.max(...defs.map((d) => d.name.length))
  return [
    `genoffice ${VERSION} — GenOffice command line`,
    '',
    'Usage: genoffice <command> [options]',
    '',
    'Commands:',
    ...defs.map((d) => `  ${d.name.padEnd(width)}  ${d.summary}`),
    '',
    'Global options:',
    '  --json     machine-readable output (one JSON object on stdout)',
    '  --help     show help for a command',
    '  --version  print the version',
    '',
    'Exit codes: 0 ok, 1 usage, 2 file, 3 conversion failed, 4 app not available',
  ].join('\n')
}

const isMain =
  typeof require !== 'undefined' && typeof module !== 'undefined' && require.main === module
if (isMain) {
  // stdout is async on a pipe: exiting right after write() drops everything past the
  // 64 KiB pipe buffer, so a `--json` consumer would see a cut JSON object
  const exit = (code: number) => process.stdout.write('', () => process.exit(code))
  runCli(process.argv.slice(2)).then(exit, (err) => {
    process.stderr.write(
      `genoffice: ${err instanceof Error ? (err.stack ?? err.message) : String(err)}\n`,
    )
    exit(EXIT.conversion)
  })
}
