import { CliError, EXIT } from './result'

export interface ParsedArgs {
  positionals: string[]
  flags: Record<string, string | true>
}

/**
 * `--key value`, `--key=value`, `--flag`, `-h`; a lone `--` ends flag parsing.
 * Names in `booleans` never take a value, so `genoffice --json info a.docx` keeps
 * `info` as the command.
 */
export function parseArgs(
  argv: readonly string[],
  booleans: ReadonlySet<string> = new Set(),
): ParsedArgs {
  const positionals: string[] = []
  const flags: Record<string, string | true> = {}
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]
    if (arg === '--') {
      positionals.push(...argv.slice(i + 1))
      break
    }
    if (arg === '-h') {
      flags.help = true
      continue
    }
    if (!arg.startsWith('--') || arg.length === 2) {
      positionals.push(arg)
      continue
    }
    const eq = arg.indexOf('=')
    if (eq !== -1) {
      const key = arg.slice(2, eq)
      const value = arg.slice(eq + 1)
      // --force=false used to store 'false', which flagBool read as present
      if (booleans.has(key)) {
        throw new CliError(
          EXIT.usage,
          `--${key} is a boolean flag; use --${key}, not --${key}=<value>`,
          undefined,
          { reason: 'invalid_argument' },
        )
      }
      flags[key] = value
      continue
    }
    const key = arg.slice(2)
    const next = argv[i + 1]
    if (!booleans.has(key) && next !== undefined && !next.startsWith('--')) {
      flags[key] = next
      i++
    } else {
      flags[key] = true
    }
  }
  return { positionals, flags }
}

export function flagString(args: ParsedArgs, name: string): string | undefined {
  const v = args.flags[name]
  return typeof v === 'string' ? v : undefined
}

export function flagBool(args: ParsedArgs, name: string): boolean {
  return args.flags[name] !== undefined
}
