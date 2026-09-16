import type { ParsedArgs } from './args'
import type { CommandResult, Warning } from './result'

export interface CommandContext {
  cwd: string
  env: NodeJS.ProcessEnv
  /** progress and diagnostics; never part of the machine-readable result */
  log: (message: string) => void
  /** advisories that belong in the result: merged into `warnings` when the command returns */
  warn: (warning: Warning) => void
}

export interface OptionDef {
  name: string
  description: string
  /** placeholder for the option's value; absent for boolean switches */
  value?: string
}

export interface CommandDef {
  name: string
  summary: string
  usage: string
  options?: OptionDef[]
  /** the command owns stdout for its whole run (a server on stdio); runCli prints nothing after it */
  quiet?: boolean
  run(args: ParsedArgs, ctx: CommandContext): Promise<CommandResult>
}

export class CommandRegistry {
  private readonly defs = new Map<string, CommandDef>()

  register(def: CommandDef): this {
    this.defs.set(def.name, def)
    return this
  }

  get(name: string): CommandDef | undefined {
    return this.defs.get(name)
  }

  list(): CommandDef[] {
    return [...this.defs.values()]
  }
}

export function commandHelp(def: CommandDef): string {
  const lines = [`Usage: genoffice ${def.usage}`, '', def.summary]
  if (def.options?.length) {
    lines.push('', 'Options:')
    const width = Math.max(...def.options.map((o) => optionLabel(o).length))
    for (const o of def.options) lines.push(`  ${optionLabel(o).padEnd(width)}  ${o.description}`)
  }
  return lines.join('\n')
}

function optionLabel(o: OptionDef): string {
  return o.value ? `--${o.name} <${o.value}>` : `--${o.name}`
}
