import { readFileSync } from 'node:fs'
import { flagString, type ParsedArgs } from './args'
import { resolveInput, type PathContext } from './fs'
import { CliError, EXIT } from './result'

/** `--ops <file>` or `--ops -` (stdin); returns the raw text and a label for error messages. */
export function readOpsInput(args: ParsedArgs, ctx: PathContext): { text: string; source: string } {
  const spec = flagString(args, 'ops')
  if (!spec)
    throw new CliError(EXIT.usage, 'missing --ops <file|->', undefined, {
      reason: 'missing_argument',
    })
  if (spec === '-') return { text: readFileSync(0, 'utf-8'), source: 'stdin' }
  const path = resolveInput(spec, ctx)
  return { text: readFileSync(path, 'utf-8'), source: path }
}
