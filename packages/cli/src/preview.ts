import { flagBool, flagString, type ParsedArgs } from './args'
import { CliError, EXIT } from './result'

/** Clip to `max` characters; the marker names how much is missing so a reader can decide whether to fetch it. */
export function clipText(text: string, max: number): { text: string; truncated: boolean } {
  if (text.length <= max) return { text, truncated: false }
  return { text: `${text.slice(0, max)}…(+${text.length - max} chars)`, truncated: true }
}

/** `--full` lifts the cap, `--max-chars n` moves it. */
export function previewChars(args: ParsedArgs, fallback: number): number {
  if (flagBool(args, 'full')) return Infinity
  const raw = flagString(args, 'max-chars')
  if (raw === undefined) return fallback
  const n = Number(raw)
  if (!Number.isInteger(n) || n < 1)
    throw new CliError(
      EXIT.usage,
      `--max-chars must be a positive integer (got ${raw})`,
      undefined,
      {
        reason: 'invalid_argument',
      },
    )
  return n
}
