import { readFileSync, readSync, statSync } from 'node:fs'
import { flagString, type ParsedArgs } from './args'
import { resolveInput, type PathContext } from './fs'
import { MAX_TRANSFER_BYTES } from './mcp/files'
import { CliError, EXIT } from './result'

/** Same ceiling as MCP file transfers: decks inline pictures as base64 in addPicture.bytes. */
export const MAX_OPS_BYTES = MAX_TRANSFER_BYTES
const STDIN_CHUNK_BYTES = 1024 * 1024

function throwIfOpsOverBudget(bytes: number, source: string): void {
  if (bytes > MAX_OPS_BYTES) {
    throw new CliError(
      EXIT.file,
      `ops input too large: ${source} is ${bytes} bytes (cap ${MAX_OPS_BYTES}); split into smaller batches`,
      { source, bytes, cap: MAX_OPS_BYTES },
      {
        reason: 'resource_limit',
        suggestion: 'split into smaller --ops batches',
      },
    )
  }
}

/** Reads a stream fd to its end, throwing as soon as the running total passes the cap. */
export function readOpsStream(fd: number, source: string): string {
  const chunks: Buffer[] = []
  const buf = Buffer.allocUnsafe(STDIN_CHUNK_BYTES)
  let total = 0
  for (;;) {
    const n = readSync(fd, buf, 0, STDIN_CHUNK_BYTES, null)
    if (n === 0) break
    total += n
    throwIfOpsOverBudget(total, source)
    chunks.push(Buffer.from(buf.subarray(0, n)))
  }
  return Buffer.concat(chunks, total).toString('utf-8')
}

/** `--ops <file>` or `--ops -` (stdin); returns the raw text and a label for error messages. */
export function readOpsInput(args: ParsedArgs, ctx: PathContext): { text: string; source: string } {
  const spec = flagString(args, 'ops')
  if (!spec)
    throw new CliError(EXIT.usage, 'missing --ops <file|->', undefined, {
      reason: 'missing_argument',
    })
  if (spec === '-') return { text: readOpsStream(0, 'stdin'), source: 'stdin' }
  const path = resolveInput(spec, ctx)
  const size = statSync(path).size
  throwIfOpsOverBudget(size, path)
  return { text: readFileSync(path, 'utf-8'), source: path }
}
