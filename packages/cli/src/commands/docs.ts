import { writeFileSync } from 'node:fs'
import { basename, dirname } from 'node:path'
import { flagBool, flagString } from '../args'
import {
  applyDocOps,
  blockRangeHtml,
  closeDocument,
  describeDocument,
  headerFooterState,
  listComments,
  listRevisions,
  openDocument,
  saveDocument,
} from '../formats/docx'
import { readInput, resolveInput, resolveOutput } from '../fs'
import { readOpsInput } from '../ops-input'
import type { CommandContext, CommandDef } from '../registry'
import { CliError, EXIT, type CommandResult } from '../result'

export const docsCommand: CommandDef = {
  name: 'docs',
  summary: 'Read or edit a .docx with the same ops the in-app AI uses.',
  usage: 'docs <read|apply> <file.docx> [options]',
  options: [
    { name: 'range', value: 'a-b', description: 'read: block index range (default: all)' },
    { name: 'html', description: 'read: also return the range as restricted HTML' },
    { name: 'full', description: 'read: whole block text instead of a 200-character preview' },
    {
      name: 'comments',
      description:
        'read: list every comment thread (ids for reply_comment / resolve_comment, anchored block, resolved state)',
    },
    { name: 'revisions', description: 'read: list pending tracked changes with their block index' },
    {
      name: 'header-footer',
      description: 'read: the header and footer text per variant ({PAGE} / {NUMPAGES} tokens)',
    },
    {
      name: 'ops',
      value: 'file',
      description:
        'apply: JSON array of ops (apply_ops entries, plus insert_content / replace_blocks with an html field, insert_image, insert_chart / edit_chart, set_header_footer, reply_comment / resolve_comment); "-" reads stdin. See `genoffice guide docs`.',
    },
    {
      name: 'dry-run',
      description: 'apply: run the batch in memory and report each step without writing',
    },
    { name: 'out', value: 'path', description: 'apply: write here instead of in place' },
    {
      name: 'force',
      description:
        'apply: overwrite an existing --out file, or write while GenOffice has the file open',
    },
  ],
  async run(args, ctx) {
    const [verb, file] = args.positionals
    switch (verb) {
      case 'read':
        return read(file, args, ctx)
      case 'apply':
        return apply(file, args, ctx)
      default:
        throw new CliError(EXIT.usage, 'expected "docs read <file>" or "docs apply <file>"')
    }
  },
}

type Args = Parameters<CommandDef['run']>[0]

function parseRange(spec: string | undefined, count: number): [number, number] | undefined {
  if (spec === undefined) return undefined
  const m = /^(\d+)(?:-(\d+))?$/.exec(spec)
  if (!m) throw new CliError(EXIT.usage, '--range must look like 3 or 3-10 (block indexes)')
  const start = Number(m[1])
  const end = m[2] === undefined ? start : Number(m[2])
  if (start > end || end >= count) {
    throw new CliError(EXIT.usage, `--range out of bounds (0-${count - 1})`)
  }
  return [start, end]
}

async function read(
  file: string | undefined,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const doc = await openDocument(readInput(path))
  try {
    const count = doc.editor.state.doc.childCount
    const range = parseRange(flagString(args, 'range'), count)
    const blocks = describeDocument(doc, range, flagBool(args, 'full'))
    const detail: Record<string, unknown> = {
      blocks: count,
      range: range ? `${range[0]}-${range[1]}` : `0-${count - 1}`,
      items: blocks,
      units: 'block indexes are 0-based positions in the body; ops target them',
    }
    if (flagBool(args, 'html')) {
      const [s, e] = range ?? [0, count - 1]
      detail.html = blockRangeHtml(doc, s, e)
    }
    if (flagBool(args, 'comments')) detail.comments = listComments(doc)
    if (flagBool(args, 'revisions')) detail.revisions = listRevisions(doc)
    if (flagBool(args, 'header-footer')) detail.headerFooter = headerFooterState(doc)
    return { summary: `${basename(path)}: ${count} blocks`, detail }
  } finally {
    closeDocument(doc)
  }
}

async function apply(
  file: string | undefined,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const { text, source } = readOpsInput(args, ctx)
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch (err) {
    throw new CliError(EXIT.usage, `${source}: not valid JSON (${(err as Error).message})`)
  }
  const ops = Array.isArray(parsed)
    ? parsed
    : parsed && typeof parsed === 'object' && Array.isArray((parsed as { ops?: unknown }).ops)
      ? (parsed as { ops: unknown[] }).ops
      : null
  if (!ops?.length) throw new CliError(EXIT.usage, `${source}: expected a non-empty array of ops`)
  const dryRun = flagBool(args, 'dry-run')
  const doc = await openDocument(readInput(path))
  try {
    const results = await applyDocOps(doc, ops as Record<string, unknown>[], {
      ctx,
      baseDir: source === 'stdin' ? undefined : dirname(source),
    })
    if (dryRun) {
      return {
        summary: `dry run: ${ops.length} ops validated, nothing written`,
        detail: { plan: results.map((r, i) => `${i}: ${r.op}: ${r.output.split('\n')[0]}`) },
      }
    }
    const output = resolveOutput(flagString(args, 'out'), ctx, {
      fallback: path,
      force: flagBool(args, 'force'),
      // in-place edits overwrite by design; --out onto another existing file needs --force
      fresh: flagString(args, 'out') !== undefined,
    })
    writeFileSync(output, await saveDocument(doc))
    return {
      summary: `applied ${ops.length} ops to ${basename(output)}`,
      outputPath: output,
      detail: { results, blocks: doc.editor.state.doc.childCount },
    }
  } finally {
    closeDocument(doc)
  }
}
