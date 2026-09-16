import { basename, dirname } from 'node:path'
import { flagBool, flagString } from '../args'
import {
  PREVIEW_CHARS,
  applyDocOps,
  type DocOpsOutcome,
  blockRangeHtml,
  closeDocument,
  describeDocument,
  headerFooterState,
  listComments,
  listFields,
  listNotes,
  listTableStyles,
  listRevisions,
  listStyles,
  openDocument,
  saveDocument,
} from '../formats/docx'
import { previewChars } from '../preview'
import { checkDocument, DOCS_CHECKS } from '../formats/docx-check'
import { listSections } from '../formats/docx-sections'
import { readInput, resolveInput, resolveOutput, writeOutput } from '../fs'
import { readOpsInput } from '../ops-input'
import type { CommandContext, CommandDef } from '../registry'
import { checkResult } from '../check'
import { CliError, EXIT, type CommandResult } from '../result'
import { BATCH_OPTIONS, batchCounts, batchMode, batchResult, failedBatch } from '../batch'
import { opSuggestion } from '../op-errors'

export const docsCommand: CommandDef = {
  name: 'docs',
  summary: 'Read or edit a .docx with the same ops the in-app AI uses.',
  usage: 'docs <read|apply|check> <file.docx> [options]',
  options: [
    { name: 'range', value: 'a-b', description: 'read: block index range (default: all)' },
    { name: 'html', description: 'read: also return the range as restricted HTML' },
    { name: 'full', description: 'read: whole block text instead of a 200-character preview' },
    {
      name: 'max-chars',
      value: 'n',
      description: 'read: preview length per block (default 200); clipped text ends in …(+n chars)',
    },
    {
      name: 'comments',
      description:
        'read: list every comment (id, author, date, parent id, anchored block and text, resolved state) for reply_comment / resolve_comment / delete_comment',
    },
    {
      name: 'revisions',
      description:
        'read: list pending tracked changes (ids for accept_changes / reject_changes, type, author, block index, text)',
    },
    {
      name: 'styles',
      description:
        'read: list the styles defined in the document (ids for applyStyle / define_style, type, parent, usage count)',
    },
    {
      name: 'header-footer',
      description: 'read: the header and footer text per variant ({PAGE} / {NUMPAGES} tokens)',
    },
    {
      name: 'sections',
      description:
        'read: page setup per section (paper, orientation, margins, columns, block range) for set_page_setup',
    },
    {
      name: 'fields',
      description:
        'read: list every field (SEQ, REF, DATE, PAGE, TOC …) with its block index, instruction, cached result and whether Word recomputes it on open',
    },
    {
      name: 'notes',
      description: 'read: list footnotes and endnotes (ids for delete_note, anchored block index)',
    },
    {
      name: 'ops',
      value: 'file',
      description:
        'apply: JSON array of ops (apply_ops entries, plus insert_content / replace_blocks with an html field, insert_image / insert_picture / insert_text_box, insert_chart / edit_chart, set_header_footer, set_page_setup / insert_section_break, set_watermark, define_style, add_comment / reply_comment / resolve_comment / delete_comment, accept_changes / reject_changes, insert_footnote / insert_endnote / delete_note); "-" reads stdin. See `genoffice guide docs`.',
    },
    {
      name: 'track',
      description:
        'apply: record content edits as tracked changes (author: --author, default "AI Assistant")',
    },
    { name: 'author', value: 'name', description: 'apply --track: author of the tracked changes' },
    {
      name: 'dry-run',
      description: 'apply: run the batch in memory and report each step without writing',
    },
    ...BATCH_OPTIONS,
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
      case 'check':
        return check(file, ctx)
      default:
        throw new CliError(
          EXIT.usage,
          'expected "docs read <file>", "docs apply <file>" or "docs check <file>"',
          undefined,
          {
            reason: verb === undefined ? 'missing_argument' : 'invalid_argument',
            suggestion: 'run `genoffice help docs`',
          },
        )
    }
  },
}

type Args = Parameters<CommandDef['run']>[0]

function parseRange(spec: string | undefined, count: number): [number, number] | undefined {
  if (spec === undefined) return undefined
  const m = /^(\d+)(?:-(\d+))?$/.exec(spec)
  if (!m) {
    throw new CliError(EXIT.usage, '--range must look like 3 or 3-10 (block indexes)', undefined, {
      reason: 'invalid_argument',
    })
  }
  const start = Number(m[1])
  const end = m[2] === undefined ? start : Number(m[2])
  if (start > end || end >= count) {
    throw new CliError(
      EXIT.usage,
      `--range out of bounds (0-${count - 1})`,
      { valid_range: [0, count - 1] },
      { reason: 'out_of_range', suggestion: `use block indexes between 0 and ${count - 1}` },
    )
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
    const blocks = describeDocument(doc, range, previewChars(args, PREVIEW_CHARS))
    const detail: Record<string, unknown> = {
      blocks: count,
      range: range ? `${range[0]}-${range[1]}` : `0-${count - 1}`,
      items: blocks,
      units: 'block indexes are 0-based positions in the body; ops target them',
    }
    if (blocks.some((b) => b.table)) detail.tableStyles = listTableStyles(doc)
    if (flagBool(args, 'html')) {
      const [s, e] = range ?? [0, count - 1]
      detail.html = blockRangeHtml(doc, s, e)
    }
    if (flagBool(args, 'comments')) detail.comments = listComments(doc)
    if (flagBool(args, 'revisions')) detail.revisions = listRevisions(doc)
    if (flagBool(args, 'styles')) detail.styles = listStyles(doc)
    if (flagBool(args, 'header-footer')) detail.headerFooter = headerFooterState(doc)
    if (flagBool(args, 'sections')) detail.sections = listSections(doc)
    if (flagBool(args, 'fields')) detail.fields = listFields(doc)
    if (flagBool(args, 'notes')) detail.notes = listNotes(doc)
    return { summary: `${basename(path)}: ${count} blocks`, detail }
  } finally {
    closeDocument(doc)
  }
}

async function check(file: string | undefined, ctx: CommandContext): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const doc = await openDocument(readInput(path))
  try {
    return checkResult(path, checkDocument(doc), DOCS_CHECKS)
  } finally {
    closeDocument(doc)
  }
}

function trackFlag(args: Args): string | undefined {
  if (!flagBool(args, 'track')) return undefined
  return flagString(args, 'author')?.trim() || 'AI Assistant'
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
    throw new CliError(
      EXIT.usage,
      `${source}: not valid JSON (${(err as Error).message})`,
      undefined,
      { reason: 'invalid_json' },
    )
  }
  const ops = Array.isArray(parsed)
    ? parsed
    : parsed && typeof parsed === 'object' && Array.isArray((parsed as { ops?: unknown }).ops)
      ? (parsed as { ops: unknown[] }).ops
      : null
  if (!ops?.length)
    throw new CliError(EXIT.usage, `${source}: expected a non-empty array of ops`, undefined, {
      reason: 'invalid_argument',
    })
  const dryRun = flagBool(args, 'dry-run')
  const mode = batchMode(args)
  const trackAuthor = trackFlag(args)
  const doc = await openDocument(readInput(path))
  try {
    let outcome: DocOpsOutcome
    try {
      outcome = await applyDocOps(doc, ops as Record<string, unknown>[], {
        ctx,
        mode,
        baseDir: source === 'stdin' ? undefined : dirname(source),
        ...(trackAuthor ? { trackAuthor } : {}),
      })
    } catch (err) {
      throw err instanceof CliError ? failedBatch(err, ops.length) : err
    }
    const { results, failures } = outcome
    const counts = batchCounts(ops.length, results.length, failures.length)
    if (!results.length) {
      const first = failures[0]!
      throw new CliError(
        EXIT.usage,
        `no ops were applied: op ${first.index} (${first.op}) rejected: ${first.error}`,
        { failures, batch: counts },
        { reason: first.reason, suggestion: opSuggestion(first, 'docs') },
      )
    }
    if (dryRun) {
      return batchResult(
        {
          summary: `dry run: ${results.length} of ${ops.length} ops validated, nothing written`,
          detail: {
            plan: results.map((r) => `${r.index}: ${r.op}: ${r.output.split('\n')[0]}`),
          },
        },
        counts,
        failures,
      )
    }
    const output = resolveOutput(flagString(args, 'out'), ctx, {
      fallback: path,
      force: flagBool(args, 'force'),
      // in-place edits overwrite by design; --out onto another existing file needs --force
      fresh: flagString(args, 'out') !== undefined,
    })
    writeOutput(output, await saveDocument(doc))
    return batchResult(
      {
        summary: `applied ${results.length} of ${ops.length} ops to ${basename(output)}`,
        outputPath: output,
        detail: { results, blocks: doc.editor.state.doc.childCount },
      },
      counts,
      failures,
    )
  } finally {
    closeDocument(doc)
  }
}
