import { readFileSync, writeFileSync } from 'node:fs'
import { basename, dirname } from 'node:path'
import { flagBool, flagString } from '../args'
import {
  applyOps,
  describeDeck,
  inlineLocalFiles,
  openDeck,
  parseOps,
  saveDeck,
} from '../formats/pptx'
import { readInput, resolveInput, resolveOutput } from '../fs'
import { outputDirectory, parseScale, renderToPngs } from '../formats/render'
import {
  auditDeckBytes,
  checkPageSpec,
  findStageFiles,
  pageIndexOf,
  replaceSlideFromSpec,
  stageContext,
} from '../formats/slide-spec'
import { parseOutline } from '@genoffice/pipelines/slides'
import { readOpsInput } from '../ops-input'
import type { TxnResult } from '@genoffice/pptx-ops'
import type { CommandContext, CommandDef } from '../registry'
import { CliError, EXIT, type CommandResult } from '../result'
import { txnDetail, txnFailure } from './txn'

export const slidesCommand: CommandDef = {
  name: 'slides',
  summary:
    'Read, edit, audit or render a .pptx with the ops and layout audit the in-app AI uses; check an outline or page spec; rebuild one slide from its spec.',
  usage:
    'slides <read|apply|audit|render|replace> <file.pptx> [options] | slides check <outline.json|page.json>',
  options: [
    {
      name: 'slide',
      value: 'n',
      description: 'read/audit/render: only this 0-based slide; replace: the slide to rebuild',
    },
    {
      name: 'full',
      description: 'read: whole text and every table row instead of previews, plus speaker notes',
    },
    { name: 'ops', value: 'file', description: 'apply: JSON ops file, or "-" for stdin' },
    { name: 'spec', value: 'file', description: 'replace: the one-page spec file to build' },
    {
      name: 'outline',
      value: 'file',
      description:
        'check <page.json> / replace: the outline the page is checked against (default: outline.json beside the page file or one folder up)',
    },
    { name: 'dry-run', description: 'apply: validate and print the plan without writing' },
    { name: 'isolation', value: 'mode', description: 'apply: atomic (default) or per_op' },
    { name: 'out', value: 'path', description: 'apply: write here instead of in place' },
    {
      name: 'force',
      description:
        'apply: overwrite an existing --out file, or write while GenOffice has the file open',
    },
    {
      name: 'scale',
      value: 'n',
      description: 'render: pixels per point (default 1 = 960x540 for a 16:9 deck; 2 doubles it)',
    },
  ],
  async run(args, ctx) {
    const [verb, file] = args.positionals
    switch (verb) {
      case 'read':
        return read(file, args, ctx)
      case 'apply':
        return apply(file, args, ctx)
      case 'audit':
        return audit(file, args, ctx)
      case 'render':
        return render(file, args, ctx)
      case 'check':
        return check(file, args, ctx)
      case 'replace':
        return replace(file, args, ctx)
      default:
        throw new CliError(
          EXIT.usage,
          'expected "slides read|apply|audit|render|replace <file.pptx>" or "slides check <outline.json|page.json>"',
        )
    }
  },
}

function slideFlag(args: Parameters<CommandDef['run']>[0], count: number): number | undefined {
  const index = slideFlagValue(args)
  if (index === undefined) return undefined
  if (index >= count) throw new CliError(EXIT.usage, `--slide out of range (0-${count - 1})`)
  return index
}

/**
 * `--slide` as a 0-based index, before any file is opened: a malformed value
 * (a shell loop that pasted "2 03") is reported as such, not as a missing file
 * or an out-of-range slide.
 */
function slideFlagValue(args: Parameters<CommandDef['run']>[0]): number | undefined {
  const only = flagString(args, 'slide')
  if (only === undefined) return undefined
  const index = Number(only)
  if (!/^\d+$/.test(only.trim()) || !Number.isInteger(index) || index < 0) {
    throw new CliError(
      EXIT.usage,
      `--slide must be one 0-based slide index (e.g. --slide 2), got "${only}"`,
    )
  }
  return index
}

async function read(
  file: string | undefined,
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const opened = await openDeck(readInput(path))
  const index = slideFlag(args, opened.deck.slides.length)
  const deck = describeDeck(opened, index, flagBool(args, 'full'))
  return {
    summary: `${basename(path)}: ${deck.slides} slides`,
    detail: {
      ...deck,
      units: 'EMU (914400 per inch); slide ids s_<n> and element ids e_* are durable op targets',
    },
  }
}

async function apply(
  file: string | undefined,
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const isolation = flagString(args, 'isolation') ?? 'atomic'
  if (isolation !== 'atomic' && isolation !== 'per_op') {
    throw new CliError(EXIT.usage, '--isolation must be atomic or per_op')
  }
  const { text, source } = readOpsInput(args, ctx)
  const ops = inlineLocalFiles(parseOps(text, source), ctx)
  const opened = await openDeck(readInput(path))
  const dryRun = flagBool(args, 'dry-run')
  // One transaction per op, as `create` does: the executor validates a batch
  // against the pre-transaction deck, so a slide added by op 0 would not exist
  // for op 1 in a single transaction. Atomic still means nothing is written
  // unless every op applied; per_op keeps what applied and reports the rest.
  const records: NonNullable<TxnResult['records']> = []
  const failures: NonNullable<TxnResult['failures']> = []
  // dry-run plan lines keep the caller's op index, so they line up with failures
  const plan: string[] = []
  for (const [index, op] of ops.entries()) {
    // the op's own isolation mode, so a per_op failure is not worded as an atomic rollback
    const r = applyOps(opened, [op], { isolation })
    if (r.applied) {
      records.push(...(r.records ?? []))
      const slide = r.records?.[0]?.slideId
      plan.push(`${index}: ${op.op}${slide ? ` on ${slide}` : ''}`)
      continue
    }
    if (isolation === 'atomic') throw txnFailure(r, index)
    failures.push(...(r.failures ?? []).map((f) => ({ ...f, index })))
  }
  const r: TxnResult = { applied: records.length > 0, records, failures }
  if (dryRun) {
    // the batch ran on the in-memory deck only: nothing is applied to the file
    r.applied = false
    r.dryRun = true
    r.plan = plan
    return {
      summary: `dry run: ${records.length} of ${ops.length} ops validated`,
      detail: txnDetail(r),
    }
  }
  if (!r.applied) throw txnFailure(r)
  const output = resolveOutput(flagString(args, 'out'), ctx, {
    fallback: path,
    force: flagBool(args, 'force'),
    // in-place edits overwrite by design; --out onto another existing file needs --force
    fresh: flagString(args, 'out') !== undefined,
  })
  writeFileSync(output, await saveDeck(opened))
  return {
    summary: `applied ${r.records?.length ?? 0} ops to ${basename(output)}`,
    outputPath: output,
    detail: txnDetail(r),
  }
}

/** Geometry audit per slide: out of bounds, text overflow, overlapping content, with durable ids. */
async function audit(
  file: string | undefined,
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const bytes = readInput(path)
  const opened = await openDeck(bytes)
  const index = slideFlag(args, opened.deck.slides.length)
  const pages = await auditDeckBytes(bytes, index)
  const failing = pages.filter((p) => p.issues.length > 0)
  const total = failing.reduce((n, p) => n + p.issues.length, 0)
  return {
    summary: failing.length
      ? `${basename(path)}: ${total} issue(s) on ${failing.length} of ${pages.length} slide(s)`
      : `${basename(path)}: ${pages.length} slide(s), no layout issues`,
    detail: {
      slides: pages,
      metrics: 'heuristic glyph widths; overflow figures are approximate',
      ids: 'element ids match `slides read` / `slides apply` targets',
    },
  }
}

/**
 * Findings for one deck-flow artifact: an outline (core hook + page plan) or a
 * single page spec, told apart by their shape. Errors exit 1 so a script can
 * gate on them; warnings are advice.
 */
async function check(
  file: string | undefined,
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const text = readFileSync(path, 'utf-8')
  const dir = dirname(path)
  if (looksLikeOutline(text)) {
    const r = parseOutline(text)
    if (!r.ok) throw new CliError(EXIT.usage, `${basename(path)}: ${r.error}`)
    const errors = r.issues.filter((i) => i.level === 'error')
    const style = findStageFiles(dir, ctx).style
    const detail = {
      kind: 'outline',
      pages: r.outline.pages.length,
      issues: r.issues,
      notes: style
        ? []
        : ['no style.md beside the outline: write the style sheet before the pages'],
      next: 'write pages/01.json … one file per outline page, running `genoffice slides check <page.json>` on each',
    }
    if (errors.length) {
      throw new CliError(
        EXIT.usage,
        `${basename(path)}: ${errors.length} error(s) in the outline`,
        detail,
      )
    }
    return {
      summary: `${basename(path)}: ${r.outline.pages.length} pages, ${r.issues.length ? `${r.issues.length} warning(s)` : 'no findings'}`,
      detail,
    }
  }
  const explicit = flagString(args, 'outline')
  const stage = stageContext(dir, ctx, explicit ? resolveInput(explicit, ctx) : undefined)
  const page = pageIndexOf(path)
  const r = await checkPageSpec(text, path, ctx, { context: stage, page })
  const outline = r.stage?.outline ?? []
  const offPalette = r.stage?.offPalette ?? []
  const count =
    r.warnings.length + r.imageFailures.length + r.audit.length + outline.length + offPalette.length
  const detail = {
    kind: 'page',
    page,
    warnings: r.warnings,
    imageFailures: r.imageFailures,
    audit: r.audit,
    outline: stage.outline ? { file: stage.outline.file, findings: outline } : null,
    style: stage.style
      ? {
          file: stage.style.file,
          offPalette,
          ...(offPalette.length
            ? {
                note: 'colors the style sheet does not name; add them to style.md or use the palette',
              }
            : {}),
        }
      : null,
    notes: stage.notes,
    metrics: 'heuristic glyph widths; overflow figures are approximate',
  }
  const errors = outline.filter((o) => o.level === 'error')
  if (errors.length) {
    throw new CliError(
      EXIT.usage,
      `${basename(path)}: ${errors.length} error(s) against ${basename(stage.outline!.file)} pages[${page}]`,
      detail,
    )
  }
  return {
    summary: count
      ? `${basename(path)}: ${count} finding(s)`
      : `${basename(path)}: page builds clean`,
    detail,
  }
}

function looksLikeOutline(text: string): boolean {
  try {
    const v = JSON.parse(text) as Record<string, unknown> | null
    if (!v || typeof v !== 'object' || Array.isArray(v) || !Array.isArray(v.pages)) return false
    // a multi-page deck spec also has "pages", but its entries carry elements, not briefs
    return !v.pages.some((p) => p && typeof p === 'object' && 'elements' in (p as object))
  } catch {
    return false
  }
}

/** One page rebuilt from its spec in place of slide --slide; the other slides keep their ids and content. */
async function replace(
  file: string | undefined,
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<CommandResult> {
  if (slideFlagValue(args) === undefined) throw new CliError(EXIT.usage, 'missing --slide <n>')
  const path = resolveInput(file, ctx)
  const specFile = flagString(args, 'spec')
  if (!specFile) throw new CliError(EXIT.usage, 'missing --spec <page.json>')
  const specPath = resolveInput(specFile, ctx)
  const opened = await openDeck(readInput(path))
  const index = slideFlag(args, opened.deck.slides.length)!
  const explicit = flagString(args, 'outline')
  const stage = stageContext(
    dirname(specPath),
    ctx,
    explicit ? resolveInput(explicit, ctx) : undefined,
  )
  const r = await replaceSlideFromSpec(
    opened,
    index,
    readFileSync(specPath, 'utf-8'),
    specPath,
    ctx,
    stage,
  )
  const output = resolveOutput(flagString(args, 'out'), ctx, {
    fallback: path,
    force: flagBool(args, 'force'),
    fresh: flagString(args, 'out') !== undefined,
  })
  writeFileSync(output, await saveDeck(opened))
  return {
    summary: `replaced slide ${index} of ${basename(output)} from ${basename(specPath)}`,
    outputPath: output,
    detail: {
      slide: index,
      issues: r.warnings,
      imageFailures: r.imageFailures,
      outline: stage.outline
        ? { file: stage.outline.file, findings: r.stage?.outline ?? [] }
        : null,
      offPalette: r.stage?.offPalette ?? [],
      notes: stage.notes,
      next: 'element ids on the rebuilt slide are new: run `genoffice slides read` before targeting them',
    },
  }
}

/** One PNG per slide, through the app's PDF export (hidden GenOffice process) and pdfium. */
async function render(
  file: string | undefined,
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const outDir = outputDirectory(flagString(args, 'out'), ctx)
  const scale = parseScale(flagString(args, 'scale'))
  const opened = await openDeck(readInput(path))
  const only = slideFlag(args, opened.deck.slides.length)
  const files = await renderToPngs(path, ctx, { outDir, scale, only, log: ctx.log })
  return {
    summary: `rendered ${files.length} slide(s) to ${outDir}`,
    outputPath: outDir,
    detail: { files: files.map(({ page, ...f }) => ({ slide: page, ...f })) },
  }
}
