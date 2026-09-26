import {
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { isAbsolute, join, resolve } from 'node:path'
import { z } from 'zod'
import { assertAllowed } from '../fs'
import { CliError, EXIT, toJsonError, type JsonError } from '../result'
import { runJson, type McpContext, type Outcome } from './run'
import type { ZodShape } from './tools'

/**
 * The staged deck flow as tools. The deck directory is the state: style.md,
 * outline.json and pages/NN.json are exactly what `genoffice slides check` and
 * `genoffice create --spec <dir>` read, so a deck started here can be finished
 * from the command line and vice versa. Each tool refuses to run before the
 * stage it depends on exists, pages are written in outline order, and a page
 * that fails its check never replaces the file on disk.
 */
export interface DeckTool {
  name: string
  description: string
  shape: ZodShape
  readOnly?: boolean
  run(args: Record<string, unknown>, ctx: McpContext): Promise<Outcome>
}

/** Path policy errors from the deck tools come back as the same JSON error a command would print. */
export async function runDeckTool(
  tool: DeckTool,
  args: Record<string, unknown>,
  ctx: McpContext,
): Promise<Outcome> {
  try {
    return await tool.run(args, ctx)
  } catch (err) {
    const error =
      err instanceof CliError
        ? err
        : new CliError(EXIT.conversion, err instanceof Error ? err.message : String(err))
    return { error: toJsonError('deck', error) }
  }
}

const jsonObject = z.record(z.string(), z.unknown())

const DIR = z
  .string()
  .describe(
    'the deck folder (created if missing); style.md, outline.json, pages/ and deck.pptx live here',
  )

export const DECK_TOOLS: DeckTool[] = [
  {
    name: 'deck_start',
    description:
      'Stage 1 of a staged deck, for a longer or design-sensitive presentation a person will present (a short deck of up to about 5 slides with concrete content is one create_pptx call instead): write the style sheet and the outline, check the outline, and get the design and spec guides plus the configured cloud capabilities back in one call. Call once per deck, before any deck_page. Fix every outline error it reports (call again with the corrected outline) before writing pages. The result lists the pages to write, one deck_page call each.',
    shape: {
      dir: DIR,
      style: z
        .string()
        .describe(
          'style.md: the deck style sheet in Markdown, naming the palette as #RRGGBB colors, the fonts, the layout variants and one sentence of style; pages are checked against these colors',
        ),
      outline: jsonObject.describe(
        'the outline object: { "core_hook", "pages": [{ "title", "type": cover|content|data|closing, "layout", "brief" with the real figures, "image_queries": [] }] } (format: guide slides spec)',
      ),
    },
    async run(args, ctx) {
      const dir = deckDir(args.dir, ctx)
      mkdirSync(join(dir, 'pages'), { recursive: true })
      writeFileSync(join(dir, 'style.md'), String(args.style))
      writeFileSync(join(dir, 'outline.json'), JSON.stringify(args.outline, null, 2))
      const check = await runJson(['slides', 'check', join(dir, 'outline.json')], ctx)
      if (check.error) return check
      const [design, spec, caps] = await Promise.all([
        runJson(['guide', 'slides', 'design'], ctx),
        runJson(['guide', 'slides', 'spec'], ctx),
        runJson(['capabilities'], ctx),
      ])
      const pages = outlinePages(dir)
      return {
        ok: {
          ...check.ok,
          detail: {
            ...check.ok.detail,
            deck: dir,
            pages: pages.map((title, index) => ({ index, title, file: pageFile(dir, index) })),
            guides: { design: design.ok?.summary ?? null, spec: spec.ok?.summary ?? null },
            capabilities: caps.ok?.detail ?? caps.error?.message ?? null,
            next: `write page 0 ("${pages[0] ?? ''}") with deck_page; one page per call, in order, re-reading style.md's palette and the page's outline entry each time`,
          },
        },
      }
    },
  },
  {
    name: 'deck_page',
    description:
      'Stage 2: one page spec for outline entry `index`, checked against that entry and the style sheet before it is kept; a page that fails the check is not written. Exactly one page per call, in outline order (the next missing page, or an earlier one to rewrite). The result lists audit, outline and off-palette findings: fix them (call again with the same index) until all three are empty, then the next page. Refuses to run before deck_start.',
    shape: {
      dir: DIR,
      index: z.number().int().nonnegative().describe('0-based outline page index'),
      page: jsonObject.describe(
        'the one-page spec object { "title", "type", "layout", "elements": [...] } echoing the outline entry\'s title, type and layout (format: guide slides spec)',
      ),
    },
    async run(args, ctx) {
      const dir = deckDir(args.dir, ctx)
      const started = requireOutline(dir)
      if (started) return started
      const pages = outlinePages(dir)
      const index = Number(args.index)
      const range = pageInRange(index, pages.length)
      if (range) return range
      const before = missingPages(dir, pages.length)
      const next = before[0]
      if (next !== undefined && index > next) {
        return fail(
          'invalid_argument',
          `page ${index} comes after missing page ${next}; pages are written in outline order`,
          { missing_pages: before },
          `write page ${next} ("${pages[next] ?? ''}") first`,
        )
      }
      const check = await checkPending(dir, index, args.page, ctx)
      const missing = missingPages(dir, pages.length)
      const progress = {
        deck: dir,
        page: index,
        file: pageFile(dir, index),
        missing_pages: missing,
      }
      if (check.error) {
        return {
          error: {
            ...check.error,
            detail: { ...check.error.detail, ...progress, kept: false },
            suggestion:
              check.error.suggestion ?? 'fix the page and call deck_page again with the same index',
          },
        }
      }
      const then = missing.length
        ? `write page ${missing[0]} ("${pages[missing[0]!] ?? ''}") with deck_page`
        : 'every page has a file: build with deck_build'
      return { ok: { ...check.ok, detail: { ...check.ok.detail, ...progress, next: then } } }
    },
  },
  {
    name: 'deck_build',
    description:
      'Stage 3: build the .pptx from the checked pages. Refuses while an outline page has no file or a page disagrees with its entry. Then look at the result with slides_render and slides_audit; fix a page with deck_replace. At most two fix rounds.',
    shape: {
      dir: DIR,
      out: z
        .string()
        .optional()
        .describe('output .pptx path (default: <dir>/deck.pptx, overwritten on rebuild)'),
    },
    async run(args, ctx) {
      const dir = deckDir(args.dir, ctx)
      const started = requireOutline(dir)
      if (started) return started
      const missing = missingPages(dir, outlinePages(dir).length)
      if (missing.length) {
        return fail(
          'missing_argument',
          `pages ${missing.join(', ')} have no file yet`,
          { missing_pages: missing },
          'write them with deck_page before building',
        )
      }
      const out = args.out
        ? assertAllowed(resolvePath(String(args.out), ctx), ctx.env, 'write')
        : join(dir, 'deck.pptx')
      const argv = [
        'create',
        '--type',
        'pptx',
        '--spec',
        join(dir, 'pages'),
        '--outline',
        join(dir, 'outline.json'),
        '--out',
        out,
      ]
      if (!args.out) argv.push('--force')
      const built = await runJson(argv, ctx)
      if (built.error) return built
      return {
        ok: {
          ...built.ok,
          detail: {
            ...built.ok.detail,
            deck: dir,
            next: `slides_render ${out} to look at every page, slides_audit for geometry; deck_replace rebuilds one page from a corrected spec`,
          },
        },
      }
    },
  },
  {
    name: 'deck_replace',
    description:
      'Fix round: rewrite one page spec, check it against its outline entry and the style sheet, and rebuild only that slide of the built deck; the other slides keep their ids and content. When the check fails, the previous page file and the deck stay as they were.',
    shape: {
      dir: DIR,
      index: z.number().int().nonnegative().describe('0-based page (slide) to rebuild'),
      page: jsonObject.describe('the corrected one-page spec object'),
      pptx: z.string().optional().describe('the built deck (default: <dir>/deck.pptx)'),
    },
    async run(args, ctx) {
      const dir = deckDir(args.dir, ctx)
      const started = requireOutline(dir)
      if (started) return started
      const pages = outlinePages(dir)
      const index = Number(args.index)
      const range = pageInRange(index, pages.length)
      if (range) return range
      const pptx = args.pptx
        ? assertAllowed(resolvePath(String(args.pptx), ctx), ctx.env, 'read')
        : join(dir, 'deck.pptx')
      if (!existsSync(pptx)) {
        return fail('file_not_found', `no built deck at ${pptx}`, undefined, 'run deck_build first')
      }
      const check = await checkPending(dir, index, args.page, ctx)
      if (check.error) {
        return { error: { ...check.error, detail: { ...check.error.detail, kept: false } } }
      }
      const replaced = await runJson(
        ['slides', 'replace', pptx, '--slide', String(index), '--spec', pageFile(dir, index)],
        ctx,
      )
      if (replaced.error) return replaced
      return {
        ok: {
          ...replaced.ok,
          detail: {
            ...replaced.ok.detail,
            check: check.ok.detail,
            next: `slides_render ${pptx} with slide ${index} to confirm`,
          },
        },
      }
    },
  },
]

/**
 * The page is checked from a `.pending` file beside the real one (so
 * `slides check` finds outline.json and style.md as usual) and only renamed
 * into place when it passes; a failing spec leaves the previous file alone.
 */
async function checkPending(
  dir: string,
  index: number,
  page: unknown,
  ctx: McpContext,
): Promise<Outcome> {
  const file = pageFile(dir, index)
  const pending = `${file}.pending`
  writeFileSync(pending, JSON.stringify(page, null, 2))
  try {
    const check = await runJson(['slides', 'check', pending, '--page', String(index)], ctx)
    if (check.ok) renameSync(pending, file)
    return check
  } finally {
    rmSync(pending, { force: true })
  }
}

function resolvePath(path: string, ctx: McpContext): string {
  return isAbsolute(path) ? path : resolve(ctx.cwd, path)
}

/** The deck folder, inside GENOFFICE_ALLOWED_ROOTS when that is set: everything the tools write lives under it. */
function deckDir(raw: unknown, ctx: McpContext): string {
  return assertAllowed(resolvePath(String(raw), ctx), ctx.env, 'write')
}

function pageFile(dir: string, index: number): string {
  return join(dir, 'pages', `${String(index + 1).padStart(2, '0')}.json`)
}

function pageInRange(index: number, count: number): Outcome | null {
  if (index < count) return null
  return fail('out_of_range', `the outline has ${count} pages; index ${index} has no entry`, {
    valid_range: [0, count - 1],
  })
}

function outlinePages(dir: string): string[] {
  try {
    const raw = JSON.parse(readFileSync(join(dir, 'outline.json'), 'utf-8')) as {
      pages?: { title?: unknown }[]
    }
    return (raw.pages ?? []).map((p) => (typeof p?.title === 'string' ? p.title : ''))
  } catch {
    return []
  }
}

/** Outline indexes whose page file is missing; extra files beyond the outline are the build's problem. */
function missingPages(dir: string, count: number): number[] {
  const present = new Set(
    existsSync(join(dir, 'pages'))
      ? readdirSync(join(dir, 'pages')).filter((f) => f.endsWith('.json'))
      : [],
  )
  const missing: number[] = []
  for (let i = 0; i < count; i++) {
    if (!present.has(`${String(i + 1).padStart(2, '0')}.json`)) missing.push(i)
  }
  return missing
}

function requireOutline(dir: string): Outcome | null {
  if (existsSync(join(dir, 'outline.json'))) return null
  return fail(
    'missing_argument',
    `${dir} has no outline.json`,
    undefined,
    'call deck_start with the style sheet and the outline first',
  )
}

function fail(
  reason: JsonError['error'],
  message: string,
  detail?: Record<string, unknown>,
  suggestion?: string,
): Outcome {
  return {
    error: {
      status: 'error',
      command: 'deck',
      code: 1,
      error: reason,
      message,
      ...(suggestion ? { suggestion } : {}),
      ...(detail ? { detail } : {}),
    },
  }
}
