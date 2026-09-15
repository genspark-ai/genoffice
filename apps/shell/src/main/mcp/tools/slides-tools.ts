import { z } from 'zod'
import { outlineToOps, parsePptxOutline, type PptxSourceFormat } from '../pptx-outline'
import { createPptxViaCli } from '../headless-cli'
import { resolveOutputPath } from './document-tools'
import { generateExtension } from './formats'
import type { FamilyDriver, SessionHost } from './session-tools'
import type { McpToolDefinition } from '../mcp-server'
import type { CliRunner } from '../cli-runner'

/**
 * Slides (pptx) tool surface for the MCP server.
 *
 * Two paths, mirroring the docx tools:
 * - headless `create_pptx`: outline -> ops, handed to the bundled `genoffice`
 *   CLI (`create --type pptx --ops`), gated behind the "background generation"
 *   setting. No deck-building code lives here;
 * - a visible deck session: the tools drive a real slides tab the user watches.
 *   Slides sessions live in the slides main process, so the control talks to
 *   main-process state directly (no renderer bridge) — see
 *   `apps/shell/src/main/mcp/slides-bridge.ts`.
 */

export interface SlidesToolDeps {
  /** directory generated files land in when the caller gives no path */
  defaultSaveDir: () => string
  /** expose the headless create_pptx tool; default true — the shell passes the user's setting */
  background?: boolean
  /** visible-deck control; absent in headless/unit runs, which drops the session tools */
  slides?: SlidesControl
  /** the bundled genoffice CLI, used by the headless tool */
  cli?: CliRunner
}

/** transaction request for the visible deck (executor semantics, ops are EMU-space) */
export interface SlidesTxnRequest {
  ops: unknown[]
  isolation?: 'atomic' | 'per_op'
  dryRun?: boolean
}

/**
 * Visible-deck control implemented in the shell main process against the slides
 * main-process sessions (`slides-bridge.ts`).
 */
export interface SlidesControl {
  /** open a fresh blank slides tab; resolves to its webContents id */
  openBlankTab: () => Promise<number>
  /** apply one transaction to the tab's session (history + journal + live render included) */
  runTxn: (wcId: number, req: SlidesTxnRequest) => Promise<unknown>
  /** read the visible deck (slides/elements with ids and geometry) */
  readDeck: (wcId: number) => Promise<unknown>
  /** save the session's deck to an absolute path */
  saveDeck: (wcId: number, path: string, overwrite: boolean) => Promise<{ path: string }>
}

const PPTX_EXT = `.${generateExtension('pptx')}`

/** the headless tool: outline -> ops -> pptx bytes written straight to disk */
function createHeadlessPptxTool(deps: SlidesToolDeps): McpToolDefinition {
  return {
    name: 'create_pptx',
    description:
      'Create a PowerPoint .pptx file from an outline and save it to disk without opening the app UI. ' +
      'By default `outline` is Markdown where `# Title` starts each slide, `## text` adds a bold line, ' +
      '`- text` a bullet (indent two spaces per level) and `1. text` a numbered bullet. Use ' +
      'format:"json" to pass {"slides":[{"title":"...","bullets":["...","...",{"text":"...","level":1}]}]} instead. ' +
      'Returns the absolute path of the written file.',
    inputSchema: {
      title: z.string().describe('presentation title, used as the file name'),
      outline: z
        .string()
        .describe(
          'Markdown outline (default) or a JSON string of {"slides":[...]} when format is "json"',
        ),
      format: z.enum(['markdown', 'json']).optional().describe('outline format; default markdown'),
      path: z
        .string()
        .optional()
        .describe('absolute output path; default is a new file in the default save folder'),
      overwrite: z
        .boolean()
        .optional()
        .describe('allow replacing an existing file at `path`; default false'),
    },
    handler: async (args) => {
      const title = String(args.title ?? '').trim()
      if (!title) throw new Error('title must not be empty')
      if (typeof args.outline !== 'string') {
        throw new Error('outline must be a string (markdown by default, JSON with format:"json")')
      }
      if (!deps.cli) throw new Error('headless generation is not available in this build')
      const format: PptxSourceFormat = args.format === 'json' ? 'json' : 'markdown'
      const ops = outlineToOps(parsePptxOutline(format, args.outline))

      const targetPath = resolveOutputPath({
        defaultSaveDir: deps.defaultSaveDir,
        ext: PPTX_EXT,
        title,
        requestedPath: typeof args.path === 'string' ? args.path : undefined,
        overwrite: args.overwrite === true,
      })

      const { summary, outputPath } = await createPptxViaCli(deps.cli, {
        ops,
        out: targetPath,
        overwrite: args.overwrite === true,
      })
      return { path: outputPath, summary }
    },
  }
}

export function createSlidesTools(deps: SlidesToolDeps, host: SessionHost): McpToolDefinition[] {
  return [
    // headless generation is opt-in, same rule as create_docx
    ...(deps.background === false ? [] : [createHeadlessPptxTool(deps)]),
    ...createDeckContentTools(deps, host),
  ]
}

/**
 * The pptx session lifecycle as seen by the shared create_session / save_session
 * tools. The content tools below address the tab this driver opened.
 */
export function slidesDriver(slides: SlidesControl): FamilyDriver {
  return {
    family: 'pptx',
    openBlankTab: () => slides.openBlankTab(),
    save: (wcId, path, overwrite) => slides.saveDeck(wcId, path, overwrite),
  }
}

/**
 * Visible-deck content tools: edit the deck the shared session opened, so the
 * user watches each page take shape. The control mutates the main-process
 * session (the same one the app's own editing and built-in AI use), so every
 * step lands in the app's undo history and the tab re-renders live.
 *
 * Only registered when the shell wired a SlidesControl — headless/unit runs
 * keep the file-only surface.
 */
function createDeckContentTools(deps: SlidesToolDeps, host: SessionHost): McpToolDefinition[] {
  const slides = deps.slides
  if (!slides) return []

  const requireActive = (): number => host.require('pptx')

  return [
    {
      name: 'read_deck',
      description:
        'Read the visible presentation: every slide with its elements (ids, types, text, geometry) ' +
        'so you can target follow-up edits. Element ids and slide indexes are what apply_slide_ops ' +
        'addresses; geometry is document-space EMU (1 px = 9525 EMU on a standard 16:9 deck).',
      inputSchema: {},
      handler: async () => slides.readDeck(requireActive()),
    },
    {
      name: 'apply_slide_ops',
      description:
        'Apply slide editing ops to the visible presentation as one transaction (atomic by default, ' +
        'plan-validated with snapshot rollback; per_op isolation skips failures instead). Ops address ' +
        '{slide, el} and use document-space EMU. Common ops: setText, setFont, setParagraphFormat, ' +
        'addElement (textbox/shapes/lines), addTable, addChart, addBlankSlide, deleteSlide, moveSlide, ' +
        'duplicateSlide, deleteElement, setFill, setStroke, setTransform, setBackground, applyTheme, ' +
        'findReplace, setTableCell, tableStructure, setTransition, setNotes. A failing atomic batch ' +
        'changes nothing — fix the op named in the error and resend the whole batch. ' +
        'Use read_deck for ids and geometry first.',
      inputSchema: {
        ops: z.array(z.any()).describe('array of op objects (at most 50 per transaction)').max(50),
        isolation: z
          .enum(['atomic', 'per_op'])
          .optional()
          .describe('atomic (default) applies all-or-nothing; per_op applies independently'),
        dryRun: z
          .boolean()
          .optional()
          .describe('validate and plan the batch without changing the deck'),
      },
      handler: async (args) => {
        const wc = requireActive()
        const result = (await slides.runTxn(wc, {
          ops: (Array.isArray(args.ops) ? args.ops : []) as unknown[],
          isolation:
            args.isolation === 'per_op' || args.isolation === 'atomic' ? args.isolation : undefined,
          dryRun: args.dryRun === true,
        })) as { applied?: boolean; failures?: Array<{ index: number; error: string }> }
        // a rejected transaction is a tool-level error so the caller reacts to it;
        // dry-run results (plan + failures) stay a normal response
        if (result?.applied === false && args.dryRun !== true) {
          const details = (result.failures ?? []).map((f) => `[${f.index}] ${f.error}`).join('\n')
          throw new Error(details || 'the transaction could not be applied')
        }
        return result
      },
    },
  ]
}
