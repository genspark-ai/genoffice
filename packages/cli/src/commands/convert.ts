import { basename, dirname, extname, join } from 'node:path'
import { PdfLoadError } from '@genoffice/pdf2docx'
import { flagBool, flagString, type ParsedArgs } from '../args'
import { csvToXlsx } from '../formats/csv'
import { convertPdf, type PdfTarget } from '../formats/pdf'
import { convertLegacyWorkbook, sheetToCsv } from '../formats/xlsx'
import { exportViaApp, type AppExportTarget } from '../formats/app-export'
import { htmlToMarkdown, markdownToDocx, markdownToHtml } from '../formats/markdown'
import { closeDocument, documentHtml, openDocument } from '../formats/docx'
import { extension, readInput, resolveInput, resolveOutput, writeOutput } from '../fs'
import type { CommandContext, CommandDef } from '../registry'
import { CliError, EXIT } from '../result'

/** Conversions that run in this process. */
const NODE_ROUTES: Record<string, readonly string[]> = {
  pdf: ['docx', 'pptx', 'xlsx'],
  csv: ['xlsx'],
  xls: ['xlsx'],
  xlsb: ['xlsx'],
  ods: ['xlsx'],
  md: ['docx', 'html'],
  markdown: ['docx', 'html'],
  docx: ['md'],
  xlsx: ['csv'],
  xlsm: ['csv'],
}

/**
 * Conversions the GenOffice binary runs for us in its hidden headless-export
 * mode: anything that needs an app renderer (page layout for pdf, the Word
 * editor's HTML export, html2docx). Mirrors HEADLESS_TARGETS in the shell.
 */
const APP_ROUTES: Record<string, readonly AppExportTarget[]> = {
  csv: ['pdf'],
  xls: ['pdf'],
  md: ['pdf'],
  markdown: ['pdf'],
  docx: ['pdf', 'html'],
  xlsx: ['pdf'],
  xlsm: ['pdf'],
  pptx: ['pdf'],
  html: ['pdf', 'docx'],
  htm: ['pdf', 'docx'],
}

const ROUTES: Record<string, readonly string[]> = Object.fromEntries(
  [...new Set([...Object.keys(NODE_ROUTES), ...Object.keys(APP_ROUTES)])].map((from) => [
    from,
    [...(NODE_ROUTES[from] ?? []), ...(APP_ROUTES[from] ?? [])],
  ]),
)

function appTarget(from: string, to: string): AppExportTarget | null {
  const target = APP_ROUTES[from]?.find((t) => t === to)
  return target ?? null
}

export const convertCommand: CommandDef = {
  name: 'convert',
  summary: 'Convert a document to another format using the GenOffice engines.',
  usage: 'convert <file> --to <format> [--out <path>] [--force] [--password <pw>] [--sheet <name>]',
  options: [
    { name: 'to', value: 'format', description: 'target format: ' + describeRoutes() },
    { name: 'out', value: 'path', description: 'output file (default: same name, new extension)' },
    { name: 'force', description: 'overwrite an existing output file' },
    { name: 'password', value: 'pw', description: 'password for an encrypted PDF' },
    {
      name: 'sheet',
      value: 'name',
      description: 'xlsx→csv: worksheet to export (default: the active one)',
    },
  ],
  async run(args, ctx) {
    const input = resolveInput(args.positionals[0], ctx)
    const from = extension(input)
    const to = flagString(args, 'to')?.toLowerCase()
    if (!to)
      throw new CliError(EXIT.usage, 'missing --to <format>', undefined, {
        reason: 'missing_argument',
      })
    const targets = ROUTES[from]
    if (!targets?.includes(to)) {
      throw new CliError(
        EXIT.usage,
        `cannot convert .${from} to .${to}`,
        { supported: describeRoutes() },
        { reason: 'unsupported', suggestion: 'pick a route from detail.supported' },
      )
    }
    // before run(): the sidecar and app routes write the output themselves
    const output = resolveOutput(flagString(args, 'out'), ctx, {
      fallback: join(dirname(input), `${basename(input, extname(input))}.${to}`),
      force: flagBool(args, 'force'),
      fresh: true,
    })
    const result = await run(input, from, to, output, args, ctx)
    if (result.bytes) writeOutput(output, result.bytes)
    return {
      summary: `converted ${basename(input)} → ${basename(output)}`,
      outputPath: output,
      detail: result.detail,
    }
  },
}

interface Produced {
  bytes?: Uint8Array
  detail: Record<string, unknown>
}

async function run(
  input: string,
  from: string,
  to: string,
  output: string,
  args: ParsedArgs,
  ctx: CommandContext,
): Promise<Produced> {
  const password = flagString(args, 'password')
  const viaApp = appTarget(from, to)
  if (viaApp) {
    const r = await exportViaApp(input, viaApp, output, { env: ctx.env, log: ctx.log })
    return { detail: { via: 'genoffice --headless-export', summary: r.summary } }
  }
  if (from === 'pdf') {
    try {
      const r = await convertPdf(readInput(input), to as PdfTarget, {
        password,
        onProgress: (page, total) => ctx.log(`page ${page}/${total}`),
      })
      return { bytes: r.bytes, detail: { pages: r.pages, warnings: r.warnings } }
    } catch (err) {
      if (err instanceof PdfLoadError) {
        const code = err.code === 'password-required' ? EXIT.file : EXIT.conversion
        throw new CliError(code, err.message, { reason: err.code })
      }
      throw err
    }
  }
  if (from === 'md' || from === 'markdown') {
    const text = Buffer.from(readInput(input)).toString('utf-8')
    const title = basename(input, extname(input))
    if (to === 'html') {
      return { bytes: Buffer.from(await markdownToHtml(text, title), 'utf-8'), detail: { title } }
    }
    return { bytes: await markdownToDocx(text, { file: input, ctx }), detail: { title } }
  }
  if (from === 'csv') {
    const sheet = basename(input, extname(input)).slice(0, 31) || 'Sheet1'
    return { bytes: await csvToXlsx(readInput(input), sheet), detail: { sheet } }
  }
  if (to === 'csv') {
    const { csv, ...detail } = await sheetToCsv(input, { sheet: flagString(args, 'sheet') })
    // UTF-8 BOM, as the app writes it, so Excel opens the file without a wizard
    return { bytes: Buffer.from('\ufeff' + csv, 'utf-8'), detail }
  }
  if (to === 'md') {
    const doc = await openDocument(readInput(input))
    try {
      const exported = documentHtml(doc)
      const markdown = await htmlToMarkdown(exported.html)
      if (exported.skipped.images > 0) {
        ctx.warn({
          code: 'images_dropped',
          message: `${exported.skipped.images} image(s) have no Markdown form and were dropped`,
          suggestion: 'use `genoffice docs read --html` when the images matter',
        })
      }
      return {
        bytes: Buffer.from(markdown.replace(/\n{3,}/g, '\n\n').trimEnd() + '\n', 'utf-8'),
        detail: {
          blocks: doc.editor.state.doc.childCount,
          skipped: exported.skipped,
          formulas_as_text: exported.formulasAsText,
        },
      }
    } finally {
      closeDocument(doc)
    }
  }
  const r = await convertLegacyWorkbook(input, output)
  return { detail: { sheets: r.sheets, cells: r.cells } }
}

function describeRoutes(): string {
  return Object.entries(ROUTES)
    .map(([from, tos]) => `${from}→${tos.join('/')}`)
    .join(', ')
}
