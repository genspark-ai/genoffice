import { existsSync } from 'node:fs'
import { basename, extname, isAbsolute, join } from 'node:path'
import { z } from 'zod'
import {
  convertViaCli,
  createFileViaCli,
  docsCheckViaCli,
  mergeTemplateViaCli,
  readDocxTextViaCli,
  readDocxPropertiesViaCli,
} from '../headless-cli'
import type { McpToolDefinition } from '../mcp-server'
import { capabilityReport, generateExtension } from './formats'
import type { FamilyDriver, SessionHost, TargetResolver } from './session-tools'
import { contentTarget } from './session-tools'
import type { CliRunner } from '../cli-runner'

/**
 * Docx MCP tool surface.
 *
 * Headless generation/reading (create_docx, read_docx) delegates to the
 * bundled `genoffice` CLI — the same engines the app ships — instead of
 * reimplementing conversion here (see ../headless-cli.ts, ../cli-runner.ts).
 * The visible document session is the part the CLI cannot do: driving the
 * editor tab the user has open. Dependencies are injected so this module stays
 * free of Electron and is unit-testable in a plain Node environment.
 */

export interface DocToolDeps {
  version: string
  /** directory generated files land in when the caller gives no path */
  defaultSaveDir: () => string
  /** expose the headless create_docx tool (generation without opening the UI);
   *  default true — the shell passes the user's "background generation" setting */
  background?: boolean
  /** open a file in the GenOffice UI; wired in M4 (optional in tests) */
  openInTab?: (filePath: string) => Promise<void> | void
  /** visible-editor control for the MCP-driven document session (optional in tests) */
  docs?: DocsControl
  /** extra formats other tool families can generate (reported by get_app_info) */
  extraFormats?: string[]
  /** the bundled genoffice CLI, used by the headless tools */
  cli?: CliRunner
  /**
   * Resolve a caller-supplied document reference (tab id or path) to the
   * webContents of that open tab, so the content tools can edit a document the
   * *user* has open rather than only the session's own blank one. Absent in
   * headless/unit runs, where a `document` argument is refused.
   */
  resolveTarget?: TargetResolver
}

/** editor commands the docs renderer bridge understands (see docs shared/ipc.ts) */
export type McpEditorCommandName =
  | 'insert_content'
  | 'replace_blocks'
  | 'apply_ops'
  | 'insert_picture'
  | 'read_document'
  | 'read_blocks'
  | 'save_document'

/**
 * Visible-editor control: opens a docs tab and pushes editor commands into it,
 * so an external agent builds/edits a document the user can watch instead of
 * writing bytes behind the UI. Implemented in the shell main process
 * (`apps/shell/src/main/mcp/docs-bridge.ts`).
 */
export interface DocsControl {
  /** open a fresh blank docs tab; resolves to its webContents id */
  openBlankTab: () => Promise<number>
  /** run one editor command in that tab and resolve its result */
  runCommand: (wcId: number, command: McpEditorCommandName, payload: unknown) => Promise<unknown>
}

const DOCX_EXT = `.${generateExtension('docx')}`

/** sanitize a title into a file base name (mirrors docs' sanitizeAiDocFileBase) */
export function sanitizeFileBase(title: string): string {
  const cleaned = String(title ?? '')
    // eslint-disable-next-line no-control-regex -- control characters are rejected on purpose
    .replace(/[/\\:*?"<>|\u0000-\u001f]/g, '_')
    .trim()
    .slice(0, 80)
    .trim()
    .replace(/^\.+|\.+$/g, '')
    .trim()
  return cleaned && cleaned !== '.' && cleaned !== '..' ? cleaned : 'Untitled'
}

/** first free path for fileName inside dir: name.docx, name-2.docx, … */
export function uniquePathIn(dir: string, fileName: string): string {
  const ext = extname(fileName)
  const base = fileName.slice(0, fileName.length - ext.length)
  let candidate = join(dir, fileName)
  for (let i = 2; existsSync(candidate); i++) candidate = join(dir, `${base}-${i}${ext}`)
  return candidate
}

/**
 * Shared output-path policy for every generating tool: sanitize the title into a
 * file base name, land in the default save folder when no path is given, require
 * absolute paths otherwise, append the extension when missing, and refuse to
 * clobber an existing file unless overwrite is set.
 *
 * A path that already carries a *different* extension is rejected rather than
 * having the right one appended: `report.pdf` silently becoming
 * `report.pdf.docx` would write a file the caller never asked for (and hide a
 * confused request behind a valid result).
 */
export function resolveOutputPath(opts: {
  defaultSaveDir: () => string
  ext: string
  title: string
  requestedPath?: string
  overwrite?: boolean
}): string {
  const ext = opts.ext.startsWith('.') ? opts.ext : `.${opts.ext}`
  const fileName = `${sanitizeFileBase(opts.title)}${ext}`
  if (!opts.requestedPath) return uniquePathIn(opts.defaultSaveDir(), fileName)

  if (!isAbsolute(opts.requestedPath)) {
    throw new Error('path must be absolute')
  }
  const requestedExt = extname(opts.requestedPath)
  if (requestedExt && requestedExt.toLowerCase() !== ext) {
    throw new Error(`path must point to a ${ext} file (got "${requestedExt}")`)
  }
  const finalPath = requestedExt ? opts.requestedPath : `${opts.requestedPath}${ext}`
  if (existsSync(finalPath) && !opts.overwrite) {
    throw new Error(`file already exists: ${finalPath} (pass overwrite:true to replace it)`)
  }
  return finalPath
}

export function resolveTargetPath(
  deps: DocToolDeps,
  title: string,
  requestedPath?: string,
  overwrite = false,
): string {
  return resolveOutputPath({
    defaultSaveDir: deps.defaultSaveDir,
    ext: DOCX_EXT,
    title,
    requestedPath,
    overwrite,
  })
}

/** the headless tool: markdown/restricted-HTML -> docx, delegated to the bundled CLI */
function createHeadlessDocxTool(deps: DocToolDeps): McpToolDefinition {
  return {
    name: 'create_docx',
    description:
      'Create a Word .docx file from content and save it to disk without opening the app UI. ' +
      'Content is Markdown (headings, lists, bold/italic, links, code blocks, tables) or a ' +
      'restricted-HTML fragment (format:"html"), converted by the same docx engine the app and ' +
      'the genoffice CLI use. Returns the absolute path of the written file.',
    inputSchema: {
      title: z.string().describe('document title, used as the file name'),
      content: z.string().describe('Markdown source (default) or a restricted-HTML fragment'),
      format: z.enum(['markdown', 'html']).optional().describe('content format; default markdown'),
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
      if (!deps.cli) throw new Error('headless generation is not available in this build')

      const format = args.format === 'html' ? 'html' : 'markdown'
      const targetPath = resolveTargetPath(
        deps,
        title,
        typeof args.path === 'string' ? args.path : undefined,
        args.overwrite === true,
      )
      const { summary, outputPath } = await createFileViaCli(deps.cli, {
        type: 'docx',
        input: {
          name: `content.${format === 'html' ? 'html' : 'md'}`,
          content: String(args.content ?? ''),
        },
        out: targetPath,
        overwrite: args.overwrite === true,
      })
      return { path: outputPath, summary }
    },
  }
}

/**
 * Headless template fill through the bundled CLI's `merge` command: one record
 * of `{{key}}` placeholders into a .docx/.pptx/.xlsx template, written to disk.
 * It writes a file, so it rides the same opt-in as create_docx/create_xlsx.
 */
function createFillTemplateTool(deps: DocToolDeps): McpToolDefinition {
  return {
    name: 'fill_template',
    description:
      'Fill {{key}} placeholders in a .docx, .pptx or .xlsx template with values and write the ' +
      'result to disk without opening the app (a one-record mail merge; loop the tool for a ' +
      'batch). Nested keys flatten to dotted paths (a.b); arrays are not expanded. Returns the ' +
      'fill report: how many placeholders were filled, which had no value, and which data keys ' +
      'went unused.',
    inputSchema: {
      template: z.string().describe('absolute path to the .docx, .pptx or .xlsx template'),
      data: z
        .record(z.string(), z.unknown())
        .describe('the values: an object whose keys are the {{key}} names'),
      path: z
        .string()
        .optional()
        .describe(
          'absolute output path; default is a new file named after the template in the default save folder',
        ),
      overwrite: z
        .boolean()
        .optional()
        .describe('allow replacing an existing file at `path`; default false'),
      strict: z
        .boolean()
        .optional()
        .describe('fail instead of writing when a placeholder has no value; default false'),
    },
    handler: async (args) => {
      const templatePath = String(args.template ?? '')
      if (!isAbsolute(templatePath)) throw new Error('template must be an absolute path')
      const templateExt = extname(templatePath).toLowerCase()
      if (!['.docx', '.pptx', '.xlsx'].includes(templateExt))
        throw new Error('template must point to a .docx, .pptx or .xlsx file')
      if (!existsSync(templatePath)) throw new Error(`file not found: ${templatePath}`)
      if (!deps.cli) throw new Error('headless generation is not available in this build')

      const stem = basename(templatePath, extname(templatePath))
      const targetPath = resolveOutputPath({
        defaultSaveDir: deps.defaultSaveDir,
        ext: templateExt,
        title: stem,
        requestedPath: typeof args.path === 'string' ? args.path : undefined,
        overwrite: args.overwrite === true,
      })
      const json = await mergeTemplateViaCli(deps.cli, {
        templatePath,
        data: args.data as Record<string, unknown>,
        out: targetPath,
        overwrite: args.overwrite === true,
        strict: args.strict === true,
      })
      const detail = json.detail ?? {}
      return { path: json.output_path ?? targetPath, summary: json.summary, ...detail }
    },
  }
}

/** Headless document conversion through the bundled CLI's `convert` command:
 *  one file, one target format, written to disk. Rides the same opt-in as the
 *  the other file-writing tools (create_docx, fill_template). The route table lives in
 *  the CLI — an unsupported pair errors with `detail.supported` naming every
 *  target the source format can reach, so the tool does not duplicate it. */
function createConvertDocumentTool(deps: DocToolDeps): McpToolDefinition {
  return {
    name: 'convert_document',
    description:
      'Convert a document to another format with the GenOffice engines and write the result to ' +
      'disk without opening the app: pdf → docx/pptx/xlsx/md, legacy .doc/.ppt → modern Open ' +
      'XML, odt/ods/ots → xlsx/docx, hwp/hwpx → docx, md → docx/html, docx → md, xlsx → csv. ' +
      'An unsupported pair is refused with the list of targets the source format supports. ' +
      'Returns the absolute path of the written file and which engine ran.',
    inputSchema: {
      path: z.string().describe('absolute path to the source document'),
      to: z
        .string()
        .describe(
          'target format extension, e.g. "docx", "pptx", "xlsx", "md", "html" or "csv" (pdf → docx/pptx/xlsx/md; md → docx/html; docx → md; xlsx → csv)',
        ),
      out: z
        .string()
        .optional()
        .describe(
          'absolute output path; default is a new file named after the source in the default save folder',
        ),
      overwrite: z.boolean().optional().describe('allow replacing an existing file; default false'),
      password: z.string().optional().describe('password when the source is an encrypted PDF'),
      sheet: z
        .string()
        .optional()
        .describe('xlsx → csv: the worksheet to export (default: the active one)'),
    },
    handler: async (args) => {
      const sourcePath = String(args.path ?? '')
      if (!isAbsolute(sourcePath)) throw new Error('path must be absolute')
      const to = String(args.to ?? '')
        .trim()
        .toLowerCase()
      if (!/^[a-z0-9]+$/.test(to)) throw new Error('to must be a format extension like "docx"')
      if (!existsSync(sourcePath)) throw new Error(`file not found: ${sourcePath}`)
      if (!deps.cli) throw new Error('headless generation is not available in this build')

      // default output: the source's stem in the save folder, target extension
      const stem = basename(sourcePath, extname(sourcePath))
      const targetPath = resolveOutputPath({
        defaultSaveDir: deps.defaultSaveDir,
        ext: `.${to}`,
        title: stem,
        requestedPath: typeof args.out === 'string' && args.out.trim() ? args.out : undefined,
        overwrite: args.overwrite === true,
      })
      const json = await convertViaCli(deps.cli, {
        path: sourcePath,
        to,
        out: targetPath,
        overwrite: args.overwrite === true,
        ...(typeof args.password === 'string' && args.password ? { password: args.password } : {}),
        ...(typeof args.sheet === 'string' && args.sheet ? { sheet: args.sheet } : {}),
      })
      const detail = json.detail ?? {}
      return {
        path: (json as { output_path?: string }).output_path ?? targetPath,
        summary: json.summary,
        ...detail,
      }
    },
  }
}

/** `docs check`, read-only: the document-quality report an agent wants before
 *  handing a file over — fields without results, broken references, a stale
 *  TOC, missing images, empty charts or headings, level skips, leftover
 *  placeholder text, pending tracked changes and open comments. No file is
 *  written, so like read_docx it is always registered. */
function createCheckDocumentTool(deps: DocToolDeps): McpToolDefinition {
  return {
    name: 'check_document',
    description:
      'Run the document-quality checks on a .docx without changing it: fields without results, ' +
      'references to missing bookmarks, a TOC that no longer matches the headings, missing ' +
      'images, empty charts, empty headings, list level skips, leftover placeholder text, ' +
      'pending tracked changes and open comments. Each finding carries its severity and ' +
      'location; an issue-free document says so explicitly.',
    inputSchema: {
      path: z.string().describe('absolute path to a .docx file'),
    },
    handler: async (args) => {
      const filePath = String(args.path ?? '')
      if (!isAbsolute(filePath)) throw new Error('path must be absolute')
      if (!existsSync(filePath)) throw new Error(`file not found: ${filePath}`)
      if (!deps.cli) throw new Error('headless generation is not available in this build')
      const json = await docsCheckViaCli(deps.cli, filePath)
      const detail = json.detail ?? {}
      return { summary: json.summary, ...detail }
    },
  }
}

export function createDocumentTools(deps: DocToolDeps, host: SessionHost): McpToolDefinition[] {
  return [
    // headless generation is opt-in: when background is off (the default),
    // clients only see the visible document session
    ...(deps.background === false
      ? []
      : [
          createHeadlessDocxTool(deps),
          createFillTemplateTool(deps),
          createConvertDocumentTool(deps),
        ]),
    createCheckDocumentTool(deps),
    {
      name: 'read_docx',
      description: 'Read a Word .docx file and return its visible text (one paragraph per line).',
      inputSchema: {
        path: z.string().describe('absolute path to a .docx file'),
      },
      handler: async (args) => {
        const filePath = String(args.path ?? '')
        if (!isAbsolute(filePath)) throw new Error('path must be absolute')
        if (extname(filePath).toLowerCase() !== DOCX_EXT)
          throw new Error('path must point to a .docx file')
        if (!existsSync(filePath)) throw new Error(`file not found: ${filePath}`)
        if (!deps.cli) throw new Error('reading .docx is not available in this build')
        const text = await readDocxTextViaCli(deps.cli, filePath)
        return { path: filePath, name: basename(filePath), text }
      },
    },
    {
      name: 'read_docx_properties',
      description:
        "Read a Word .docx file's document properties: core (title, author, dates), app and " +
        'custom properties, as stored in the file.',
      inputSchema: {
        path: z.string().describe('absolute path to a .docx file'),
      },
      handler: async (args) => {
        const filePath = String(args.path ?? '')
        if (!isAbsolute(filePath)) throw new Error('path must be absolute')
        if (extname(filePath).toLowerCase() !== DOCX_EXT)
          throw new Error('path must point to a .docx file')
        if (!existsSync(filePath)) throw new Error(`file not found: ${filePath}`)
        if (!deps.cli) throw new Error('reading .docx is not available in this build')
        const properties = await readDocxPropertiesViaCli(deps.cli, filePath)
        return { path: filePath, name: basename(filePath), properties }
      },
    },
    {
      name: 'open_in_genoffice',
      description: 'Open an existing file in the running GenOffice app, focusing its tab.',
      inputSchema: {
        path: z.string().describe('absolute path to the file to open'),
      },
      handler: async (args) => {
        const filePath = String(args.path ?? '')
        if (!isAbsolute(filePath)) throw new Error('path must be absolute')
        if (!existsSync(filePath)) throw new Error(`file not found: ${filePath}`)
        if (!deps.openInTab) throw new Error('opening files is not available in this build')
        await deps.openInTab(filePath)
        return { ok: true, path: filePath }
      },
    },
    {
      name: 'get_app_info',
      description:
        'Report GenOffice version, the default save folder, the document formats this server can ' +
        "generate, and the editor's full open/save/export format matrix per family.",
      inputSchema: {},
      handler: () => ({
        name: 'GenOffice',
        version: deps.version,
        defaultSaveDir: deps.defaultSaveDir(),
        // Every format listed here is written by a headless create/read tool, and
        // all of them are opt-in: with background generation off there is no
        // generate tool at all, so advertising docx/pptx/xlsx would describe
        // capabilities this client cannot reach.
        formats: deps.background === false ? [] : ['docx', ...(deps.extraFormats ?? [])],
        // editor truth vs. what this server exposes — see tools/formats.ts
        families: capabilityReport({ generating: deps.background !== false }),
      }),
    },
    ...createDocxContentTools(deps, host),
  ]
}

/**
 * The docx session lifecycle as seen by the shared create_session / save_session
 * tools. The content tools below address the tab this driver opened.
 */
export function documentDriver(docs: DocsControl): FamilyDriver {
  return {
    family: 'docx',
    openBlankTab: () => docs.openBlankTab(),
    save: async (wcId, path, overwrite) =>
      docs.runCommand(wcId, 'save_document', { path, overwrite }),
  }
}

/**
 * Visible-editing content tools: edit the document the shared session opened, so
 * the user watches it take shape. These reuse the built-in agent's editor
 * pipeline (the same restricted-HTML parser, ops executor and save path), so the
 * result matches what the in-app AI produces.
 *
 * Only registered when the shell wired a DocsControl — headless/unit runs keep
 * the phase-1 file-only surface.
 */
function createDocxContentTools(deps: DocToolDeps, host: SessionHost): McpToolDefinition[] {
  const docs = deps.docs
  if (!docs) return []

  const target = contentTarget(host, deps.resolveTarget, 'docx')
  const documentField = z
    .string()
    .optional()
    .describe(
      "tab id or path of an open Word document to edit instead of the session's own tab " +
        '(from open_documents {action:"list"}); the edit lands in that live document, and a ' +
        'mutating tool switches the UI to that tab first so the user sees it happen',
    )

  return [
    {
      name: 'insert_content',
      description:
        'Insert content into the visible document as HTML (headings, paragraphs, bold/italic, lists, ' +
        'tables, links). Appends at the end unless afterBlockIndex is given.',
      inputSchema: {
        html: z.string().describe('restricted HTML fragment to insert'),
        afterBlockIndex: z
          .number()
          .int()
          .optional()
          .describe('insert after this block index; default appends at the end'),
        document: documentField,
      },
      handler: async (args) => {
        const wc = await target(args.document, { focus: true })
        return docs.runCommand(wc, 'insert_content', args)
      },
    },
    {
      name: 'replace_blocks',
      description:
        'Replace a range of blocks in the visible document with new HTML content. ' +
        'Use read_document to learn block indexes, then read_blocks to read the full original ' +
        'text before rewriting it.',
      inputSchema: {
        startBlockIndex: z.number().int().describe('first block index to replace (inclusive)'),
        endBlockIndex: z.number().int().describe('last block index to replace (inclusive)'),
        html: z.string().describe('restricted HTML fragment the range is replaced with'),
        document: documentField,
      },
      handler: async (args) => {
        const wc = await target(args.document, { focus: true })
        return docs.runCommand(wc, 'replace_blocks', args)
      },
    },
    {
      name: 'apply_ops',
      description:
        'Apply formatting commands to the visible document (font, paragraph format, heading level, ' +
        'find/replace, list/indent, etc.). `ops` is the same batch format the built-in AI editor accepts; ' +
        'the whole batch is atomic. Use read_document for block indexes.',
      inputSchema: {
        ops: z.array(z.any()).describe('array of op objects'),
        dryRun: z
          .boolean()
          .optional()
          .describe('validate and plan the batch without changing the document'),
        document: documentField,
      },
      handler: async (args) => {
        const wc = await target(args.document, { focus: true })
        return docs.runCommand(wc, 'apply_ops', {
          ops: args.ops,
          dryRun: args.dryRun === true,
        })
      },
    },
    {
      name: 'insert_picture',
      description:
        'Insert a picture into the visible document from a direct image link (png/jpg/gif), ' +
        'at a given size, inline after a block or floating at a position with text wrap. ' +
        'Use read_document for block indexes.',
      inputSchema: {
        url: z.string().describe('direct image link (png/jpg/gif)'),
        afterBlockIndex: z
          .number()
          .int()
          .optional()
          .describe('insert after this block (-1 = start of document); default: end of document'),
        width: z
          .union([z.number(), z.string()])
          .optional()
          .describe(
            'display width, "2.54cm", "1in", "72pt", "96px" or a bare number of points; ' +
              'default natural size capped at 480px',
          ),
        height: z
          .union([z.number(), z.string()])
          .optional()
          .describe('display height; omit to keep the aspect ratio'),
        float: z
          .object({
            anchor: z
              .enum(['paragraph', 'page', 'margin'])
              .optional()
              .describe(
                'what x/y measure from: the anchor paragraph (default), the page corner, ' +
                  'or the margin box corner',
              ),
            x: z
              .union([z.number(), z.string()])
              .optional()
              .describe(
                'horizontal offset, "2.54cm", "1in", "72pt", "96px" or a bare number of points',
              ),
            y: z
              .union([z.number(), z.string()])
              .optional()
              .describe(
                'vertical offset, "2.54cm", "1in", "72pt", "96px" or a bare number of points',
              ),
            wrap: z
              .enum(['square', 'tight', 'topAndBottom', 'behind', 'inFront'])
              .optional()
              .describe('how text wraps around the floating picture (default square)'),
          })
          .optional()
          .describe(
            'omit for an inline picture in its own paragraph; give it to float the picture ' +
              'at a position',
          ),
        altText: z.string().optional().describe('alternative text (accessibility)'),
        document: documentField,
      },
      handler: async (args) => {
        const wc = await target(args.document, { focus: true })
        return docs.runCommand(wc, 'insert_picture', args)
      },
    },
    {
      name: 'read_document',
      description:
        'Read the visible document as a block list: one "index|type|text" line per block, where the ' +
        'index is what apply_ops and insert_content address. Long blocks are clipped for reading; ' +
        'use read_blocks for the full text of a block range.',
      inputSchema: { document: documentField },
      handler: async (args) => {
        const wc = await target(args.document)
        return docs.runCommand(wc, 'read_document', {})
      },
    },
    {
      name: 'read_blocks',
      description:
        'Read the full text of a block range in the visible document as restricted HTML, without ' +
        'the preview clipping of read_document. Use read_document first for block indexes, and read ' +
        'a block with this before rewriting it. Long ranges are paged: a truncated result says ' +
        'which offset to continue from; concatenate the slices in order to get the full HTML.',
      inputSchema: {
        startBlockIndex: z.number().int().describe('start block index (0-based, inclusive)'),
        endBlockIndex: z.number().int().describe('end block index (inclusive)'),
        offset: z
          .number()
          .int()
          .optional()
          .describe(
            'character offset to continue a truncated read; default 0 (use the offset from the previous truncation notice)',
          ),
        document: documentField,
      },
      handler: async (args) => {
        const wc = await target(args.document)
        return docs.runCommand(wc, 'read_blocks', {
          startBlockIndex: args.startBlockIndex,
          endBlockIndex: args.endBlockIndex,
          ...(args.offset === undefined ? {} : { offset: args.offset }),
        })
      },
    },
  ]
}

/** exported for tests */
export const __internal = { resolveTargetPath }
