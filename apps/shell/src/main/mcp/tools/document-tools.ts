import { existsSync } from 'node:fs'
import { basename, extname, isAbsolute, join } from 'node:path'
import { z } from 'zod'
import { createFileViaCli, readDocxTextViaCli } from '../headless-cli'
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
  'insert_content' | 'replace_blocks' | 'apply_ops' | 'read_document' | 'save_document'

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

export function createDocumentTools(deps: DocToolDeps, host: SessionHost): McpToolDefinition[] {
  return [
    // headless generation is opt-in: when background is off (the default),
    // clients only see the visible document session
    ...(deps.background === false ? [] : [createHeadlessDocxTool(deps)]),
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
        'Use read_document to learn block indexes.',
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
      name: 'read_document',
      description:
        'Read the visible document as a block list: one "index|type|text" line per block, where the ' +
        'index is what apply_ops and insert_content address. Long blocks are clipped for reading; ' +
        'the list is an overview, not a lossless copy (use read_docx for full text).',
      inputSchema: { document: documentField },
      handler: async (args) => {
        const wc = await target(args.document)
        return docs.runCommand(wc, 'read_document', {})
      },
    },
  ]
}

/** exported for tests */
export const __internal = { resolveTargetPath }
