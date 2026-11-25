// Modified for ForaOffice from GenOffice v0.12.852.
import { existsSync } from 'node:fs'
import { basename, extname, isAbsolute, join } from 'node:path'
import { z } from 'zod'
import {
  createFileViaCli,
  mergeTemplateViaCli,
  readDocxTextViaCli,
  readDocxPropertiesViaCli,
} from '../headless-cli'
import type { McpInvocation, McpToolDefinition } from '../mcp-server'
import {
  docsToolCatalog,
  DOCS_READ_TOOLS,
  docsToolExecutionSchema,
} from '../../../../../docs/src/shared/tool-catalog'
import { DOCS_RUNTIME_TOOLS } from '../../../../../docs/src/shared/runtime-tools'
import { formatDocsMcpResult } from '../../../../../docs/src/shared/mcp-result'
import { DOCS_RESOURCE_TOOLS } from '../../../../../docs/src/shared/resource-tools'
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
  listDocumentTargets?: () => Promise<number[]>
}

/** editor commands the docs renderer bridge understands (see docs shared/ipc.ts) */
export type McpEditorCommandName =
  | 'insert_content'
  | 'replace_blocks'
  | 'apply_ops'
  | 'insert_picture'
  | 'read_document'
  | 'save_document'
  | 'execute_tool'
  | 'list_tools'

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
  runCommand: (
    wcId: number,
    command: McpEditorCommandName,
    payload: unknown,
    invocation?: McpInvocation,
  ) => Promise<unknown>
  closeCaller?: (callerId: string) => void
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

export function createDocumentTools(deps: DocToolDeps, host: SessionHost): McpToolDefinition[] {
  return [
    // headless generation is opt-in: when background is off (the default),
    // clients only see the visible document session
    ...(deps.background === false
      ? []
      : [createHeadlessDocxTool(deps), createFillTemplateTool(deps)]),
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
    save: async (wcId, path, overwrite, invocation) =>
      docs.runCommand(wcId, 'save_document', { path, overwrite }, invocation),
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
  const legacyCommands = new Set([
    'insert_content',
    'replace_blocks',
    'apply_ops',
    'insert_picture',
  ])
  const discover = async (invocation: McpInvocation) => {
    const targets = new Set(await deps.listDocumentTargets?.())
    const active = host.current()
    if (active?.family === 'docx') targets.add(active.wcId)
    const lists = await Promise.allSettled(
      [...targets].map((wcId) => docs.runCommand(wcId, 'list_tools', {}, invocation)),
    )
    invocation.signal.throwIfAborted()
    return new Set<string>(
      lists.flatMap((list) =>
        list.status === 'fulfilled' && Array.isArray(list.value) ? list.value : [],
      ),
    )
  }
  const withDocument = (schema: Record<string, unknown>) => ({
    ...schema,
    type: 'object' as const,
    properties: {
      ...(schema.properties as Record<string, unknown>),
      document: {
        type: 'string' as const,
        description: 'An open document id or path; omit for the active session document.',
      },
    },
  })
  return [
    ...[...docsToolCatalog(), ...DOCS_RUNTIME_TOOLS, ...DOCS_RESOURCE_TOOLS].map(
      (tool): McpToolDefinition => ({
        name: tool.name,
        description: tool.description,
        discover,
        formatResult: formatDocsMcpResult,
        jsonSchema: withDocument(tool.inputSchema),
        inputSchema: z.fromJSONSchema(withDocument(docsToolExecutionSchema(tool))),
        handler: async (args, invocation) => {
          const { document, ...input } = args
          const wc = await target(document, { focus: !DOCS_READ_TOOLS.has(tool.name) })
          return legacyCommands.has(tool.name)
            ? docs.runCommand(wc, tool.name as McpEditorCommandName, input, invocation)
            : docs.runCommand(wc, 'execute_tool', { name: tool.name, input }, invocation)
        },
      }),
    ),
    {
      name: 'read_document',
      discover,
      description:
        "Read the current live document context and refresh this caller's editing baseline.",
      inputSchema: {
        document: z
          .string()
          .optional()
          .describe('An open document id or path; omit for the active session document.'),
      },
      handler: async (args, invocation) =>
        docs.runCommand(await target(args.document), 'read_document', {}, invocation),
    },
  ]
}

/** exported for tests */
export const __internal = { resolveTargetPath }
