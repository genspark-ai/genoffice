import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js'
import { defaultRegistry, VERSION } from '../cli'
import type { CommandRegistry } from '../registry'
import { DECK_TOOLS, runDeckTool } from './deck'
import { attachOutputs, defaultOut, materializeInputs, urlErrorOutcome } from './remote'
import {
  createContext,
  disposeContext,
  imageBlocks,
  runJson,
  runWithInline,
  type McpContext,
  type Outcome,
} from './run'
import { buildArgv, resolveTools, toolShape, type ResolvedTool } from './tools'

const ABOUT =
  'GenOffice: create, read, convert, edit and render Office documents locally (docx, xlsx, pptx, pdf, md, html, csv). The app need not be running; render, convert-to-pdf and create_pdf start a hidden GenOffice process for a few seconds.'
const WORKFLOW = [
  'Editing: read the file with the *_read tool, write the ops with the op reference from guide (or the genoffice://guide/* resources), then *_apply. A rejected op names its index and reason; fix that op and resend the whole batch.',
  'A new presentation for a person: deck_start (style sheet + outline), deck_page once per page in order, deck_build, then slides_render to look and slides_audit for geometry, deck_replace to fix a page. Edits to an existing deck: slides_read + slides_apply, keeping its design.',
]

/** What the client shows the model about this server before any tool is called. */
export const INSTRUCTIONS = [
  ABOUT,
  'Paths are absolute, or relative to the working directory the server was started in. Only search, image and media send data off the machine, to the provider configured in GenOffice.',
  ...WORKFLOW,
  'A file open in a GenOffice tab is not written unless force is set. Do not call open unless the user asks to see the file.',
].join('\n')

/** The http-mode variant: the client is on another machine, so files travel as URLs and result content. */
export function remoteInstructions(baseUrl: string): string {
  const base = baseUrl.replace(/\/$/, '')
  return [
    ABOUT,
    `This server runs on another machine: paths you know are not visible to it. To work on a file you have, upload it first (curl -T report.docx ${base}/files/ — the reply carries its url) and pass that url wherever a tool takes a file; any other http(s) URL the server can reach works too. Relative paths and deck folders live in a private scratch directory of this session.`,
    'Omit out: the file a tool writes comes back in the result as output_url (download it with curl -o) and, when small, as an embedded resource with the bytes. Only search, image and media send data to the cloud provider configured in GenOffice.',
    ...WORKFLOW,
  ].join('\n')
}

const GUIDES: { uri: string; name: string; argv: string[]; description: string }[] = [
  {
    uri: 'genoffice://guide/docs',
    name: 'Word ops reference',
    argv: ['guide', 'docs'],
    description: 'every docs_apply op with its fields and the restricted-HTML rules',
  },
  {
    uri: 'genoffice://guide/sheets',
    name: 'Excel ops reference',
    argv: ['guide', 'sheets'],
    description: 'every sheet_apply DSL op with its fields',
  },
  {
    uri: 'genoffice://guide/slides',
    name: 'PowerPoint ops reference',
    argv: ['guide', 'slides'],
    description: 'the slides_apply op groups and vocabulary',
  },
  {
    uri: 'genoffice://guide/slides/design',
    name: 'Deck design guide',
    argv: ['guide', 'slides', 'design'],
    description: 'the staged deck workflow: style sheet, outline, one page at a time, build, QC',
  },
  {
    uri: 'genoffice://guide/slides/spec',
    name: 'Deck spec format',
    argv: ['guide', 'slides', 'spec'],
    description: 'the outline and one-page spec JSON the deck_* tools take',
  },
]

export interface ServerOptions {
  registry?: CommandRegistry
}

export function createMcpServer(ctx: McpContext, opts: ServerOptions = {}): McpServer {
  const registry = opts.registry ?? defaultRegistry()
  const remote = ctx.mode === 'http'
  const server = new McpServer(
    { name: 'genoffice', version: VERSION },
    { instructions: remote ? remoteInstructions(ctx.baseUrl ?? '') : INSTRUCTIONS },
  )

  for (const tool of resolveTools(registry)) {
    // there is no GenOffice window in front of a remote client
    if (remote && tool.name === 'open') continue
    server.registerTool(
      tool.name,
      {
        description: tool.description,
        inputSchema: toolShape(tool, ctx.mode),
        annotations: {
          readOnlyHint: tool.readOnly === true,
          openWorldHint: tool.openWorld === true,
        },
      },
      async (args) => {
        const outcome = await runTool(tool, { ...(args as Record<string, unknown>) }, ctx)
        return toResult(outcome, { images: tool.images, plainText: tool.plainText, ctx })
      },
    )
  }

  for (const tool of DECK_TOOLS) {
    server.registerTool(
      tool.name,
      {
        description: tool.description,
        inputSchema: tool.shape,
        annotations: { readOnlyHint: tool.readOnly === true, openWorldHint: false },
      },
      async (args) =>
        toResult(await runDeckTool(tool, args as Record<string, unknown>, ctx), { ctx }),
    )
  }

  for (const g of GUIDES) {
    server.registerResource(
      g.name,
      g.uri,
      { description: g.description, mimeType: 'text/plain' },
      async (uri) => {
        const r = await runJson(g.argv, ctx)
        const text = r.ok ? r.ok.summary : `${r.error.message}`
        return { contents: [{ uri: uri.href, mimeType: 'text/plain', text }] }
      },
    )
  }
  return server
}

async function runTool(
  tool: ResolvedTool,
  args: Record<string, unknown>,
  ctx: McpContext,
): Promise<Outcome> {
  if (ctx.mode === 'http') {
    try {
      await materializeInputs(tool, args, ctx)
    } catch (err) {
      return urlErrorOutcome(tool, err)
    }
    defaultOut(tool, args, ctx)
  }
  const { argv, inline } = buildArgv(tool, args)
  return runWithInline(argv, inline, ctx)
}

function toResult(
  outcome: Outcome,
  opts: { images?: boolean; plainText?: boolean; ctx: McpContext },
): CallToolResult {
  if (outcome.error) {
    return { isError: true, content: [{ type: 'text', text: JSON.stringify(outcome.error) }] }
  }
  const ok = outcome.ok
  const outputs = opts.ctx.mode === 'http' ? attachOutputs(ok, opts.ctx) : []
  const content: CallToolResult['content'] = [
    { type: 'text', text: opts.plainText ? ok.summary : JSON.stringify(ok) },
    ...outputs,
  ]
  if (opts.images) {
    const { images, omitted } = imageBlocks(ok)
    content.push(...images)
    if (omitted)
      content.push({
        type: 'text',
        text: `${omitted} more page(s) written to disk but not attached; render them one at a time`,
      })
  }
  return { content }
}

export interface ServeOptions extends ServerOptions {
  cwd: string
  env: NodeJS.ProcessEnv
  log: (message: string) => void
}

/** Serves on stdin/stdout until the client closes the transport. */
export async function serveStdio(opts: ServeOptions): Promise<void> {
  const ctx = createContext(opts)
  const server = createMcpServer(ctx, opts)
  const transport = new StdioServerTransport()
  try {
    await new Promise<void>((resolve, reject) => {
      server.server.onclose = () => resolve()
      server.connect(transport).catch(reject)
    })
  } finally {
    disposeContext(ctx)
  }
}
