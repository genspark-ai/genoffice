import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js'
import { defaultRegistry, VERSION } from '../cli'
import type { CommandRegistry } from '../registry'
import { DECK_TOOLS, runDeckTool } from './deck'
import {
  createContext,
  disposeContext,
  imageBlocks,
  runJson,
  runWithInline,
  type McpContext,
  type Outcome,
} from './run'
import { buildArgv, resolveTools, toolShape } from './tools'

/** What the client shows the model about this server before any tool is called. */
export const INSTRUCTIONS = [
  'GenOffice: create, read, convert, edit and render Office documents locally (docx, xlsx, pptx, pdf, md, html, csv). The app need not be running; render, convert-to-pdf and create_pdf start a hidden GenOffice process for a few seconds.',
  'Paths are absolute, or relative to the working directory the server was started in. Only search, image and media send data off the machine, to the provider configured in GenOffice.',
  'Editing: read the file with the *_read tool, write the ops with the op reference from guide (or the genoffice://guide/* resources), then *_apply. A rejected op names its index and reason; fix that op and resend the whole batch.',
  'A new presentation for a person: deck_start (style sheet + outline), deck_page once per page in order, deck_build, then slides_render to look and slides_audit for geometry, deck_replace to fix a page. Edits to an existing deck: slides_read + slides_apply, keeping its design.',
  'A file open in a GenOffice tab is not written unless force is set. Do not call open unless the user asks to see the file.',
].join('\n')

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
  const server = new McpServer(
    { name: 'genoffice', version: VERSION },
    { instructions: INSTRUCTIONS },
  )

  for (const tool of resolveTools(registry)) {
    server.registerTool(
      tool.name,
      {
        description: tool.description,
        inputSchema: toolShape(tool),
        annotations: {
          readOnlyHint: tool.readOnly === true,
          openWorldHint: tool.openWorld === true,
        },
      },
      async (args) => {
        const { argv, inline } = buildArgv(tool, args as Record<string, unknown>)
        const outcome = await runWithInline(argv, inline, ctx)
        return toResult(outcome, { images: tool.images, plainText: tool.plainText })
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
      async (args) => toResult(await runDeckTool(tool, args as Record<string, unknown>, ctx)),
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

function toResult(
  outcome: Outcome,
  opts: { images?: boolean; plainText?: boolean } = {},
): CallToolResult {
  if (outcome.error) {
    return { isError: true, content: [{ type: 'text', text: JSON.stringify(outcome.error) }] }
  }
  const ok = outcome.ok
  const content: CallToolResult['content'] = [
    { type: 'text', text: opts.plainText ? ok.summary : JSON.stringify(ok) },
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
