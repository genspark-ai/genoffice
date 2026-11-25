import { randomUUID } from 'node:crypto'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js'
import { z } from 'zod'
import { docsToolCatalog, DOCS_READ_TOOLS, docsToolExecutionSchema } from '../shared/tool-catalog'
import { DOCS_DRAFT_CONTROLS, DOCS_RUNTIME_TOOLS } from '../shared/runtime-tools'
import { DOCS_RESOURCE_TOOLS } from '../shared/resource-tools'
import type { McpCommandMessage } from '../shared/ipc'
import { DocsOperationError, isDocsOperationResult } from '../shared/operation-result'
import type { DocsQueryResult } from '../shared/document-query'
import { monitorMcpToolList } from '../shared/mcp-tool-list'
import { formatDocsMcpResult } from '../shared/mcp-result'
import type { AgentToolDef } from '@genoffice/agent-core'

export interface DocsCommandEndpoint {
  request(message: McpCommandMessage, signal: AbortSignal): Promise<unknown>
  closeCaller(callerId: string): void
}

export interface DocsMcpServerOptions {
  version: string
  endpoint: DocsCommandEndpoint
  /** A binding created by the host for the selected document/action, never a tool argument. */
  bindingId?: string
  /** Host-created read grant: writes are neither advertised nor executable. */
  readOnly?: boolean
}

/** A pre-bound document MCP server. The host supplies a standard SDK transport and authorized endpoint. */
export function createDocsMcpServer(options: DocsMcpServerOptions): McpServer {
  const server = new McpServer({ name: 'ForaOffice Docs', version: options.version })
  const callerId = randomUUID()
  const lifetime = new AbortController()
  server.server.onclose = () => {
    lifetime.abort()
    options.endpoint.closeCaller(callerId)
  }
  const tools: AgentToolDef[] = [
    ...docsToolCatalog(),
    ...DOCS_RUNTIME_TOOLS,
    ...DOCS_RESOURCE_TOOLS,
    {
      name: 'read_document',
      description:
        "Read authorized live unsaved content and refresh this caller's context without retargeting a captured selection.",
      inputSchema: { type: 'object', properties: {} },
    },
    {
      name: 'save_document',
      description:
        'Save the bound live Word document, including its document-side state and assets. An unaccepted draft must be settled first.',
      inputSchema: {
        type: 'object',
        properties: { path: { type: 'string' }, overwrite: { type: 'boolean', default: false } },
        required: ['path'],
      },
    },
  ]
  const grantedReadTool = (name: string) =>
    name === 'read_document' || (DOCS_READ_TOOLS.has(name) && !DOCS_DRAFT_CONTROLS.has(name))
  const grantedTools = tools.filter((tool) => !options.readOnly || grantedReadTool(tool.name))
  const discover = async (signal: AbortSignal) => {
    const result = await options.endpoint.request(
      {
        requestId: randomUUID(),
        callerId,
        bindingId: options.bindingId,
        command: 'list_tools',
        payload: {},
      },
      signal,
    )
    signal.throwIfAborted()
    if (!Array.isArray(result)) throw new Error('Invalid document tool list')
    const names = new Set(result)
    return grantedTools.filter((tool) => names.has(tool.name))
  }
  const listed = monitorMcpToolList(
    lifetime.signal,
    async () => (await discover(lifetime.signal)).map((tool) => tool.name),
    () => server.server.sendToolListChanged(),
  )
  for (const tool of grantedTools) {
    server.registerTool(
      tool.name,
      {
        description: tool.description,
        inputSchema: z.fromJSONSchema(docsToolExecutionSchema(tool)),
        annotations: {
          readOnlyHint: DOCS_READ_TOOLS.has(tool.name) || tool.name === 'read_document',
        },
      },
      async (input, extra) => {
        const command =
          tool.name === 'read_document' || tool.name === 'save_document'
            ? tool.name
            : 'execute_tool'
        try {
          const result = await options.endpoint.request(
            {
              requestId: randomUUID(),
              callerId,
              bindingId: options.bindingId,
              command,
              payload: command === 'execute_tool' ? { name: tool.name, input } : input,
            },
            AbortSignal.any([extra.signal, lifetime.signal]),
          )
          if (options.readOnly && isDocsOperationResult(result) && !result.isError) {
            if (tool.name === 'query_document') {
              const query = JSON.parse(result.output) as DocsQueryResult
              for (const item of query.items)
                item.supportedOperations = item.supportedOperations.filter(grantedReadTool)
              return formatDocsMcpResult({ ...result, output: JSON.stringify(query) })
            }
            if (tool.name === 'get_document_context')
              return formatDocsMcpResult({ ...result, editable: false })
          }
          return formatDocsMcpResult(result)
        } catch (error) {
          if (error instanceof DocsOperationError) return formatDocsMcpResult(error.result)
          return {
            content: [
              {
                type: 'text' as const,
                text: error instanceof Error ? error.message : String(error),
              },
            ],
            isError: true,
          }
        }
      },
    )
  }
  // The high-level SDK only serializes ZodObject; composed object schemas
  // (for example required kind OR engineTarget) otherwise become {} on tools/list.
  // Serve the authoritative JSON Schema through the SDK's standard request handler.
  server.server.setRequestHandler(ListToolsRequestSchema, async (_request, extra) => {
    const tools = await discover(AbortSignal.any([extra.signal, lifetime.signal]))
    listed(tools.map((tool) => tool.name))
    return {
      tools: tools.map((tool) => ({
        ...tool,
        inputSchema: { ...tool.inputSchema, type: 'object' as const },
        annotations: {
          readOnlyHint: DOCS_READ_TOOLS.has(tool.name) || tool.name === 'read_document',
        },
      })),
    }
  })
  return server
}
