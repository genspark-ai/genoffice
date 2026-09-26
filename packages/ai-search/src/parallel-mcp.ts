import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import type { Transport } from '@modelcontextprotocol/sdk/shared/transport.js'

/** Anonymous Search MCP: the keyless Parallel path, also the default fallback before DuckDuckGo. */
export async function parallelMcpSearch(query: string): Promise<unknown> {
  const controller = new AbortController()
  // Bound the entire handshake + tool call, including streamed response bodies.
  const timer = setTimeout(() => controller.abort(), 15000)
  const client = new Client({ name: 'genoffice', version: '0.1.0' })
  const transport = new StreamableHTTPClientTransport(new URL('https://search.parallel.ai/mcp'), {
    fetch: (input, init) =>
      fetch(input, {
        ...init,
        signal: AbortSignal.any(
          init?.signal ? [controller.signal, init.signal] : [controller.signal],
        ),
      }),
  })
  const options = { signal: controller.signal, timeout: 15000 }
  try {
    // SDK's sessionId getter includes undefined; its Transport interface uses an
    // optional string instead, which conflicts with sheets' exact optional types.
    await client.connect(transport as Transport, options)
    const result = await client.callTool(
      { name: 'web_search', arguments: { objective: query, search_queries: [query] } },
      undefined,
      options,
    )
    if (result.isError) return null
    if (result.structuredContent) return result.structuredContent
    // MCP also permits the structured result serialized into a text block.
    if (Array.isArray(result.content)) {
      for (const block of result.content) {
        if (block.type !== 'text' || typeof block.text !== 'string') continue
        try {
          return JSON.parse(block.text) as unknown
        } catch {
          // Non-JSON explanatory text is not a search result.
        }
      }
    }
    return null
  } finally {
    try {
      if (transport.sessionId && !controller.signal.aborted) await transport.terminateSession()
    } catch {
      // Cleanup failure must not discard usable search results.
    } finally {
      controller.abort()
      clearTimeout(timer)
      await client.close().catch(() => {})
    }
  }
}
