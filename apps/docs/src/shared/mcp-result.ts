import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js'
import { isDocsOperationResult } from './operation-result'

/** Standard MCP image content keeps binary media out of textual/structured diagnostics. */
export function formatDocsMcpResult(result: unknown): CallToolResult {
  if (isDocsOperationResult(result)) {
    const { media, display, ...metadata } = result
    return {
      content: [
        { type: 'text', text: JSON.stringify(metadata) },
        ...(media ? [{ type: 'image' as const, data: media.data, mimeType: media.mimeType }] : []),
      ],
      structuredContent: metadata,
      ...(display ? { _meta: { 'io.foraoffice/display': display } } : {}),
      ...(result.isError ? { isError: true } : {}),
    }
  }
  return {
    content: [{ type: 'text', text: JSON.stringify(result) }],
    ...(result !== null && typeof result === 'object' && !Array.isArray(result)
      ? { structuredContent: result as Record<string, unknown> }
      : {}),
  }
}
