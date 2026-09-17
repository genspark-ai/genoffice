import { randomBytes } from 'node:crypto'
import { existsSync, readFileSync, statSync } from 'node:fs'
import { basename, extname, join } from 'node:path'
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js'
import type { JsonOk } from '../result'
import { fetchToFile, isHttpUrl, mimeOf } from './files'
import type { McpContext, Outcome } from './run'
import { PATH_KEYS, type ResolvedTool } from './tools'

/**
 * What changes when the client is on another machine (http mode): URL inputs
 * are fetched onto this side first, a missing `out` lands in the session
 * scratch directory, and every file a tool wrote is handed back as a download
 * URL plus, when small, the bytes themselves.
 */

const INLINE_MAX_BYTES = 2 * 1024 * 1024

export async function materializeInputs(
  tool: ResolvedTool,
  args: Record<string, unknown>,
  ctx: McpContext,
): Promise<void> {
  const keys = [
    ...(tool.positionals ?? []).filter((p) => PATH_KEYS.has(p.key)).map((p) => p.key),
    ...tool.params.filter((p) => p.kind === 'string' && PATH_KEYS.has(p.option)).map((p) => p.key),
  ]
  for (const key of keys) {
    const value = args[key]
    if (!isHttpUrl(value)) continue
    args[key] = ctx.files?.resolveOwnUrl(value) ?? (await fetchToFile(value, ctx.scratchDir))
  }
}

/** A remote client has no useful path to give; outputs go to the scratch dir and come back in the result. */
export function defaultOut(
  tool: ResolvedTool,
  args: Record<string, unknown>,
  ctx: McpContext,
): void {
  if (!tool.params.some((p) => p.key === 'out')) return
  if (args.out !== undefined && args.out !== null && args.out !== '') return
  const id = randomBytes(4).toString('hex')
  if (tool.command === 'render' || tool.verb === 'render') {
    args.out = join(ctx.scratchDir, `render-${id}`)
    return
  }
  if (tool.command !== 'create') return
  const type = tool.fixed?.[tool.fixed.indexOf('--type') + 1] ?? 'docx'
  const from = typeof args.from === 'string' ? basename(args.from, extname(args.from)) : ''
  args.out = join(ctx.scratchDir, `out-${id}`, `${from || 'document'}.${type}`)
}

export function urlErrorOutcome(tool: ResolvedTool, err: unknown): Outcome {
  return {
    error: {
      status: 'error',
      command: tool.command,
      code: 2,
      error: 'file_not_found',
      message: err instanceof Error ? err.message : String(err),
      suggestion: 'upload the file with POST /files and pass the URL it returns',
    },
  }
}

/**
 * Adds `output_url` (and a `url` per detail.files entry) to the envelope and
 * returns the content blocks that carry the output file: the bytes as an
 * embedded resource when small, a resource_link otherwise.
 */
export function attachOutputs(ok: JsonOk, ctx: McpContext): CallToolResult['content'] {
  const files = ctx.files
  const base = ctx.baseUrl
  if (!files || !base) return []
  const content: CallToolResult['content'] = []
  const listed = (ok.detail?.files as { path?: string; url?: string }[] | undefined) ?? []
  for (const f of listed) {
    if (f.path && isFile(f.path)) f.url = files.urlFor(files.expose(f.path), base)
  }
  const out = ok.output_path
  if (!out || !isFile(out)) return content
  const stored = files.expose(out)
  const url = files.urlFor(stored, base)
  ;(ok as JsonOk & { output_url: string }).output_url = url
  const size = statSync(out).size
  const mimeType = mimeOf(out)
  if (size <= INLINE_MAX_BYTES) {
    content.push({
      type: 'resource',
      resource: { uri: url, mimeType, blob: readFileSync(out).toString('base64') },
    })
  } else {
    content.push({ type: 'resource_link', uri: url, name: stored.name, mimeType })
  }
  return content
}

function isFile(path: string): boolean {
  return existsSync(path) && statSync(path).isFile()
}
