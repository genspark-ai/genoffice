import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, extname, join } from 'node:path'
import type { IncomingMessage, ServerResponse } from 'node:http'
import type { CliRunner } from '../main/mcp/cli-runner'
import { tokenValid } from './static'

/**
 * Headless conversion API (issue #2080's third delivery form, lite): the
 * existing `genoffice` CLI wrapped in HTTP. A client uploads document bytes,
 * the server runs the same CLI an agent would (`convert` / `info`, `--json`)
 * on a private temp copy, and the response is either the converted file or
 * the CLI's JSON envelope verbatim. The CLI keeps owning format routing and
 * validation; this layer only moves bytes.
 */

const MAX_UPLOAD_BYTES = 200 * 1024 * 1024
const CONVERT_TIMEOUT_MS = 180_000

/** the CLI's exit-code convention (packages/cli/src/result.ts) mapped to HTTP */
const EXIT_STATUS: Record<number, number> = {
  0: 200,
  1: 400, // usage / rejected ops
  2: 409, // file not found / output exists / open in GUI
  3: 422, // conversion failed
  4: 503, // app not available
}

export interface ApiResult {
  status: number
  contentType: string
  /** binary payload (converted file) or text (JSON envelope / error) */
  body: Buffer | string
  disposition?: string
}

export interface WebApiDeps {
  runner: CliRunner
  workDir?: () => Promise<string>
}

function sanitizeName(name: string | null): string | null {
  if (!name) return null
  const base = basename(name).trim()
  if (!base || base.length > 255 || base.includes('\0') || /^[\s.]+$/.test(base)) return null
  return base
}

/** run one CLI conversion over an uploaded document */
async function convert(
  deps: WebApiDeps,
  name: string,
  to: string | null,
  bytes: Buffer,
): Promise<ApiResult> {
  if (!to || !/^[\w-]+$/.test(to)) {
    return {
      status: 400,
      contentType: 'application/json',
      body: '{"error":"missing or invalid ?to=<fmt>"}',
    }
  }
  const dir = await (deps.workDir ? deps.workDir() : mkdtemp(join(tmpdir(), 'genoffice-api-')))
  try {
    const input = join(dir, name)
    await writeFile(input, bytes)
    const output = join(dir, `${basename(name, extname(name))}.${to}`)
    const outcome = await deps.runner.run(
      ['convert', input, '--to', to, '--out', output, '--force', '--json'],
      { timeoutMs: CONVERT_TIMEOUT_MS, cwd: dir },
    )
    if (!outcome.ok || outcome.json?.status !== 'ok') {
      return {
        status: EXIT_STATUS[outcome.code] ?? 500,
        contentType: 'application/json',
        body: outcome.json
          ? JSON.stringify(outcome.json)
          : JSON.stringify({
              status: 'error',
              code: outcome.code,
              message: outcome.stderr.slice(0, 2000) || 'conversion failed',
            }),
      }
    }
    const converted = await readFile(output)
    return {
      status: 200,
      contentType: 'application/octet-stream',
      body: converted,
      disposition: `attachment; filename="${basename(output)}"`,
    }
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}

/** run one CLI info/read-style command over an uploaded document */
async function info(deps: WebApiDeps, name: string, bytes: Buffer): Promise<ApiResult> {
  const dir = await (deps.workDir ? deps.workDir() : mkdtemp(join(tmpdir(), 'genoffice-api-')))
  try {
    const input = join(dir, name)
    await writeFile(input, bytes)
    const outcome = await deps.runner.run(['info', input, '--json'], { cwd: dir })
    const status = outcome.ok ? 200 : (EXIT_STATUS[outcome.code] ?? 500)
    return {
      status,
      contentType: 'application/json',
      body:
        outcome.stdout ||
        JSON.stringify({
          status: 'error',
          code: outcome.code,
          message: outcome.stderr.slice(0, 2000),
        }),
    }
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}

/** the request core, separated from node:http so tests drive it directly */
export async function handleApiRequest(
  deps: WebApiDeps,
  method: string,
  url: string,
  body: Buffer,
): Promise<ApiResult> {
  const parsed = new URL(url, 'http://localhost')
  const name = sanitizeName(parsed.searchParams.get('name'))
  if (method !== 'POST') {
    return { status: 405, contentType: 'application/json', body: '{"error":"POST only"}' }
  }
  if (!name || extname(name) === '') {
    return {
      status: 400,
      contentType: 'application/json',
      body: '{"error":"?name=<file.ext> with an extension is required"}',
    }
  }
  if (body.length === 0) {
    return {
      status: 400,
      contentType: 'application/json',
      body: '{"error":"request body must carry the document bytes"}',
    }
  }
  if (body.length > MAX_UPLOAD_BYTES) {
    return {
      status: 413,
      contentType: 'application/json',
      body: '{"error":"upload exceeds the 200MB cap"}',
    }
  }
  switch (parsed.pathname) {
    case '/api/convert':
      return convert(deps, name, parsed.searchParams.get('to'), body)
    case '/api/info':
      return info(deps, name, body)
    default:
      return {
        status: 404,
        contentType: 'application/json',
        body: '{"error":"unknown endpoint; /api/convert and /api/info exist"}',
      }
  }
}

function send(res: ServerResponse, result: ApiResult): void {
  const headers: Record<string, string> = { 'Content-Type': result.contentType }
  if (result.disposition) headers['Content-Disposition'] = result.disposition
  res.writeHead(result.status, headers)
  res.end(result.body)
}

/** true when the path belongs to the API and was served; false → static/websocket */
export async function serveWebApi(
  req: IncomingMessage,
  res: ServerResponse,
  options: { token: string; runner: CliRunner; port: number },
): Promise<boolean> {
  const url = req.url ?? '/'
  if (!url.startsWith('/api/')) return false
  if (!tokenValid(req, options.token, options.port)) {
    res.writeHead(401, { 'Content-Type': 'application/json' })
    res.end('{"error":"token required"}')
    return true
  }
  if (req.method !== 'POST') {
    send(res, { status: 405, contentType: 'application/json', body: '{"error":"POST only"}' })
    return true
  }
  const chunks: Buffer[] = []
  let size = 0
  let overflow = false
  for await (const chunk of req) {
    size += (chunk as Buffer).length
    if (size > MAX_UPLOAD_BYTES) {
      overflow = true
      break
    }
    chunks.push(chunk as Buffer)
  }
  if (overflow) {
    send(res, {
      status: 413,
      contentType: 'application/json',
      body: '{"error":"upload exceeds the 200MB cap"}',
    })
    return true
  }
  const result = await handleApiRequest(
    { runner: options.runner },
    req.method ?? 'POST',
    url,
    Buffer.concat(chunks),
  )
  send(res, result)
  return true
}
