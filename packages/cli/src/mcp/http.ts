import { randomBytes, randomUUID, timingSafeEqual } from 'node:crypto'
import { createReadStream, createWriteStream, rmSync, statSync } from 'node:fs'
import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http'
import type { Socket } from 'node:net'
import { tmpdir } from 'node:os'
import { delimiter, join } from 'node:path'
import { pipeline } from 'node:stream/promises'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js'
import { FileStore, MAX_TRANSFER_BYTES, mimeOf, safeName } from './files'
import { createContext, disposeContext, type McpContext } from './run'
import { createMcpServer, type ServeOptions } from './server'

/**
 * `genoffice mcp --http <port>`: the same tools over Streamable HTTP for clients
 * on other machines. Routes:
 *   POST /mcp            the MCP session (GET for the notification stream, DELETE to end it)
 *   PUT  /files/<name>   upload a file (also POST /files?name=), reply { url, name, size }
 *   GET  /files/<id>/<name>   download an upload or a tool output
 *   GET  /health
 * Each MCP session gets its own scratch directory, working directory and deck
 * state, so two clients never see each other's files. With GENOFFICE_ALLOWED_ROOTS
 * unset, the tools are confined to the server's own file store.
 */
export interface HttpServeOptions extends ServeOptions {
  port: number
  host?: string
  /** required as `Authorization: Bearer <token>` on every route but /health */
  token?: string
}

export interface HttpHandle {
  url: string
  port: number
  close(): Promise<void>
}

interface Session {
  transport: StreamableHTTPServerTransport
  server: McpServer
  ctx: McpContext
}

const MAX_JSON_BYTES = 32 * 1024 * 1024
const LOOPBACK_HOSTS = new Set(['127.0.0.1', 'localhost', '::1', '[::1]'])

export async function startHttp(opts: HttpServeOptions): Promise<HttpHandle> {
  const host = opts.host ?? '127.0.0.1'
  const files = new FileStore(
    join(tmpdir(), `genoffice-mcp-http-${process.pid}-${randomBytes(4).toString('hex')}`),
  )
  const roots = opts.env.GENOFFICE_ALLOWED_ROOTS?.trim()
  const env = {
    ...opts.env,
    GENOFFICE_ALLOWED_ROOTS: roots ? `${roots}${delimiter}${files.root}` : files.root,
  }
  const sessions = new Map<string, Session>()
  const sockets = new Set<Socket>()
  const log = opts.log

  const json = (res: ServerResponse, status: number, body: unknown): void => {
    res.statusCode = status
    res.setHeader('content-type', 'application/json')
    res.end(JSON.stringify(body))
  }
  const rpcError = (res: ServerResponse, status: number, code: number, message: string): void =>
    json(res, status, { jsonrpc: '2.0', error: { code, message }, id: null })

  const authorized = (req: IncomingMessage): boolean => {
    if (!opts.token) return true
    const given = /^Bearer\s+(.+)$/i.exec(req.headers.authorization ?? '')?.[1]?.trim() ?? ''
    const a = Buffer.from(given)
    const b = Buffer.from(opts.token)
    return a.length === b.length && timingSafeEqual(a, b)
  }

  const baseUrlOf = (req: IncomingMessage): string => {
    const proto = header(req.headers['x-forwarded-proto']) ?? 'http'
    const hostHeader = header(req.headers['x-forwarded-host']) ?? header(req.headers.host)
    return `${proto}://${hostHeader ?? `${host}:${handle.port}`}`
  }

  const openSession = async (req: IncomingMessage, body: unknown): Promise<Session> => {
    const id = randomUUID()
    const ctx = createContext({
      cwd: join(files.root, 'sessions', id),
      scratchDir: join(files.root, 'sessions', id),
      env,
      log,
      mode: 'http',
      files,
      baseUrl: baseUrlOf(req),
    })
    const server = createMcpServer(ctx, opts)
    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: () => id })
    const session: Session = { transport, server, ctx }
    transport.onclose = () => {
      if (sessions.delete(id)) {
        disposeContext(ctx)
        log(`[mcp] session ${id} closed`)
      }
    }
    sessions.set(id, session)
    try {
      await server.connect(transport)
    } catch (err) {
      sessions.delete(id)
      disposeContext(ctx)
      throw err
    }
    log(`[mcp] session ${id} opened for ${JSON.stringify(clientName(body))}`)
    return session
  }

  const handleMcp = async (req: IncomingMessage, res: ServerResponse): Promise<void> => {
    const method = req.method ?? 'GET'
    const sessionId = header(req.headers['mcp-session-id'])
    const existing = sessionId ? sessions.get(sessionId) : undefined
    if (method === 'GET' || method === 'DELETE') {
      if (!existing) {
        rpcError(
          res,
          sessionId ? 404 : 400,
          sessionId ? -32001 : -32000,
          sessionId ? 'Session not found' : 'Mcp-Session-Id header is required',
        )
        return
      }
      await existing.transport.handleRequest(req, res)
      return
    }
    if (method !== 'POST') {
      res.statusCode = 405
      res.setHeader('allow', 'GET, POST, DELETE')
      res.end()
      return
    }
    const body = await readJson(req)
    if (sessionId && !existing) {
      rpcError(res, 404, -32001, `Session not found: ${sessionId}`)
      return
    }
    if (!existing && !isInitialize(body)) {
      rpcError(res, 400, -32000, 'Initialize the session before sending MCP messages')
      return
    }
    const session = existing ?? (await openSession(req, body))
    await session.transport.handleRequest(req, res, body)
  }

  const handleUpload = async (
    req: IncomingMessage,
    res: ServerResponse,
    name: string,
  ): Promise<void> => {
    const declared = Number(req.headers['content-length'] ?? 0)
    if (declared > MAX_TRANSFER_BYTES) {
      json(res, 413, { error: `uploads are limited to ${MAX_TRANSFER_BYTES} bytes` })
      return
    }
    const path = files.uploadTarget(name)
    let received = 0
    req.on('data', (chunk: Buffer) => {
      received += chunk.length
      if (received > MAX_TRANSFER_BYTES) req.destroy(new Error('upload too large'))
    })
    try {
      await pipeline(req, createWriteStream(path))
    } catch (err) {
      rmSync(join(path, '..'), { recursive: true, force: true })
      json(res, 413, { error: err instanceof Error ? err.message : String(err) })
      return
    }
    const stored = files.expose(path)
    const url = files.urlFor(stored, baseUrlOf(req))
    log(`[mcp] upload ${stored.name} (${received} bytes) → ${url}`)
    json(res, 201, { url, name: stored.name, size: received })
  }

  const handleDownload = (req: IncomingMessage, res: ServerResponse, id: string): void => {
    const stored = files.get(id)
    if (!stored) {
      json(res, 404, { error: 'no such file (it may have expired)' })
      return
    }
    const size = statSync(stored.path).size
    res.statusCode = 200
    res.setHeader('content-type', mimeOf(stored.path))
    res.setHeader('content-length', String(size))
    res.setHeader(
      'content-disposition',
      `attachment; filename*=UTF-8''${encodeURIComponent(stored.name)}`,
    )
    if (req.method === 'HEAD') {
      res.end()
      return
    }
    createReadStream(stored.path).pipe(res)
  }

  const route = async (req: IncomingMessage, res: ServerResponse): Promise<void> => {
    const url = new URL(req.url ?? '/', 'http://placeholder')
    if (url.pathname === '/health' && req.method === 'GET') {
      json(res, 200, { status: 'ok', server: 'genoffice', sessions: sessions.size })
      return
    }
    // no token and a loopback bind: refuse Host headers a rebound DNS name would carry
    if (!opts.token && LOOPBACK_HOSTS.has(host)) {
      const hostname = new URL(`http://${header(req.headers.host) ?? ''}`).hostname
      if (!LOOPBACK_HOSTS.has(hostname)) {
        json(res, 403, { error: 'host not allowed' })
        return
      }
    }
    if (!authorized(req)) {
      res.setHeader('www-authenticate', 'Bearer')
      json(res, 401, { error: 'missing or invalid bearer token' })
      return
    }
    if (url.pathname === '/mcp') return handleMcp(req, res)

    const upload = /^\/files\/?([^/]*)$/.exec(url.pathname)
    if (upload && (req.method === 'POST' || req.method === 'PUT')) {
      const name = decodeURIComponent(upload[1] ?? '') || url.searchParams.get('name') || ''
      return handleUpload(
        req,
        res,
        safeName(name || header(req.headers['x-filename']), 'upload.bin'),
      )
    }
    const download = /^\/files\/([a-f0-9]{16})\/[^/]+$/.exec(url.pathname)
    if (download && (req.method === 'GET' || req.method === 'HEAD')) {
      return handleDownload(req, res, download[1]!)
    }
    json(res, 404, { error: 'not found' })
  }

  const httpServer: Server = createServer((req, res) => {
    route(req, res).catch((err) => {
      log(`[mcp] ${req.method} ${req.url}: ${String(err)}`)
      if (!res.headersSent) json(res, 500, { error: 'internal error' })
      else res.end()
    })
  })
  httpServer.on('connection', (socket: Socket) => {
    sockets.add(socket)
    socket.on('close', () => sockets.delete(socket))
  })

  const handle: HttpHandle = {
    url: '',
    port: opts.port,
    async close() {
      for (const [id, s] of [...sessions]) {
        sessions.delete(id)
        await s.transport.close().catch(() => undefined)
        disposeContext(s.ctx)
      }
      const closed = new Promise<void>((resolve) => httpServer.close(() => resolve()))
      for (const socket of sockets) socket.destroy()
      await closed
      files.dispose()
    },
  }

  await new Promise<void>((resolve, reject) => {
    httpServer.once('error', reject)
    httpServer.listen(opts.port, host, () => {
      httpServer.off('error', reject)
      const address = httpServer.address()
      handle.port = typeof address === 'object' && address ? address.port : opts.port
      const shown = host.includes(':') ? `[${host}]` : host
      handle.url = `http://${shown}:${handle.port}`
      resolve()
    })
  })
  if (!opts.token && !LOOPBACK_HOSTS.has(host)) {
    log(
      '[mcp] warning: reachable from other machines without --token; anyone on the network can use the tools',
    )
  }
  log(`[mcp] listening on ${handle.url}/mcp (uploads: PUT ${handle.url}/files/<name>)`)
  return handle
}

/** Serves until SIGINT or SIGTERM. */
export async function serveHttp(opts: HttpServeOptions): Promise<void> {
  const handle = await startHttp(opts)
  await new Promise<void>((resolve) => {
    const stop = (): void => {
      process.off('SIGINT', stop)
      process.off('SIGTERM', stop)
      void handle.close().finally(resolve)
    }
    process.once('SIGINT', stop)
    process.once('SIGTERM', stop)
  })
}

function header(value: string | string[] | undefined): string | undefined {
  return Array.isArray(value) ? value[0] : value
}

function isInitialize(body: unknown): boolean {
  const messages = Array.isArray(body) ? body : [body]
  return messages.some((m) => (m as { method?: string })?.method === 'initialize')
}

function clientName(body: unknown): unknown {
  const messages = Array.isArray(body) ? body : [body]
  const init = messages.find((m) => (m as { method?: string })?.method === 'initialize') as
    { params?: { clientInfo?: unknown } } | undefined
  return init?.params?.clientInfo ?? 'unknown client'
}

function readJson(req: IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    let size = 0
    req.on('data', (chunk: Buffer) => {
      size += chunk.length
      if (size > MAX_JSON_BYTES) {
        reject(new Error('request body too large'))
        req.destroy()
        return
      }
      chunks.push(chunk)
    })
    req.on('end', () => {
      const raw = Buffer.concat(chunks).toString('utf8')
      if (!raw) return resolve(undefined)
      try {
        resolve(JSON.parse(raw))
      } catch {
        reject(new Error('invalid JSON body'))
      }
    })
    req.on('error', reject)
  })
}
