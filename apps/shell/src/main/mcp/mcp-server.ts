import { randomUUID } from 'node:crypto'
import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http'
import type { Socket } from 'node:net'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { SSEServerTransport } from '@modelcontextprotocol/sdk/server/sse.js'
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js'
import type { ZodRawShape } from 'zod'

/**
 * Local MCP server for GenOffice.
 *
 * Lives in the Electron main process, so it only exists while the app runs.
 * Speaks both the current Streamable HTTP transport (`/mcp`) and the legacy SSE
 * transport (`/sse` + `/messages`) for older clients, and binds loopback only.
 *
 * Implemented on plain `node:http` (the same shape as the sheets capture
 * server) rather than express: the SDK transports accept Node request/response
 * objects directly, so express would add a dependency for nothing.
 */

export interface McpToolDefinition {
  name: string
  description: string
  /** zod raw shape, as the SDK expects (not a z.object()) */
  inputSchema?: ZodRawShape
  /** return value is serialized as JSON text; a thrown error becomes isError */
  handler: (args: Record<string, unknown>) => Promise<unknown> | unknown
}

export interface McpServerOptions {
  port: number
  /** fixed tool set shared by every client session (omit when using toolsFactory) */
  tools?: McpToolDefinition[]
  /**
   * Build a fresh tool set per client session instead of sharing `tools`.
   * Tools that hold session state (the visible-editing session host) pass this
   * so two connected clients never fight over one active session. Read the
   * current settings inside the factory: it runs at connect time.
   */
  toolsFactory?: () => McpToolDefinition[]
  logger?: (message: string) => void
  /** reported to clients in the initialize handshake; defaults to the app version when wired */
  version?: string
  /** EADDRINUSE retry tuning (exposed so tests do not wait out the backoff) */
  maxListenAttempts?: number
  listenRetryDelayMs?: number
}

export const DEFAULT_MCP_PORT = 3093
const LISTEN_RETRY_DELAY_MS = 1500
const MAX_BODY_BYTES = 32 * 1024 * 1024

export class McpServerService {
  private readonly tools: McpToolDefinition[]
  private readonly toolsFactory?: () => McpToolDefinition[]
  private readonly logger: (message: string) => void
  private readonly version: string
  private readonly maxListenAttempts: number
  private readonly listenRetryDelayMs: number
  private httpServer?: Server
  private port: number
  private running = false

  /** per-session servers/transports, keyed by session id */
  private readonly sessionServers = new Map<string, McpServer>()
  private readonly streamableTransports = new Map<string, StreamableHTTPServerTransport>()
  private readonly legacyTransports = new Map<string, SSEServerTransport>()
  private readonly sockets = new Set<Socket>()

  /** bumps on stop so an in-flight start cannot bind after cancellation */
  private lifecycleGeneration = 0

  constructor(options: McpServerOptions) {
    this.tools = options.tools ?? []
    this.toolsFactory = options.toolsFactory
    this.port = options.port
    this.logger = options.logger ?? (() => undefined)
    this.version = options.version ?? '0.1.0'
    this.maxListenAttempts = options.maxListenAttempts ?? 5
    this.listenRetryDelayMs = options.listenRetryDelayMs ?? LISTEN_RETRY_DELAY_MS
  }

  isRunning(): boolean {
    return this.running
  }

  getPort(): number {
    return this.port
  }

  /** URL a client should connect to */
  getUrl(): string {
    return `http://127.0.0.1:${this.port}`
  }

  // ── tool registration ─────────────────────────────────────────────────────

  private createSessionServer(): McpServer {
    const server = new McpServer({ name: 'GenOffice', version: this.version })
    // a factory gives each connected client its own tool instances (session
    // state lives in their closures); otherwise the fixed set is shared
    const tools = this.toolsFactory ? this.toolsFactory() : this.tools
    for (const tool of tools) {
      server.registerTool(
        tool.name,
        {
          description: tool.description,
          ...(tool.inputSchema ? { inputSchema: tool.inputSchema } : {}),
        },
        (async (args: Record<string, unknown>) => {
          const started = Date.now()
          try {
            const result = await tool.handler(args ?? {})
            this.logger(`[mcp] tool ${tool.name} ok (${Date.now() - started}ms)`)
            return { content: [{ type: 'text' as const, text: stringifyResult(result) }] }
          } catch (error) {
            const message = error instanceof Error ? error.message : String(error)
            this.logger(`[mcp] tool ${tool.name} error (${Date.now() - started}ms): ${message}`)
            return { content: [{ type: 'text' as const, text: message }], isError: true }
          }
        }) as never,
      )
    }
    return server
  }

  // ── request handling ──────────────────────────────────────────────────────

  private handleRequest = (req: IncomingMessage, res: ServerResponse): void => {
    void this.route(req, res).catch((error) => {
      this.logger(`[mcp] unhandled request error: ${String(error)}`)
      if (!res.headersSent) {
        res.statusCode = 500
        res.setHeader('content-type', 'application/json')
        res.end(JSON.stringify({ error: 'internal error' }))
      }
    })
  }

  private async route(req: IncomingMessage, res: ServerResponse): Promise<void> {
    const url = new URL(req.url ?? '/', `http://127.0.0.1:${this.port}`)
    // DNS-rebinding defense: the server has no auth, so only loopback host
    // headers are accepted. Requests without an Origin (curl, local CLI) pass.
    if (!this.checkHost(req, res) || !this.checkOrigin(req, res)) return

    if (url.pathname === '/health' && req.method === 'GET') {
      return this.json(res, 200, {
        status: 'ok',
        server: 'GenOffice',
        transport: 'StreamableHTTP + SSE',
        port: this.port,
      })
    }

    if (url.pathname === '/mcp') return this.handleStreamable(req, res)

    if (url.pathname === '/sse' && req.method === 'GET') return this.handleLegacySse(req, res)

    if (url.pathname === '/messages' && req.method === 'POST') {
      return this.handleLegacyMessage(req, res, url)
    }

    this.json(res, 404, { error: 'not found' })
  }

  private async handleStreamable(req: IncomingMessage, res: ServerResponse): Promise<void> {
    const method = req.method ?? 'GET'
    if (method !== 'POST' && method !== 'GET' && method !== 'DELETE') {
      res.statusCode = 405
      res.setHeader('allow', 'GET, POST, DELETE')
      res.end()
      return
    }

    const sessionId = headerValue(req.headers['mcp-session-id'])

    // GET/DELETE must belong to a live session
    if (method === 'GET' || method === 'DELETE') {
      const transport = sessionId ? this.streamableTransports.get(sessionId) : undefined
      if (!transport) {
        this.json(res, sessionId ? 404 : 400, {
          jsonrpc: '2.0',
          error: {
            code: sessionId ? -32001 : -32000,
            message: sessionId ? 'Session not found' : 'Mcp-Session-Id header is required',
          },
          id: null,
        })
        return
      }
      await transport.handleRequest(req, res)
      return
    }

    // POST
    const body = await readJsonBody(req)
    const existing = sessionId ? this.streamableTransports.get(sessionId) : undefined
    if (sessionId && !existing) {
      this.json(res, 404, {
        jsonrpc: '2.0',
        error: { code: -32001, message: `Session not found: ${sessionId}` },
        id: null,
      })
      return
    }
    if (!existing && !isInitializeRequest(body)) {
      this.json(res, 400, {
        jsonrpc: '2.0',
        error: { code: -32000, message: 'Initialize the session before sending MCP messages' },
        id: null,
      })
      return
    }

    let transport = existing
    if (!transport) {
      const newSessionId = sessionId ?? randomUUID()
      const created: StreamableHTTPServerTransport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => newSessionId,
        onsessioninitialized: (sid) => {
          this.logger(`[mcp] session initialized: ${sid}`)
        },
      })
      created.onclose = () => {
        this.streamableTransports.delete(newSessionId)
        this.sessionServers.delete(newSessionId)
      }
      this.streamableTransports.set(newSessionId, created)
      const server = this.createSessionServer()
      this.sessionServers.set(newSessionId, server)
      try {
        await server.connect(created)
      } catch (error) {
        this.streamableTransports.delete(newSessionId)
        this.sessionServers.delete(newSessionId)
        throw error
      }
      transport = created
    }

    await transport.handleRequest(req, res, body)
  }

  private async handleLegacySse(req: IncomingMessage, res: ServerResponse): Promise<void> {
    res.setHeader('content-type', 'text/event-stream')
    res.setHeader('cache-control', 'no-cache')
    res.setHeader('connection', 'keep-alive')
    res.setHeader('x-accel-buffering', 'no')

    const transport = new SSEServerTransport('/messages', res)
    const sessionId = transport.sessionId
    this.legacyTransports.set(sessionId, transport)

    const heartbeat = setInterval(() => {
      if (!res.writableEnded) {
        try {
          res.write(': heartbeat\n\n')
        } catch {
          clearInterval(heartbeat)
        }
      }
    }, 15000)

    const cleanup = (): void => {
      clearInterval(heartbeat)
      this.legacyTransports.delete(sessionId)
      this.sessionServers.delete(sessionId)
    }
    res.on('close', cleanup)
    res.on('error', cleanup)

    const server = this.createSessionServer()
    this.sessionServers.set(sessionId, server)
    try {
      await server.connect(transport)
    } catch (error) {
      cleanup()
      this.logger(`[mcp] legacy SSE connect failed: ${String(error)}`)
      if (!res.headersSent) res.statusCode = 500
      res.end()
    }
  }

  private async handleLegacyMessage(
    req: IncomingMessage,
    res: ServerResponse,
    url: URL,
  ): Promise<void> {
    const sessionId = url.searchParams.get('sessionId')
    if (!sessionId) {
      this.json(res, 400, { error: 'Missing sessionId parameter' })
      return
    }
    const transport = this.legacyTransports.get(sessionId)
    if (!transport) {
      this.json(res, 400, { error: `No transport found for sessionId ${sessionId}` })
      return
    }
    const body = await readJsonBody(req)
    // express.json() would have consumed the stream; passing the parsed body
    // keeps the SDK from re-reading an exhausted request
    await transport.handlePostMessage(req, res, body)
  }

  // ── guards ────────────────────────────────────────────────────────────────

  private checkHost(req: IncomingMessage, res: ServerResponse): boolean {
    const host = req.headers.host
    try {
      const parsed = new URL(`http://${host ?? ''}`)
      const validHost = parsed.hostname === '127.0.0.1' || parsed.hostname === 'localhost'
      const validPort = parsed.port === String(this.port)
      if (validHost && validPort) return true
    } catch {
      // fall through to rejection
    }
    this.logger(`[mcp] rejected invalid Host header: ${host ?? '(missing)'}`)
    this.json(res, 403, { error: 'Invalid host' })
    return false
  }

  private checkOrigin(req: IncomingMessage, res: ServerResponse): boolean {
    const origin = req.headers.origin
    if (!origin) return true
    try {
      const parsed = new URL(origin)
      const ok =
        (parsed.protocol === 'http:' || parsed.protocol === 'https:') &&
        (parsed.hostname === '127.0.0.1' || parsed.hostname === 'localhost') &&
        parsed.port === String(this.port)
      if (ok) return true
    } catch {
      // fall through to rejection
    }
    this.logger(`[mcp] rejected invalid Origin header: ${origin}`)
    this.json(res, 403, { error: 'Invalid origin' })
    return false
  }

  private json(res: ServerResponse, status: number, body: unknown): void {
    res.statusCode = status
    res.setHeader('content-type', 'application/json')
    res.end(JSON.stringify(body))
  }

  // ── lifecycle ─────────────────────────────────────────────────────────────

  async start(port = this.port): Promise<void> {
    if (this.running) return
    this.port = port
    const generation = ++this.lifecycleGeneration

    let lastError: unknown
    for (let attempt = 1; attempt <= this.maxListenAttempts; attempt++) {
      try {
        await this.listenOnce(generation)
        this.running = true
        this.logger(`[mcp] listening on ${this.getUrl()} (streamable /mcp, legacy /sse)`)
        return
      } catch (error) {
        lastError = error
        if (generation !== this.lifecycleGeneration) {
          throw new Error('MCP server start cancelled', { cause: error })
        }
        if ((error as NodeJS.ErrnoException)?.code !== 'EADDRINUSE') throw error
        this.logger(`[mcp] port ${this.port} in use, retry ${attempt}/${this.maxListenAttempts}`)
        if (attempt < this.maxListenAttempts) {
          await new Promise((resolve) => setTimeout(resolve, this.listenRetryDelayMs))
        }
      }
    }
    throw lastError
  }

  private listenOnce(generation: number): Promise<void> {
    return new Promise((resolve, reject) => {
      const server = createServer(this.handleRequest)
      let settled = false

      server.on('connection', (socket: Socket) => {
        this.sockets.add(socket)
        socket.on('close', () => this.sockets.delete(socket))
      })

      server.on('error', (error: NodeJS.ErrnoException) => {
        if (settled) {
          this.logger(`[mcp] server error after listen: ${String(error)}`)
          return
        }
        settled = true
        try {
          server.close()
        } catch {
          // nothing to clean up
        }
        reject(error)
      })

      // loopback only: the server is unauthenticated, so it must never be
      // reachable from another machine
      server.listen(this.port, '127.0.0.1', () => {
        if (generation !== this.lifecycleGeneration) {
          settled = true
          server.close()
          reject(new Error('MCP server start cancelled'))
          return
        }
        settled = true
        this.httpServer = server
        resolve()
      })
    })
  }

  /** Best-effort synchronous teardown for app shutdown. */
  stopSync(): void {
    void this.stop()
  }

  async stop(): Promise<void> {
    ++this.lifecycleGeneration
    if (!this.httpServer) {
      this.running = false
      return
    }

    for (const transport of this.legacyTransports.values()) {
      try {
        void transport.close()
      } catch {
        // ignore
      }
    }
    this.legacyTransports.clear()
    for (const transport of this.streamableTransports.values()) {
      try {
        await transport.close()
      } catch {
        // ignore
      }
    }
    this.streamableTransports.clear()
    this.sessionServers.clear()

    for (const socket of this.sockets) socket.destroy()
    this.sockets.clear()

    const server = this.httpServer
    this.httpServer = undefined
    this.running = false
    await new Promise<void>((resolve) => server.close(() => resolve()))
    this.logger('[mcp] stopped')
  }
}

function stringifyResult(result: unknown): string {
  if (typeof result === 'string') return result
  try {
    return JSON.stringify(result, null, 2)
  } catch {
    return String(result)
  }
}

function headerValue(value: string | string[] | undefined): string | undefined {
  if (Array.isArray(value)) return value[0]
  return value
}

function isInitializeRequest(body: unknown): boolean {
  if (Array.isArray(body))
    return body.some((msg) => (msg as { method?: string })?.method === 'initialize')
  return (body as { method?: string })?.method === 'initialize'
}

function readJsonBody(req: IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    let size = 0
    req.on('data', (chunk: Buffer) => {
      size += chunk.length
      if (size > MAX_BODY_BYTES) {
        reject(new Error('request body too large'))
        req.destroy()
        return
      }
      chunks.push(chunk)
    })
    req.on('end', () => {
      const raw = Buffer.concat(chunks).toString('utf8')
      if (!raw) {
        resolve(undefined)
        return
      }
      try {
        resolve(JSON.parse(raw))
      } catch (error) {
        reject(error)
      }
    })
    req.on('error', reject)
  })
}
