/**
 * GenOffice Web Server — standalone HTTP service that mirrors the Electron
 * main-process IPC surface so the same renderer code can run in a browser
 * without launching Electron.
 *
 * The boot path is intentionally thin:
 *
 *   1. Wire every capability module into the shared handler registry.
 *   2. Bring up a single HTTP server that multiplexes:
 *        - GET  /health                 — health probe
 *        - GET  /api/channels           — list registered channels
 *        - GET  /api/collab/sessions    — collab session snapshot
 *        - POST /api/ipc/:channel       — JSON-over-HTTP IPC invoke
 *        - GET  /api/ipc/events         — SSE event stream per session
 *        - POST /api/ai/stream          — Agent-Loop SSE stream
 *      and falls back to the static docs renderer when no API route hits.
 *
 * The capability code lives in capability-specific sub-directories
 * (`apps/web-server/src/{ai,projects,docs,slides,sheets,pdf,markdown,shell,
 * collab,enterprise,common,anydoc,web}`). See LUM-553 for the refactor plan.
 */
import { createServer, type IncomingMessage, ServerResponse } from 'node:http'
import { createReadStream, existsSync, statSync } from 'node:fs'
import { resolve, extname } from 'node:path'

import {
  APPS,
  COLLAB_SESSIONS,
  HOST,
  MIME_TYPES,
  PORT,
  STATIC_ROOT,
  decodeTransportValue,
  initRecentState,
  encodeTransportValue,
  getHandler,
  handlerCount,
  listChannels,
} from './common/index.js'
import {
  registerAiHandlers,
  AI_STREAM_SESSIONS,
  runProviderStream,
} from './ai/index.js'
import type { AiSettings, AiStreamChunk } from '@genoffice/ai-provider'
import { registerProjectHandlers } from './projects/index.js'
import { registerDocsHandlers } from './docs/index.js'
import { registerSheetsHandlers } from './sheets/index.js'
import { registerSlidesHandlers } from './slides/index.js'
import { registerPdfHandlers } from './pdf/index.js'
import { registerMarkdownHandlers } from './markdown/index.js'
import { registerHtmlHandlers } from './html/index.js'
import { registerShellHandlers } from './shell/index.js'
import { registerCollabHandlers } from './collab/index.js'
import { registerEnterpriseHandlers } from './enterprise/index.js'
import { registerAnydocHandlers } from './anydoc/index.js'
import { registerWebHandlers } from './web/index.js'

// ----- capability wiring ----------------------------------------------------
initRecentState()
registerAiHandlers()
registerProjectHandlers()
registerDocsHandlers()
registerSheetsHandlers()
registerSlidesHandlers()
registerPdfHandlers()
registerMarkdownHandlers()
registerHtmlHandlers()
registerShellHandlers()
registerCollabHandlers()
registerEnterpriseHandlers()
registerAnydocHandlers()
registerWebHandlers()

// ----- HTTP helpers --------------------------------------------------------
function sendJson(response: ServerResponse, status: number, payload: unknown): void {
  response.writeHead(status, { 'Content-Type': 'application/json' })
  response.end(JSON.stringify(payload))
}

function sendIpcError(response: ServerResponse, error: unknown, structuredAware: boolean = true): void {
  const errObj: { message: string; code?: string; channel?: string; reason?: string } = {
    message: (error as Error)?.message ? String((error as Error).message) : String(error),
  }
  if (structuredAware) {
    const anyErr = error as { code?: unknown; channel?: unknown; reason?: unknown }
    if (typeof anyErr?.code === 'string') errObj.code = anyErr.code
    if (typeof anyErr?.channel === 'string') errObj.channel = anyErr.channel
    if (typeof anyErr?.reason === 'string') errObj.reason = anyErr.reason
  }
  const status = errObj.code === 'WEB_UNSUPPORTED' ? 501 : 500
  sendJson(response, status, { error: errObj })
}

async function readBody(request: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    request.on('data', (chunk: Buffer) => chunks.push(chunk))
    request.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')))
    request.on('error', reject)
  })
}

// SSE plumbing — mirrors the legacy single-file implementation. Each
// session keeps its own Set<ServerResponse> so renderer tabs share one
// push channel.
const sessionConnections = new Map<string, Set<ServerResponse>>()
const PENDING_FRAMES = new Map<string, string[]>()
const SSE_HEARTBEAT_MS = 25000

function pushSseEvent(session: string, channel: string, args: unknown[]): void {
  const encodedArgs = args.map((arg) => encodeTransportValue(arg))
  const frame = `data: ${JSON.stringify({ channel, args: encodedArgs })}\n\n`
  const connections = sessionConnections.get(session)
  if (connections) {
    for (const response of connections) {
      try {
        response.write(frame)
      } catch {}
    }
  } else {
    const pending = PENDING_FRAMES.get(session) || []
    pending.push(frame)
    if (pending.length > 100) pending.shift()
    PENDING_FRAMES.set(session, pending)
  }
}

// ----- request handling ----------------------------------------------------
const server = createServer(async (request, response) => {
  const url = new URL(request.url || '/', `http://${request.headers.host}`)

  response.setHeader('Access-Control-Allow-Origin', '*')
  response.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
  response.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-IPC-Session')

  if (request.method === 'OPTIONS') {
    response.writeHead(204)
    response.end()
    return
  }

  if (url.pathname === '/health' && request.method === 'GET') {
    sendJson(response, 200, {
      status: 'ok',
      version: '0.8.0',
      mode: 'web-server',
      implementedChannels: handlerCount(),
      features: ['ai', 'collab', 'files', 'projects'],
    })
    return
  }

  if (url.pathname === '/api/channels' && request.method === 'GET') {
    sendJson(response, 200, { channels: listChannels() })
    return
  }

  if (url.pathname === '/api/collab/sessions' && request.method === 'GET') {
    const sessions = [...COLLAB_SESSIONS.entries()].map(([docId, session]) => ({
      docId,
      users: [...session.users],
      lastActivity: session.lastActivity,
    }))
    sendJson(response, 200, { sessions })
    return
  }

  if (url.pathname.startsWith('/api/ipc/') && request.method === 'POST') {
    const encodedChannel = url.pathname.slice('/api/ipc/'.length)
    let channel: string
    try {
      channel = decodeURIComponent(encodedChannel)
    } catch {
      sendJson(response, 400, {
        error: { message: 'Invalid IPC channel encoding', code: 'IPC_INVALID_CHANNEL' },
      })
      return
    }
    const session = request.headers['x-ipc-session'] as string | undefined

    try {
      const body = await readBody(request)
      const { args = [] } = JSON.parse(body || '{}')
      const decodedArgs = (args as unknown[]).map((arg) => decodeTransportValue(arg))

      const handler = getHandler(channel)
      if (handler) {
        const event = {
          processId: 0,
          frameId: 0,
          sender: {
            id: -1,
            isDestroyed: () => false,
            send: (ch: string, ...a: unknown[]) => {
              const encodedArgs = a.map((arg) => encodeTransportValue(arg))
              if (session) pushSseEvent(session, ch, encodedArgs)
            },
          },
        }

        const result = await handler(event, ...decodedArgs)
        const encodedResult = encodeTransportValue(result)
        sendJson(response, 200, { ok: true, result: encodedResult })
      } else {
        sendJson(response, 404, {
          error: { message: `No handler for '${channel}'`, code: 'IPC_NO_HANDLER' },
        })
      }
    } catch (error) {
      sendIpcError(response, error)
    }
    return
  }

  if (url.pathname === '/api/ipc/events' && request.method === 'GET') {
    const session = url.searchParams.get('session')
    if (!session) {
      sendJson(response, 400, { error: { message: 'Missing session' } })
      return
    }

    response.writeHead(200, {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      Connection: 'keep-alive',
    })
    response.write(': connected\n\n')

    const pending = PENDING_FRAMES.get(session)
    if (pending) {
      for (const frame of pending) response.write(frame)
      PENDING_FRAMES.delete(session)
    }

    if (!sessionConnections.has(session)) {
      sessionConnections.set(session, new Set())
    }
    sessionConnections.get(session)!.add(response)

    const heartbeat = setInterval(() => {
      try {
        response.write(': heartbeat\n\n')
      } catch {
        clearInterval(heartbeat)
      }
    }, SSE_HEARTBEAT_MS)

    request.on('close', () => {
      clearInterval(heartbeat)
      sessionConnections.get(session)?.delete(response)
      if (sessionConnections.get(session)?.size === 0) {
        sessionConnections.delete(session)
      }
    })
    return
  }

  if (url.pathname === '/api/ai/stream' && request.method === 'POST') {
    let sessionAbort: AbortController | undefined
    let requestId: string | undefined
    try {
      const body = await readBody(request)
      const req = JSON.parse(body || '{}') as {
        requestId?: string
        sessionId?: string
        settings?: AiSettings
        system?: string
        messages?: Parameters<typeof runProviderStream>[2]
        tools?: Parameters<typeof runProviderStream>[3]
        maxTokens?: number
      }
      requestId = req.requestId || `sse-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
      // Import inside the handler to grab the live settings the AI module
      // has just persisted (avoids a duplicate cached copy).
      const { aiSettings } = await import('./ai/index.js').then((m) => m) as {
        aiSettings: AiSettings
      }
      void aiSettings // (kept to surface the live reference if needed)
      // The renderer can include its own settings override; otherwise use
      // the server's persisted ones.
      const settings: AiSettings = req.settings || (await import('./ai/chat.js' as string).catch(() => null))?.aiSettings
        || (await import('./ai/index.js') as { aiSettings?: AiSettings }).aiSettings
        || (req.settings as AiSettings)

      response.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        Connection: 'keep-alive',
        'X-Request-Id': requestId,
      })

      const send = (chunk: AiStreamChunk) => {
        try {
          response.write(`data: ${JSON.stringify({ ...chunk, requestId })}\n\n`)
        } catch {}
      }

      // Track the session so /api/ai/stream/cancel can abort it.
      sessionAbort = new AbortController()
      AI_STREAM_SESSIONS.set(requestId, { abort: sessionAbort, chunks: 0 })

      // If the client disconnects, stop the upstream call too.
      request.on('close', () => {
        sessionAbort?.abort()
        AI_STREAM_SESSIONS.delete(requestId)
      })

      await runProviderStream(settings, req.system || '', req.messages || [], req.tools || [], req.maxTokens, {
        onAbort: (c) => {
          sessionAbort = c
        },
        send,
      })
    } catch (error) {
      sendIpcError(response, error)
      sessionAbort?.abort()
    } finally {
      try { response.end() } catch {}
    }
    return
  }

  if (url.pathname === '/api/ai/stream/cancel' && request.method === 'POST') {
    try {
      const body = await readBody(request)
      const { requestId } = JSON.parse(body || '{}') as { requestId?: string }
      if (!requestId) {
        sendJson(response, 400, { error: { message: 'requestId required' } })
        return
      }
      const session = AI_STREAM_SESSIONS.get(requestId)
      if (session) {
        session.abort.abort()
        AI_STREAM_SESSIONS.delete(requestId)
      }
      sendJson(response, 200, { ok: true, aborted: !!session })
    } catch (error) {
      sendIpcError(response, error)
    }
    return
  }

  // ----- static / SPA fallback ---------------------------------------------
  const pathMatch = url.pathname.match(
    /^\/(docs|sheets|slides|pdf|markdown|html|shell)(?:\/(.*))?$/,
  )
  const isManagementRoute =
    url.pathname === '/' || url.pathname === '/manage' || url.pathname === '/management'
  const appName = isManagementRoute
    ? 'shell'
    : url.searchParams.get('app') || pathMatch?.[1] || 'shell'
  const relativePath = pathMatch
    ? pathMatch[2] || 'index.html'
    : url.pathname.replace(/^\/+/, '') || 'index.html'
  let filePath = resolve(STATIC_ROOT, appName, 'out', 'renderer', relativePath)

  if (!existsSync(filePath)) {
    filePath = resolve(STATIC_ROOT, 'docs', 'out', 'renderer', relativePath)
  }

  if (!existsSync(filePath) && relativePath.startsWith('assets/')) {
    for (const candidateApp of APPS) {
      const candidatePath = resolve(STATIC_ROOT, candidateApp, 'out', 'renderer', relativePath)
      if (existsSync(candidatePath)) {
        filePath = candidatePath
        break
      }
    }
  }

  if (existsSync(filePath) && statSync(filePath).isFile()) {
    const ext = extname(filePath)
    response.writeHead(200, { 'Content-Type': MIME_TYPES[ext] || 'application/octet-stream' })
    createReadStream(filePath).pipe(response)
    return
  }

  const indexPath = resolve(STATIC_ROOT, appName, 'out', 'renderer', 'index.html')
  if (existsSync(indexPath)) {
    response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' })
    createReadStream(indexPath).pipe(response)
    return
  }

  response.writeHead(404, { 'Content-Type': 'text/html; charset=utf-8' })
  response.end(
    `<!doctype html><meta charset="utf-8"><title>GenOffice Web Server</title>` +
    `<style>body{font-family:system-ui;max-width:640px;margin:48px auto;padding:0 24px;color:#222;line-height:1.55}` +
    `code{background:#f4f4f4;padding:2px 6px;border-radius:3px;font-size:0.92em}</style>` +
    `<h1>GenOffice Web Server</h1>` +
    `<p>The renderer apps were not found at <code>${STATIC_ROOT}</code>.</p>` +
    `<p>Either run <code>npm run build:all</code> at the repo root and keep it on the same disk layout, ` +
    `or set <code>WEB_STATIC_ROOT=/path/to/apps</code> to point at an apps directory you mounted.</p>`,
  )
})

// Detect whether the renderer apps are available at STATIC_ROOT. When the
// binary is shipped standalone (pkg) the operator is expected to set
// WEB_STATIC_ROOT; in dev the apps live next to the source. We surface
// this in the boot log so misconfiguration is obvious instead of silent.
const shellIndex = resolve(STATIC_ROOT, 'shell', 'out', 'renderer', 'index.html')
const staticReady = existsSync(shellIndex)
const staticHint = staticReady
  ? `║   📁 Static root: ${STATIC_ROOT}                    ║\n`
  : `║   ⚠️  No renderer apps at ${STATIC_ROOT}          ║\n` +
    `║      Set WEB_STATIC_ROOT=/path/to/apps or run npm run build:all  ║\n`

server.listen(PORT, HOST, () => {
  console.log(`
╔═══════════════════════════════════════════════════════════╗
║                                                           ║
║   GenOffice Web Server v0.8.0 (Enhanced)                ║
║                                                           ║
    URL: http://${HOST}:${PORT}
║   📁 Mode: Standalone (No Electron)                        ║
║                                                           ║
║   Apps: ${APPS.slice(0, 4).join(', ')}...
${staticHint}║                                                           ║
║   📊 Channels: ${String(handlerCount()).padEnd(25)}   ║
║   🔗 Features: AI, Collab, Files, Projects, AnyDoc        ║
║                                                           ║
║   Endpoints:                                              ║
║   • GET  /health              Health check                 ║
║   • GET  /api/channels       List channels                ║
║   • POST /api/ai/stream       Agent Loop SSE               ║
║   • GET  /api/collab/sessions Collaboration status        ║
║   • POST /api/ipc/:channel   IPC invoke                  ║
║   • GET  /api/ipc/events     SSE events                  ║
║                                                           ║
║   Agent Core Integration:                                 ║
║   ✅ createHttpTransport()  - HTTP Transport for AgentLoop ║
║                                                           ║
╚═══════════════════════════════════════════════════════════╝
`)
})

process.on('SIGTERM', () => {
  server.close(() => process.exit(0))
})
process.on('SIGINT', () => {
  server.close(() => process.exit(0))
})
