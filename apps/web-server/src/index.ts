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
  encodeTransportValue,
  getHandler,
  handlerCount,
  listChannels,
} from './common/index.js'
import { registerAiHandlers, generateAgentResponse } from './ai/index.js'
import { registerProjectHandlers } from './projects/index.js'
import { registerDocsHandlers } from './docs/index.js'
import { registerSheetsHandlers } from './sheets/index.js'
import { registerSlidesHandlers } from './slides/index.js'
import { registerPdfHandlers } from './pdf/index.js'
import { registerMarkdownHandlers } from './markdown/index.js'
import { registerShellHandlers } from './shell/index.js'
import { registerCollabHandlers } from './collab/index.js'
import { registerEnterpriseHandlers } from './enterprise/index.js'
import { registerAnydocHandlers } from './anydoc/index.js'
import { registerWebHandlers } from './web/index.js'

// ----- capability wiring ----------------------------------------------------
registerAiHandlers()
registerProjectHandlers()
registerDocsHandlers()
registerSheetsHandlers()
registerSlidesHandlers()
registerPdfHandlers()
registerMarkdownHandlers()
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
  const encodedArgs = args.map(arg => encodeTransportValue(arg))
  const frame = `data: ${JSON.stringify({ channel, args: encodedArgs })}\n\n`
  const connections = sessionConnections.get(session)
  if (connections) {
    for (const response of connections) {
      try { response.write(frame) } catch {}
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
    const channel = url.pathname.slice('/api/ipc/'.length)
    const session = request.headers['x-ipc-session'] as string | undefined

    try {
      const body = await readBody(request)
      const { args = [] } = JSON.parse(body || '{}')
      const decodedArgs = (args as unknown[]).map(arg => decodeTransportValue(arg))

      const handler = getHandler(channel)
      if (handler) {
        const event = {
          processId: 0,
          frameId: 0,
          sender: {
            id: -1,
            isDestroyed: () => false,
            send: (ch: string, ...a: unknown[]) => {
              const encodedArgs = a.map(arg => encodeTransportValue(arg))
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
      sendJson(response, 500, { error: { message: String((error as Error)?.message) } })
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
      'Connection': 'keep-alive',
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
      try { response.write(': heartbeat\n\n') } catch { clearInterval(heartbeat) }
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
    try {
      const body = await readBody(request)
      const { requestId, system, messages, tools } = JSON.parse(body || '{}')

      response.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'X-Request-Id': requestId || '',
      })

      const responseText = generateAgentResponse(messages || [])
      const words = responseText.split(/([\s，。、！？]+)/)
      let delay = 50

      const streamChunk = (type: string, data: Record<string, unknown>) => {
        response.write(`data: ${JSON.stringify({ requestId, type, ...data })}\n\n`)
      }

      const pingInterval = setInterval(() => {
        try { response.write(`data: ${JSON.stringify({ requestId, type: 'ping' })}\n\n`) } catch {}
      }, 30000)

      let wordIndex = 0
      const sendWord = () => {
        if (wordIndex >= words.length) {
          clearInterval(pingInterval)
          response.write(`data: ${JSON.stringify({ requestId, type: 'done', stopReason: 'stop' })}\n\n`)
          response.end()
          return
        }

        streamChunk('delta', { text: words[wordIndex] })
        wordIndex++

        if (wordIndex === Math.floor(words.length / 2) && tools && tools.length > 0) {
          const toolCall = {
            id: `tool-${Date.now()}`,
            name: tools[0].name,
            input: {},
          }
          streamChunk('tool-call', { toolCall })
        }

        setTimeout(sendWord, delay)
      }

      sendWord()

      request.on('close', () => {
        clearInterval(pingInterval)
      })
    } catch (error) {
      sendJson(response, 500, { error: { message: String((error as Error)?.message) } })
    }
    return
  }

  // ----- static / SPA fallback ---------------------------------------------
  const appName = url.searchParams.get('app') || 'docs'
  let filePath = resolve(STATIC_ROOT, appName, 'out', 'renderer', url.pathname === '/' ? 'index.html' : url.pathname)

  if (!existsSync(filePath)) {
    filePath = resolve(STATIC_ROOT, 'docs', 'out', 'renderer', url.pathname === '/' ? 'index.html' : url.pathname)
  }

  if (existsSync(filePath) && statSync(filePath).isFile()) {
    const ext = extname(filePath)
    response.writeHead(200, { 'Content-Type': MIME_TYPES[ext] || 'application/octet-stream' })
    createReadStream(filePath).pipe(response)
    return
  }

  const indexPath = resolve(STATIC_ROOT, 'docs', 'out', 'renderer', 'index.html')
  if (existsSync(indexPath)) {
    response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' })
    createReadStream(indexPath).pipe(response)
    return
  }

  response.writeHead(404, { 'Content-Type': 'text/html; charset=utf-8' })
  response.end('<h1>GenOffice Web Server</h1><p>Please build the apps first: npm run build:all</p>')
})

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
║                                                           ║
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

process.on('SIGTERM', () => { server.close(() => process.exit(0)) })
process.on('SIGINT', () => { server.close(() => process.exit(0)) })
