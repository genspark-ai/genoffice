#!/usr/bin/env node

/**
 * GenOffice MCP stdio bridge.
 *
 * MCP clients that only speak stdio (Claude Desktop, Cursor) can't reach the
 * GenOffice localhost server directly. This script bridges stdin/stdout
 * JSON-RPC to the running app's legacy SSE transport (/sse + /messages).
 *
 * The GenOffice app must be running with the MCP server enabled
 * (Settings > General > Local MCP server).
 *
 * Usage:
 *   node scripts/mcp-stdio-bridge.js [--port 3093] [--host 127.0.0.1]
 *
 * Example mcp.json entry:
 *   {
 *     "mcpServers": {
 *       "genoffice": {
 *         "command": "node",
 *         "args": ["/path/to/genoffice/scripts/mcp-stdio-bridge.js"]
 *       }
 *     }
 *   }
 */

const http = require('node:http')
const readline = require('node:readline')

const DEFAULT_PORT = 3093
const DEFAULT_HOST = '127.0.0.1'

function parseArgs(argv) {
  let port = Number(process.env.GENOFFICE_MCP_PORT) || DEFAULT_PORT
  let host = process.env.GENOFFICE_MCP_HOST || DEFAULT_HOST
  for (let i = 2; i < argv.length; i++) {
    if (argv[i] === '--port' && argv[i + 1]) port = parseInt(argv[++i], 10)
    else if (argv[i] === '--host' && argv[i + 1]) host = argv[++i]
  }
  return { port, host }
}

const config = parseArgs(process.argv)
const baseUrl = `http://${config.host}:${config.port}`

let sessionId = null
let sseConnection = null
let connectingPromise = null
let reconnectTimer = null
let stopping = false

// The client performs the MCP handshake once, against the first SSE session.
// A reconnect creates a *new*, uninitialized server session, so the cached
// handshake is replayed on it — otherwise every tool call after a transient
// disconnect fails as "not initialized" while the client believes it is ready.
// Responses to a replayed initialize are dropped: the client already has its
// handshake result and would see a duplicate response.
let handshakeMessages = null
let handshakeSession = null
const suppressedResponseIds = new Set()

/** stderr only: stdout is the JSON-RPC channel */
function log(message) {
  process.stderr.write(`[genoffice-mcp-bridge] ${message}\n`)
}

function sendResponse(response) {
  process.stdout.write(JSON.stringify(response) + '\n')
}

function httpRequest(method, path, body) {
  return new Promise((resolve, reject) => {
    const url = new URL(path, baseUrl)
    const req = http.request(
      {
        hostname: url.hostname,
        port: url.port,
        path: url.pathname + url.search,
        method,
        headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
      },
      (res) => {
        let data = ''
        res.on('data', (chunk) => (data += chunk))
        res.on('end', () => {
          try {
            resolve({ status: res.statusCode, data: data ? JSON.parse(data) : null })
          } catch {
            resolve({ status: res.statusCode, data })
          }
        })
      },
    )
    req.on('error', reject)
    if (body) req.write(JSON.stringify(body))
    req.end()
  })
}

function connectSSE() {
  if (sessionId && sseConnection && !sseConnection.destroyed) return Promise.resolve(sessionId)
  if (connectingPromise) return connectingPromise

  connectingPromise = new Promise((resolve, reject) => {
    const url = new URL('/sse', baseUrl)
    let settled = false
    let sessionTimeout

    log(`connecting to ${url.href}`)
    const req = http.get(url.href, (res) => {
      if (res.statusCode !== 200) {
        res.resume()
        settled = true
        reject(new Error(`SSE connection failed: ${res.statusCode}`))
        return
      }

      sseConnection = res
      let buffer = ''

      sessionTimeout = setTimeout(() => {
        if (!settled) {
          settled = true
          req.destroy()
          reject(new Error('timed out waiting for the SSE session id'))
        }
      }, 5000)

      res.on('data', (chunk) => {
        buffer += chunk.toString()
        const lines = buffer.split('\n')
        buffer = lines.pop() || ''
        for (const line of lines) {
          if (line.startsWith('data: ')) {
            try {
              const parsed = JSON.parse(line.slice(6))
              // drop the answer to a replayed handshake
              if (parsed && parsed.id !== undefined && suppressedResponseIds.has(parsed.id)) {
                suppressedResponseIds.delete(parsed.id)
                continue
              }
              sendResponse(parsed)
            } catch {
              // endpoint URL / heartbeat comment: not JSON-RPC
            }
          }
          if (line.includes('sessionId=')) {
            const match = line.match(/sessionId=([a-zA-Z0-9-]+)/)
            if (match) {
              sessionId = match[1]
              log(`session ${sessionId}`)
              if (!settled) {
                settled = true
                clearTimeout(sessionTimeout)
                resolve(sessionId)
              }
            }
          }
        }
      })

      res.on('error', (err) => log(`SSE error: ${err.message}`))
      res.on('close', () => {
        clearTimeout(sessionTimeout)
        sseConnection = null
        sessionId = null
        if (!settled) {
          settled = true
          reject(new Error('SSE connection closed before the session initialized'))
        }
        if (!stopping) scheduleReconnect()
      })
    })

    req.on('error', (error) => {
      clearTimeout(sessionTimeout)
      if (!settled) {
        settled = true
        reject(error)
      }
    })
  }).finally(() => {
    connectingPromise = null
  })

  return connectingPromise
}

function scheduleReconnect() {
  if (stopping || reconnectTimer) return
  reconnectTimer = setTimeout(() => {
    reconnectTimer = null
    connectSSE().catch((error) => {
      log(`reconnect failed: ${error.message}`)
      scheduleReconnect()
    })
  }, 5000)
}

/** Remember the client's handshake so a reconnect can re-establish it. */
function cacheHandshake(message) {
  const method = message && typeof message === 'object' ? message.method : undefined
  if (method !== 'initialize' && method !== 'notifications/initialized') return
  const cached = (handshakeMessages ?? []).filter((m) => m.method !== method)
  cached.push(message)
  handshakeMessages = cached
  // the exchanged pair completes the handshake on the session that carried it
  if (method === 'notifications/initialized') handshakeSession = sessionId
}

/** Re-run the cached handshake against a freshly created session. */
async function replayHandshake() {
  const messages = handshakeMessages
  const target = sessionId
  if (!messages || messages.length === 0) {
    handshakeSession = target
    return
  }
  // Mark the session as handshaked up front: sendToServer below would otherwise
  // see the stale handshakeSession and re-enter this function.
  handshakeSession = target
  log('replaying session handshake after reconnect')
  for (const message of messages) {
    if (message.id !== undefined && message.id !== null) {
      suppressedResponseIds.add(message.id)
    }
    try {
      await sendToServer(message)
    } catch (error) {
      log(`handshake replay failed: ${error.message}`)
    }
  }
}

async function sendToServer(message) {
  if (!sessionId) await connectSSE()
  // a reconnect hands us a new, uninitialized session: re-establish the cached
  // handshake before the client's next call lands on it
  if (handshakeSession !== null && handshakeSession !== sessionId) await replayHandshake()
  cacheHandshake(message)
  // legacy SSE delivers responses only on the stream; the POST ack body must
  // not be written to stdout or it corrupts the newline-delimited protocol
  const response = await httpRequest('POST', `/messages?sessionId=${sessionId}`, message)
  if (response.status < 200 || response.status >= 300) {
    const detail = typeof response.data === 'string' ? response.data : JSON.stringify(response.data)
    throw new Error(`server returned HTTP ${response.status}${detail ? `: ${detail}` : ''}`)
  }
}

async function handleRequest(request) {
  try {
    await sendToServer(request)
  } catch (error) {
    // A notification (no id) must never be answered: an unsolicited line would
    // read as a protocol violation to the client. Only log it.
    if (request.id === undefined || request.id === null) {
      log(`notification ${request.method} failed: ${error.message}`)
      return
    }
    sendResponse({
      jsonrpc: '2.0',
      id: request.id,
      error: { code: -32603, message: error.message },
    })
  }
}

async function main() {
  log(`bridging stdio to ${baseUrl}`)
  try {
    const health = await httpRequest('GET', '/health')
    if (health.status === 200) log('GenOffice MCP server is reachable')
    else log('warning: unexpected /health response; is GenOffice running?')
  } catch {
    log(
      'warning: cannot reach GenOffice. Start the app and enable Settings > General > Local MCP server.',
    )
  }

  try {
    await connectSSE()
  } catch (error) {
    log(`SSE connection error: ${error.message}`)
  }

  const rl = readline.createInterface({ input: process.stdin, terminal: false })
  rl.on('line', (line) => {
    if (!line.trim()) return
    try {
      void handleRequest(JSON.parse(line))
    } catch (error) {
      log(`invalid JSON: ${error.message}`)
    }
  })

  const shutdown = (reason) => {
    if (stopping) return
    stopping = true
    log(`${reason}, exiting`)
    if (reconnectTimer) clearTimeout(reconnectTimer)
    if (sseConnection) sseConnection.destroy()
    process.exit(0)
  }
  rl.on('close', () => shutdown('stdin closed'))
  process.on('SIGINT', () => shutdown('SIGINT'))
  process.on('SIGTERM', () => shutdown('SIGTERM'))
}

main().catch((error) => {
  log(`fatal: ${error.message}`)
  process.exit(1)
})
