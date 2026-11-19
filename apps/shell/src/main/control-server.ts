import { createHash, randomBytes, timingSafeEqual } from 'node:crypto'
import { chmodSync, unlinkSync, writeFileSync } from 'node:fs'
import { createServer, type Server, type Socket } from 'node:net'
import { join } from 'node:path'
import {
  CONTROL_FILE,
  CONTROL_MAX_REQUEST_BYTES,
  CONTROL_PROTOCOL,
  type ControlEndpoint,
  type ControlEnvelope,
  type ControlReply,
  type ControlRequest,
} from '@genoffice/cli/control-protocol'

/**
 * Local request/reply channel for the genoffice CLI (`open --slide`,
 * `selection`). Security model: the socket lives in userData (per-user
 * directory, 0600 on POSIX; a per-user named pipe on Windows), every request
 * must carry the random token from userData/control.json (0600), one request
 * per connection, 64 KB cap, and the only commands are "show this spot" and
 * "read the selection" — nothing writes a file or runs code from the request.
 */
export type ControlHandler = (request: ControlRequest) => Promise<ControlReply>

export interface ControlServer {
  endpoint: ControlEndpoint
  close(): void
}

export function controlEndpointPath(userData: string): string {
  if (process.platform === 'win32') {
    const hash = createHash('sha1').update(userData).digest('hex').slice(0, 16)
    return `\\\\.\\pipe\\genoffice-control-${hash}`
  }
  return join(userData, 'control.sock')
}

export function parseEnvelope(line: string): ControlEnvelope | null {
  let raw: unknown
  try {
    raw = JSON.parse(line)
  } catch {
    return null
  }
  if (!raw || typeof raw !== 'object') return null
  const env = raw as Partial<ControlEnvelope>
  if (typeof env.token !== 'string' || !env.request || typeof env.request !== 'object') return null
  const req = env.request as Partial<ControlRequest>
  if ((req.cmd !== 'open' && req.cmd !== 'selection') || typeof req.path !== 'string') return null
  return env as ControlEnvelope
}

function tokenMatches(given: string, expected: string): boolean {
  const a = Buffer.from(given)
  const b = Buffer.from(expected)
  return a.length === b.length && timingSafeEqual(a, b)
}

export function startControlServer(
  userData: string,
  handler: ControlHandler,
): Promise<ControlServer> {
  const endpoint = controlEndpointPath(userData)
  const token = randomBytes(24).toString('hex')
  const file = join(userData, CONTROL_FILE)
  if (process.platform !== 'win32') {
    try {
      unlinkSync(endpoint)
    } catch {}
  }
  const server: Server = createServer((socket) => serve(socket, token, handler))
  return new Promise((resolve, reject) => {
    server.once('error', reject)
    server.listen(endpoint, () => {
      server.removeListener('error', reject)
      server.on('error', () => {})
      if (process.platform !== 'win32') {
        try {
          chmodSync(endpoint, 0o600)
        } catch {}
      }
      const info: ControlEndpoint = {
        protocol: CONTROL_PROTOCOL,
        pid: process.pid,
        endpoint,
        token,
      }
      writeFileSync(file, JSON.stringify(info), { mode: 0o600 })
      resolve({
        endpoint: info,
        close() {
          server.close()
          try {
            unlinkSync(file)
          } catch {}
          if (process.platform !== 'win32') {
            try {
              unlinkSync(endpoint)
            } catch {}
          }
        },
      })
    })
  })
}

function serve(socket: Socket, token: string, handler: ControlHandler): void {
  let buffer = ''
  let done = false
  const finish = (reply: ControlReply | null) => {
    if (done) return
    done = true
    if (reply) socket.end(JSON.stringify(reply) + '\n')
    else socket.destroy()
  }
  socket.setEncoding('utf8')
  socket.setTimeout(30_000, () => finish(null))
  socket.on('error', () => finish(null))
  socket.on('data', (chunk: string) => {
    if (done) return
    buffer += chunk
    if (buffer.length > CONTROL_MAX_REQUEST_BYTES) return finish(null)
    const nl = buffer.indexOf('\n')
    if (nl < 0) return
    const envelope = parseEnvelope(buffer.slice(0, nl))
    if (!envelope || !tokenMatches(envelope.token, token)) return finish(null)
    handler(envelope.request).then(finish, (err: unknown) =>
      finish({
        ok: false,
        error: {
          reason: 'app_unavailable',
          message: err instanceof Error ? err.message : String(err),
        },
      }),
    )
  })
}
