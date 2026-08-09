/**
 * A loopback HTTP `CONNECT` proxy that forwards to a SOCKS5 proxy.
 *
 * The gsk CLI is a child process on Node's built-in fetch, which only honours
 * a proxy through `NODE_USE_ENV_PROXY` — and that is undici's env proxy
 * support, which speaks HTTP `CONNECT` and nothing else. So a user whose
 * settings say `socks5://127.0.0.1:7897` got a working browser and a working
 * main process while the CLI dialled genspark.ai direct, which on the kind of
 * network that needs a proxy means every CLI call fails.
 *
 * This bridge hands the child an `http://` proxy on loopback and performs the
 * SOCKS5 handshake on its behalf. Deliberately narrow: `CONNECT` only (every
 * genspark endpoint is https, so the CLI never needs plain-HTTP forwarding),
 * bound to 127.0.0.1 so nothing off-machine can reach it, and torn down
 * whenever the proxy changes.
 */
import { createServer, type Server } from 'node:http'
import { connect, type Socket } from 'node:net'

/** SOCKS4 has no domain-name mode worth supporting here; callers fall back to direct. */
const SUPPORTED_SCHEMES = new Set(['socks5:', 'socks:'])

export function isBridgeableSocksUrl(url: string): boolean {
  try {
    return SUPPORTED_SCHEMES.has(new URL(url).protocol.toLowerCase())
  } catch {
    return false
  }
}

/** Reads exactly N bytes at a time off a socket, so the handshake can be written sequentially. */
class ByteReader {
  private buffer: Buffer = Buffer.alloc(0)
  private want = 0
  private settle: ((value: Buffer) => void) | null = null
  private fail: ((error: Error) => void) | null = null
  private error: Error | null = null

  constructor(socket: Socket) {
    socket.on('data', (chunk: Buffer) => {
      this.buffer = Buffer.concat([this.buffer, chunk])
      this.drain()
    })
    socket.on('error', (err: Error) => this.abort(err))
    socket.on('close', () => this.abort(new Error('SOCKS proxy closed the connection')))
  }

  read(count: number): Promise<Buffer> {
    if (this.error) return Promise.reject(this.error)
    return new Promise<Buffer>((resolve, reject) => {
      this.want = count
      this.settle = resolve
      this.fail = reject
      this.drain()
    })
  }

  private drain(): void {
    if (!this.settle || this.buffer.length < this.want) return
    const taken = this.buffer.subarray(0, this.want)
    this.buffer = this.buffer.subarray(this.want)
    const resolve = this.settle
    this.settle = null
    this.fail = null
    resolve(taken)
  }

  private abort(err: Error): void {
    this.error ??= err
    const reject = this.fail
    this.settle = null
    this.fail = null
    reject?.(err)
  }
}

/** RFC 1928 reply codes, so a failure says what the proxy actually refused. */
const SOCKS_REPLY: Record<number, string> = {
  1: 'general SOCKS server failure',
  2: 'connection not allowed by ruleset',
  3: 'network unreachable',
  4: 'host unreachable',
  5: 'connection refused',
  6: 'TTL expired',
  7: 'command not supported',
  8: 'address type not supported',
}

/** Open a tunnel to host:port through a SOCKS5 proxy (RFC 1928, with RFC 1929 auth). */
async function socks5Tunnel(proxy: URL, host: string, port: number): Promise<Socket> {
  const socket = connect({
    host: proxy.hostname,
    port: Number(proxy.port) || 1080,
  })
  const reader = new ByteReader(socket)
  try {
    await new Promise<void>((resolve, reject) => {
      socket.once('connect', resolve)
      socket.once('error', reject)
    })

    const user = decodeURIComponent(proxy.username)
    const pass = decodeURIComponent(proxy.password)
    const credentials = user.length > 0 || pass.length > 0
    // offer no-auth, plus username/password when the url carries credentials
    socket.write(credentials ? Buffer.from([5, 2, 0, 2]) : Buffer.from([5, 1, 0]))
    const greeting = await reader.read(2)
    if (greeting[0] !== 5) throw new Error('SOCKS proxy is not speaking SOCKS5')
    const method = greeting[1]
    if (method === 2) {
      if (!credentials) throw new Error('SOCKS proxy wants a username and password')
      const u = Buffer.from(user, 'utf8')
      const p = Buffer.from(pass, 'utf8')
      socket.write(Buffer.concat([Buffer.from([1, u.length]), u, Buffer.from([p.length]), p]))
      const auth = await reader.read(2)
      if (auth[1] !== 0) throw new Error('SOCKS proxy rejected the username and password')
    } else if (method !== 0) {
      throw new Error('SOCKS proxy offered no authentication method we support')
    }

    // CONNECT by domain name: the proxy resolves, which is the point of using it
    const name = Buffer.from(host, 'utf8')
    if (name.length > 255) throw new Error('Host name is too long for SOCKS5')
    const request = Buffer.concat([
      Buffer.from([5, 1, 0, 3, name.length]),
      name,
      Buffer.from([(port >> 8) & 0xff, port & 0xff]),
    ])
    socket.write(request)

    const reply = await reader.read(4)
    if (reply[1] !== 0) {
      throw new Error(`SOCKS proxy refused the connection: ${SOCKS_REPLY[reply[1]!] ?? 'unknown'}`)
    }
    // drain the bound address so the tunnel starts at the first payload byte
    const type = reply[3]
    if (type === 1) await reader.read(4 + 2)
    else if (type === 4) await reader.read(16 + 2)
    else if (type === 3) {
      const length = await reader.read(1)
      await reader.read(length[0]! + 2)
    } else throw new Error('SOCKS proxy replied with an unknown address type')

    return socket
  } catch (err) {
    socket.destroy()
    throw err
  }
}

let server: Server | null = null
let listeningUrl = ''
let startedFor = ''

/**
 * Start (or reuse) the bridge for a SOCKS url and return the loopback
 * `http://` proxy the child should use. Returns '' when the url is not one we
 * can bridge, so callers fall back to no proxy rather than a broken one.
 */
export async function startSocksBridge(socksUrl: string): Promise<string> {
  if (!isBridgeableSocksUrl(socksUrl)) return ''
  if (server && startedFor === socksUrl) return listeningUrl
  stopSocksBridge()
  const proxy = new URL(socksUrl)

  const next = createServer()
  // A plain request means something other than CONNECT reached us; the CLI
  // only ever tunnels https, so answer rather than leave the socket hanging.
  next.on('request', (_req, res) => {
    res.writeHead(405, { Connection: 'close' })
    res.end('This proxy only supports CONNECT.\n')
  })
  next.on('connect', (req, clientSocket, head) => {
    const [host, rawPort] = splitHostPort(req.url ?? '')
    if (!host) {
      clientSocket.end('HTTP/1.1 400 Bad Request\r\n\r\n')
      return
    }
    socks5Tunnel(proxy, host, rawPort)
      .then((upstream) => {
        clientSocket.write('HTTP/1.1 200 Connection Established\r\n\r\n')
        if (head.length > 0) upstream.write(head)
        upstream.pipe(clientSocket)
        clientSocket.pipe(upstream)
        const drop = () => {
          upstream.destroy()
          clientSocket.destroy()
        }
        upstream.on('error', drop)
        clientSocket.on('error', drop)
      })
      .catch((err: unknown) => {
        console.warn('[proxy] SOCKS bridge could not reach', host, '-', String(err))
        clientSocket.end('HTTP/1.1 502 Bad Gateway\r\n\r\n')
      })
  })
  next.on('error', (err) => console.warn('[proxy] SOCKS bridge server error:', err))

  await new Promise<void>((resolve, reject) => {
    next.once('error', reject)
    // 127.0.0.1 only: this is an open proxy, and it must stay on this machine
    next.listen(0, '127.0.0.1', resolve)
  })
  // never let an idle bridge hold the app open at quit; an in-flight tunnel
  // still keeps its own sockets referenced
  next.unref()
  const address = next.address()
  if (address === null || typeof address === 'string') {
    next.close()
    return ''
  }
  server = next
  startedFor = socksUrl
  listeningUrl = `http://127.0.0.1:${address.port}`
  return listeningUrl
}

export function stopSocksBridge(): void {
  server?.close()
  server = null
  listeningUrl = ''
  startedFor = ''
}

/** `host:port`, or a bracketed IPv6 literal, as CONNECT sends it. */
export function splitHostPort(target: string): [host: string, port: number] {
  const bracketed = /^\[([^\]]+)\]:(\d+)$/.exec(target)
  if (bracketed) return [bracketed[1]!, Number(bracketed[2])]
  const index = target.lastIndexOf(':')
  if (index <= 0) return ['', 0]
  const port = Number(target.slice(index + 1))
  return Number.isInteger(port) && port > 0 && port <= 65535
    ? [target.slice(0, index), port]
    : ['', 0]
}
