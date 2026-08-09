import { createServer as createTcpServer, connect, type Server, type Socket } from 'node:net'

import { afterEach, describe, expect, it } from 'vitest'

import {
  isBridgeableSocksUrl,
  splitHostPort,
  startSocksBridge,
  stopSocksBridge,
} from '../src/socks-bridge'

const closers: Array<() => void> = []

afterEach(() => {
  stopSocksBridge()
  while (closers.length > 0) closers.pop()?.()
})

function listen(server: Server): Promise<number> {
  return new Promise((resolve) => {
    server.listen(0, '127.0.0.1', () => {
      const address = server.address()
      resolve(typeof address === 'object' && address !== null ? address.port : 0)
    })
  })
}

/** An origin that greets whoever connects, so the tunnel can be proven end to end. */
async function startOrigin(): Promise<number> {
  const server = createTcpServer((socket) => {
    socket.on('data', (chunk) => socket.write(`echo:${chunk.toString('utf8')}`))
    socket.write('hello\n')
  })
  closers.push(() => server.close())
  return listen(server)
}

interface SocksOptions {
  /** demand username/password auth (RFC 1929) */
  requireAuth?: { user: string; pass: string }
  /** reply code to refuse with instead of tunnelling */
  refuseWith?: number
}

/** A minimal RFC 1928 server: enough to exercise the bridge's client half. */
async function startSocks(options: SocksOptions = {}): Promise<number> {
  const server = createTcpServer((socket) => {
    let stage: 'greeting' | 'auth' | 'request' | 'tunnel' = 'greeting'
    let upstream: Socket | null = null
    socket.on('data', (chunk: Buffer) => {
      if (stage === 'greeting') {
        const wantsAuth = options.requireAuth !== undefined
        socket.write(Buffer.from([5, wantsAuth ? 2 : 0]))
        stage = wantsAuth ? 'auth' : 'request'
        return
      }
      if (stage === 'auth') {
        const userLen = chunk[1]!
        const user = chunk.subarray(2, 2 + userLen).toString('utf8')
        const passLen = chunk[2 + userLen]!
        const pass = chunk.subarray(3 + userLen, 3 + userLen + passLen).toString('utf8')
        const ok = user === options.requireAuth?.user && pass === options.requireAuth?.pass
        socket.write(Buffer.from([1, ok ? 0 : 1]))
        if (!ok) {
          socket.end()
          return
        }
        stage = 'request'
        return
      }
      if (stage === 'request') {
        // 05 01 00 03 <len> <host> <port>
        const nameLen = chunk[4]!
        const host = chunk.subarray(5, 5 + nameLen).toString('utf8')
        const port = chunk.readUInt16BE(5 + nameLen)
        if (options.refuseWith !== undefined) {
          socket.write(Buffer.from([5, options.refuseWith, 0, 1, 0, 0, 0, 0, 0, 0]))
          socket.end()
          return
        }
        upstream = connect({ host, port }, () => {
          socket.write(Buffer.from([5, 0, 0, 1, 127, 0, 0, 1, 0, 0]))
          stage = 'tunnel'
          upstream!.pipe(socket)
        })
        upstream.on('error', () => socket.destroy())
        return
      }
      upstream?.write(chunk)
    })
    socket.on('error', () => undefined)
  })
  closers.push(() => server.close())
  return listen(server)
}

/** Issue a CONNECT through the bridge and return the status line plus the first payload. */
function connectThrough(
  bridge: string,
  target: string,
): Promise<{ status: string; payload: string }> {
  const { port } = new URL(bridge)
  return new Promise((resolve, reject) => {
    const socket = connect({ host: '127.0.0.1', port: Number(port) }, () => {
      socket.write(`CONNECT ${target} HTTP/1.1\r\nHost: ${target}\r\n\r\n`)
    })
    let buffer = ''
    socket.setTimeout(4000, () => {
      socket.destroy()
      reject(new Error('timed out'))
    })
    socket.on('data', (chunk) => {
      buffer += chunk.toString('utf8')
      const split = buffer.indexOf('\r\n\r\n')
      if (split === -1) return
      const status = buffer.slice(0, buffer.indexOf('\r\n'))
      const payload = buffer.slice(split + 4)
      if (!status.includes('200') || payload.length > 0) {
        socket.destroy()
        resolve({ status, payload })
      }
    })
    socket.on('error', reject)
  })
}

describe('isBridgeableSocksUrl', () => {
  it('accepts socks5, which is the only scheme the bridge speaks', () => {
    expect(isBridgeableSocksUrl('socks5://127.0.0.1:7897')).toBe(true)
    expect(isBridgeableSocksUrl('socks://127.0.0.1:7897')).toBe(true)
  })

  it('rejects everything else, so callers fall back rather than bridge blindly', () => {
    expect(isBridgeableSocksUrl('http://127.0.0.1:7890')).toBe(false)
    // socks4 has no domain-name mode worth supporting
    expect(isBridgeableSocksUrl('socks4://127.0.0.1:1080')).toBe(false)
    expect(isBridgeableSocksUrl('')).toBe(false)
    expect(isBridgeableSocksUrl('not a url')).toBe(false)
  })
})

describe('splitHostPort', () => {
  it('splits the forms CONNECT actually sends', () => {
    expect(splitHostPort('www.genspark.ai:443')).toEqual(['www.genspark.ai', 443])
    expect(splitHostPort('[::1]:8443')).toEqual(['::1', 8443])
  })

  it('rejects targets with no usable port', () => {
    expect(splitHostPort('www.genspark.ai')).toEqual(['', 0])
    expect(splitHostPort('host:0')).toEqual(['', 0])
    expect(splitHostPort('host:99999')).toEqual(['', 0])
  })
})

describe('startSocksBridge', () => {
  it('tunnels a CONNECT through the SOCKS proxy to the origin', async () => {
    const origin = await startOrigin()
    const socks = await startSocks()
    const bridge = await startSocksBridge(`socks5://127.0.0.1:${socks}`)

    expect(bridge).toMatch(/^http:\/\/127\.0\.0\.1:\d+$/)
    const { status, payload } = await connectThrough(bridge, `localhost:${origin}`)
    expect(status).toContain('200')
    expect(payload).toBe('hello\n')
  })

  it('authenticates when the url carries credentials', async () => {
    const origin = await startOrigin()
    const socks = await startSocks({ requireAuth: { user: 'me', pass: 'pw' } })
    const bridge = await startSocksBridge(`socks5://me:pw@127.0.0.1:${socks}`)

    const { status } = await connectThrough(bridge, `localhost:${origin}`)
    expect(status).toContain('200')
  })

  it('answers 502 when the proxy refuses, rather than hanging the child', async () => {
    const origin = await startOrigin()
    // 5 = connection refused
    const socks = await startSocks({ refuseWith: 5 })
    const bridge = await startSocksBridge(`socks5://127.0.0.1:${socks}`)

    const { status } = await connectThrough(bridge, `localhost:${origin}`)
    expect(status).toContain('502')
  })

  it('reuses the same port for the same url and rebinds for a different one', async () => {
    const socks = await startSocks()
    const first = await startSocksBridge(`socks5://127.0.0.1:${socks}`)
    expect(await startSocksBridge(`socks5://127.0.0.1:${socks}`)).toBe(first)

    const other = await startSocks()
    const second = await startSocksBridge(`socks5://127.0.0.1:${other}`)
    expect(second).not.toBe(first)
  })

  it('returns nothing for a url it cannot bridge', async () => {
    expect(await startSocksBridge('http://127.0.0.1:7890')).toBe('')
    expect(await startSocksBridge('')).toBe('')
  })
})
