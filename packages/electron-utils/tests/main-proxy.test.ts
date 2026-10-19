import { createServer, type RequestListener } from 'node:http'
import type { AddressInfo } from 'node:net'
import { describe, expect, it } from 'vitest'
import { createMainProcessProxyDispatcher, mergeNoProxy } from '../src/main-proxy'

describe('mergeNoProxy', () => {
  it('keeps the caller entries in front of the loopback defaults', () => {
    expect(mergeNoProxy('internal.example, *.corp')).toBe(
      'internal.example, *.corp,localhost,127.0.0.1,[::1]',
    )
  })

  it('ships the loopback defaults alone when the user set nothing', () => {
    expect(mergeNoProxy(undefined)).toBe('localhost,127.0.0.1,[::1]')
  })
})

describe('createMainProcessProxyDispatcher', () => {
  /** a loopback http server; `label` is what it answers with */
  const httpServer = async (label: string): Promise<{ url: string; close: () => void }> => {
    const listener: RequestListener = (_req, res) => {
      res.end(label)
    }
    const server = createServer(listener)
    await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
    const { port } = server.address() as AddressInfo
    return { url: `http://127.0.0.1:${port}`, close: () => server.close() }
  }

  it('sends loopback traffic directly even though a proxy is configured', async () => {
    const origin = await httpServer('loopback direct')
    // port 9 (discard) is not listening: a request that reached the "proxy"
    // would fail with ECONNREFUSED instead of this assertion passing
    const dispatcher = await createMainProcessProxyDispatcher('http://127.0.0.1:9')
    const { statusCode, body } = await dispatcher.request({
      origin: origin.url,
      path: '/',
      method: 'GET',
    })
    expect(statusCode).toBe(200)
    expect(await body.text()).toBe('loopback direct')
    await dispatcher.close()
    origin.close()
  })

  it('still sends everything off loopback to the proxy', async () => {
    // nothing listens on port 9, so the dispatch failing fast there is the
    // proxy being consulted at all — the loopback target above never touches it
    const dispatcher = await createMainProcessProxyDispatcher('http://127.0.0.1:9')
    await expect(
      dispatcher.request({ origin: 'http://example.com/', path: '/', method: 'GET' }),
    ).rejects.toThrow()
    await dispatcher.close()
  })

  it('honours the caller NO_PROXY entries the undici way', async () => {
    const origin = await httpServer('exempt host')
    process.env.no_proxy = '127.0.0.1'
    try {
      const dispatcher = await createMainProcessProxyDispatcher('http://127.0.0.1:9')
      const { statusCode } = await dispatcher.request({
        origin: origin.url,
        path: '/',
        method: 'GET',
      })
      expect(statusCode).toBe(200)
      await dispatcher.close()
    } finally {
      delete process.env.no_proxy
    }
    origin.close()
  })
})
