import { afterEach, describe, expect, it } from 'vitest'
import { z } from 'zod'
import { request as httpRequest } from 'node:http'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { SSEClientTransport } from '@modelcontextprotocol/sdk/client/sse.js'
import { McpServerService, type McpToolDefinition } from '../../src/main/mcp/mcp-server'

/**
 * M2: transport-level behaviour of the local MCP server.
 *
 * Uses the SDK's own client against the service in-process, so the assertions
 * cover the real handshake rather than a mock.
 */

const TOOLS: McpToolDefinition[] = [
  {
    name: 'echo',
    description: 'echo the given text back',
    inputSchema: { text: z.string() },
    handler: (args) => ({ echoed: String(args.text) }),
  },
  {
    name: 'boom',
    description: 'always throws',
    handler: () => {
      throw new Error('tool exploded')
    },
  },
]

let service: McpServerService | undefined

afterEach(async () => {
  await service?.stop()
  service = undefined
})

/** find a free port by binding an ephemeral one, then releasing it */
async function freePort(): Promise<number> {
  const { createServer } = await import('node:http')
  return new Promise((resolve, reject) => {
    const server = createServer()
    server.on('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const address = server.address()
      const port = typeof address === 'object' && address ? address.port : 0
      server.close(() => resolve(port))
    })
  })
}

async function startService(): Promise<number> {
  const port = await freePort()
  service = new McpServerService({ port, tools: TOOLS })
  await service.start()
  return port
}

describe('MCP server core', () => {
  it('answers /health on loopback', async () => {
    const port = await startService()
    const res = await fetch(`http://127.0.0.1:${port}/health`)
    expect(res.status).toBe(200)
    const body = (await res.json()) as { status: string; server: string }
    expect(body.status).toBe('ok')
    expect(body.server).toBe('GenOffice')
  })

  it('completes initialize and lists the registered tools', async () => {
    const port = await startService()
    const client = new Client({ name: 'test-client', version: '0.0.0' })
    const transport = new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`))
    await client.connect(transport)

    const { tools } = await client.listTools()
    expect(tools.map((t) => t.name).sort()).toEqual(['boom', 'echo'])

    await client.close()
  })

  it('runs a tool and returns its JSON result', async () => {
    const port = await startService()
    const client = new Client({ name: 'test-client', version: '0.0.0' })
    await client.connect(new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`)))

    const result = await client.callTool({ name: 'echo', arguments: { text: 'hi' } })
    const content = result.content as Array<{ type: string; text?: string }>
    expect(content[0].type).toBe('text')
    expect(content[0].text).toContain('hi')

    await client.close()
  })

  it('reports a throwing tool as isError without killing the session', async () => {
    const port = await startService()
    const client = new Client({ name: 'test-client', version: '0.0.0' })
    await client.connect(new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`)))

    const result = await client.callTool({ name: 'boom', arguments: {} })
    expect(result.isError).toBe(true)
    const content = result.content as Array<{ type: string; text?: string }>
    expect(content[0].text).toContain('tool exploded')

    // the session still works
    const ok = await client.callTool({ name: 'echo', arguments: { text: 'still alive' } })
    expect(ok.isError).toBeFalsy()

    await client.close()
  })

  it('rejects a foreign Host header', async () => {
    const port = await startService()
    // fetch/undici refuses to override Host, so go through node:http directly
    const status = await new Promise<number>((resolve, reject) => {
      const req = httpRequest(
        {
          host: '127.0.0.1',
          port,
          path: '/health',
          method: 'GET',
          headers: { host: `evil.example.com:${port}` },
        },
        (res) => {
          res.resume()
          res.on('end', () => resolve(res.statusCode ?? 0))
        },
      )
      req.on('error', reject)
      req.end()
    })
    expect(status).toBe(403)
  })

  it('rejects a foreign Origin header', async () => {
    const port = await startService()
    const res = await fetch(`http://127.0.0.1:${port}/health`, {
      headers: { origin: 'http://evil.example.com' },
    })
    expect(res.status).toBe(403)
  })

  it('serves the legacy SSE transport (GET /sse + POST /messages)', async () => {
    const port = await startService()
    const client = new Client({ name: 'sse-test', version: '0.0.0' })
    const transport = new SSEClientTransport(new URL(`http://127.0.0.1:${port}/sse`))
    await client.connect(transport)

    const { tools } = await client.listTools()
    expect(tools.map((t) => t.name).sort()).toEqual(['boom', 'echo'])

    const result = await client.callTool({ name: 'echo', arguments: { text: 'over sse' } })
    const content = result.content as Array<{ type: string; text?: string }>
    expect(content[0].text).toContain('over sse')

    await client.close()
  })

  it('bind failure on an occupied port surfaces as EADDRINUSE after retries', async () => {
    const port = await freePort()
    service = new McpServerService({ port, tools: TOOLS })
    await service.start()

    const second = new McpServerService({
      port,
      tools: TOOLS,
      maxListenAttempts: 3,
      listenRetryDelayMs: 10,
    })
    await expect(second.start()).rejects.toMatchObject({ code: 'EADDRINUSE' })
  })

  it('toolsFactory runs once per client session, giving each its own tool state', async () => {
    const port = await freePort()
    let builds = 0
    // a stateful tool: each build gets a fresh counter closure
    service = new McpServerService({
      port,
      toolsFactory: () => {
        builds++
        let count = 0
        return [
          {
            name: 'count',
            description: 'increments a session-local counter',
            handler: () => ({ count: ++count }),
          },
        ]
      },
    })
    await service.start()

    const connect = async (): Promise<Client> => {
      const c = new Client({ name: 'factory-test', version: '0.0.0' })
      await c.connect(new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`)))
      return c
    }
    const text = (r: { content: unknown }): string =>
      (r.content as Array<{ text?: string }>).map((c) => c.text ?? '').join('')

    const a = await connect()
    const b = await connect()
    expect(builds).toBe(2) // one factory call per session

    // each client's counter is independent: two calls on A reach 2, B starts at 1
    await a.callTool({ name: 'count', arguments: {} })
    const aSecond = await a.callTool({ name: 'count', arguments: {} })
    expect(text(aSecond)).toContain('2')
    const bFirst = await b.callTool({ name: 'count', arguments: {} })
    expect(text(bFirst)).toContain('1')

    await a.close()
    await b.close()
  })
})
