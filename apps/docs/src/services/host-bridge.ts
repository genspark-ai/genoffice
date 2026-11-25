import { randomBytes, randomUUID, timingSafeEqual } from 'node:crypto'
import { createServer, type IncomingMessage, type ServerResponse } from 'node:http'
import { z } from 'zod'
import type { DocsCommandEndpoint } from './docs-mcp'
import { DocsOperationError, isDocsOperationResult } from '../shared/operation-result'
import { DOCS_COMMAND_MAX_BYTES } from '../shared/media-policy'

export interface DocsHostBridge {
  readonly url: string
  /** Pass through the child environment; never include in tool results or activity logs. */
  readonly token: string
  close(): Promise<void>
}

const requestSchema = z.object({
  jsonrpc: z.literal('2.0'),
  method: z.string(),
  id: z.union([z.string(), z.number(), z.null()]).optional(),
  params: z.unknown().optional(),
})
const bridgeCommand = z.enum([
  'list_tools',
  'execute_tool',
  'read_document',
  'save_document',
  'close_caller',
])
const paramsSchema = z
  .object({
    message: z
      .object({
        command: bridgeCommand,
        payload: z.unknown().optional(),
      })
      .strict(),
  })
  .strict()

async function readBody(request: IncomingMessage): Promise<unknown> {
  const chunks: Buffer[] = []
  let size = 0
  for await (const chunk of request) {
    const bytes = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)
    size += bytes.length
    if (size > DOCS_COMMAND_MAX_BYTES) throw new RangeError('Host request exceeds the size limit')
    chunks.push(bytes)
  }
  return JSON.parse(Buffer.concat(chunks).toString('utf8')) as unknown
}

const rpcError = (id: string | number | null, code: number, message: string, data?: unknown) => ({
  jsonrpc: '2.0',
  id,
  error: { code, message, ...(data === undefined ? {} : { data }) },
})

/** Authenticated, local JSON-RPC host IPC. Public Agent calls remain standard stdio MCP. */
export async function createDocsHostBridge(options: {
  endpoint: DocsCommandEndpoint
  bindingId: string
  releaseScope(bindingId: string): void
}): Promise<DocsHostBridge> {
  const token = randomBytes(32).toString('hex')
  const expected = Buffer.from(`Bearer ${token}`)
  const callerId = randomUUID()
  const active = new Set<AbortController>()
  let origin = ''
  let revoked = false
  let closing: Promise<void> | undefined
  const revoke = () => {
    if (revoked) return
    revoked = true
    for (const controller of active) controller.abort()
    options.endpoint.closeCaller(callerId)
    options.releaseScope(options.bindingId)
  }
  const dispatch = async (value: unknown, signal: AbortSignal) => {
    const parsed = requestSchema.safeParse(value)
    if (!parsed.success) return rpcError(null, -32600, 'Invalid Request')
    const request = parsed.data
    const notification = request.id === undefined
    let reply: unknown
    if (request.method !== 'docs.request')
      reply = rpcError(request.id ?? null, -32601, 'Method not found')
    else {
      const params = paramsSchema.safeParse(request.params)
      if (!params.success) reply = rpcError(request.id ?? null, -32602, 'Invalid params')
      else if (revoked) reply = rpcError(request.id ?? null, -32001, 'The document grant is closed')
      else if (params.data.message.command === 'close_caller') {
        revoke()
        reply = { jsonrpc: '2.0', id: request.id ?? null, result: { closed: true } }
      } else {
        try {
          const result = await options.endpoint.request(
            {
              requestId: randomUUID(),
              callerId,
              bindingId: options.bindingId,
              command: params.data.message.command,
              payload: params.data.message.payload ?? {},
            },
            signal,
          )
          reply = { jsonrpc: '2.0', id: request.id ?? null, result: result ?? null }
        } catch (error) {
          reply = rpcError(
            request.id ?? null,
            -32000,
            error instanceof Error ? error.message : String(error),
            error instanceof DocsOperationError ? error.result : undefined,
          )
        }
      }
    }
    return notification ? undefined : reply
  }
  const handle = async (request: IncomingMessage, response: ServerResponse) => {
    response.setHeader('cache-control', 'no-store')
    const send = (status: number, body?: unknown) => {
      if (response.destroyed) return
      response.statusCode = status
      response.setHeader('content-type', 'application/json')
      response.end(body === undefined ? undefined : JSON.stringify(body))
    }
    if (
      request.headers.host !== new URL(origin).host ||
      (request.headers.origin !== undefined && request.headers.origin !== origin)
    ) {
      send(403, { error: 'Invalid host or origin' })
      return
    }
    const supplied = Buffer.from(request.headers.authorization ?? '')
    if (supplied.length !== expected.length || !timingSafeEqual(supplied, expected)) {
      response.setHeader('www-authenticate', 'Bearer')
      send(401, { error: 'Unauthorized' })
      return
    }
    if (request.method !== 'POST' || request.url !== '/docs') {
      send(404, { error: 'Not found' })
      return
    }
    const controller = new AbortController()
    active.add(controller)
    const onClose = () => {
      if (!response.writableEnded) controller.abort()
    }
    response.on('close', onClose)
    try {
      let body: unknown
      try {
        body = await readBody(request)
      } catch (error) {
        send(
          error instanceof RangeError ? 413 : 400,
          rpcError(
            null,
            error instanceof RangeError ? -32602 : -32700,
            error instanceof RangeError ? 'Host request exceeds the size limit' : 'Parse error',
          ),
        )
        return
      }
      if (Array.isArray(body)) {
        if (!body.length) {
          send(200, rpcError(null, -32600, 'Invalid Request'))
          return
        }
        const replies = (
          await Promise.all(body.map((value) => dispatch(value, controller.signal)))
        ).filter((value) => value !== undefined)
        send(replies.length ? 200 : 204, replies.length ? replies : undefined)
      } else {
        const reply = await dispatch(body, controller.signal)
        send(reply === undefined ? 204 : 200, reply)
      }
    } finally {
      active.delete(controller)
      response.off('close', onClose)
    }
  }
  const server = createServer((request, response) => {
    void handle(request, response).catch(() => {
      if (!response.destroyed) {
        response.writeHead(500, { 'content-type': 'application/json' })
        response.end(JSON.stringify(rpcError(null, -32603, 'Internal error')))
      }
    })
  })
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject)
    server.listen(0, '127.0.0.1', () => {
      server.off('error', reject)
      resolve()
    })
  })
  const address = server.address()
  if (!address || typeof address === 'string')
    throw new Error('The host bridge did not obtain a local port')
  origin = `http://127.0.0.1:${address.port}`
  return {
    url: `${origin}/docs`,
    token,
    close() {
      if (!closing) {
        revoke()
        closing = new Promise<void>((resolve, reject) => {
          server.close((error) => (error ? reject(error) : resolve()))
          server.closeAllConnections()
        })
      }
      return closing
    },
  }
}

const responseSchema = z.object({
  jsonrpc: z.literal('2.0'),
  id: z.string(),
  result: z.unknown().optional(),
  error: z
    .object({ code: z.number(), message: z.string(), data: z.unknown().optional() })
    .optional(),
})

export function connectDocsHostBridge(
  connection: Pick<DocsHostBridge, 'url' | 'token'>,
): DocsCommandEndpoint {
  const target = new URL(connection.url)
  if (
    target.protocol !== 'http:' ||
    target.hostname !== '127.0.0.1' ||
    target.pathname !== '/docs' ||
    target.username ||
    target.password ||
    target.search ||
    target.hash
  )
    throw new Error('The editor host endpoint must be an explicit local bridge')
  let closed = false
  const post = async (
    command: z.infer<typeof bridgeCommand>,
    payload: unknown,
    signal?: AbortSignal,
  ) => {
    const id = randomUUID()
    const response = await fetch(target, {
      method: 'POST',
      signal,
      redirect: 'error',
      headers: { 'content-type': 'application/json', authorization: `Bearer ${connection.token}` },
      body: JSON.stringify({
        jsonrpc: '2.0',
        id,
        method: 'docs.request',
        params: { message: { command, payload } },
      }),
    })
    if (!response.ok) throw new Error(`Editor host request failed (HTTP ${response.status})`)
    const result = responseSchema.parse(await response.json())
    if (result.id !== id || 'error' in result === 'result' in result)
      throw new Error('The editor host returned an invalid response')
    if (result.error) {
      if (isDocsOperationResult(result.error.data) && result.error.data.isError)
        throw new DocsOperationError(result.error.message, result.error.data)
      throw new Error(result.error.message)
    }
    return result.result
  }
  return {
    request(message, signal) {
      if (closed) return Promise.reject(new Error('The document connection is closed'))
      if (
        message.command !== 'list_tools' &&
        message.command !== 'execute_tool' &&
        message.command !== 'read_document' &&
        message.command !== 'save_document'
      )
        return Promise.reject(
          new Error('The command is not available over the document connection'),
        )
      return post(message.command, message.payload, signal)
    },
    closeCaller() {
      if (closed) return
      closed = true
      void post('close_caller', {}).catch(() => {
        /* The host may already have revoked this grant. */
      })
    },
  }
}
