import { afterEach, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { createDocsHttpMcpServer, type DocsHttpMcpServer } from '../src/services/docs-http-mcp'
import { createDocsRuntime, type DocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { createDocsCommandRouter } from '../src/renderer/runtime/command-router'

let editor: Editor | undefined
let runtime: DocsRuntime | undefined
let service: DocsHttpMcpServer | undefined
const clients: Client[] = []

afterEach(async () => {
  await service?.close()
  await Promise.all(clients.splice(0).map((client) => client.close()))
  runtime?.dispose()
  editor?.destroy()
  service = undefined
})

async function fixture() {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<p>Unsaved words</p>',
  })
  const active = editor
  runtime = createDocsRuntime({
    current: () => ({ editor: active, identity: active, numIds: { bullet: null, ordered: null } }),
  })
  const scope = runtime.capture()
  const router = createDocsCommandRouter(runtime)
  router.bindScope('document-grant', scope)
  service = await createDocsHttpMcpServer({
    version: 'test',
    documentId: scope.documentId,
    bindingId: 'document-grant',
    releaseScope: (id) => router.releaseScope(id),
    endpoint: {
      request: (message, signal) => router.handle(message, signal),
      closeCaller: (id) => router.disconnect(id),
    },
  })
  return { active, service, router }
}

async function connect(server: DocsHttpMcpServer) {
  const client = new Client({ name: 'sdk-http-test', version: '1' })
  clients.push(client)
  const transport = new StreamableHTTPClientTransport(new URL(server.connection.url), {
    requestInit: { headers: server.connection.headers },
  })
  await client.connect(transport)
  return { client, transport }
}

it('exposes standard HTTP MCP tools over the live unsaved document using returned connection information', async () => {
  const { active, service } = await fixture()
  const { client } = await connect(service)
  expect(service.connection.transport).toBe('streamable-http')
  const { tools } = await client.listTools()
  const names = tools.map((tool) => tool.name)
  expect(names).toContain('read_document')
  expect(names).not.toContain('read_attachment')
  expect(names).not.toContain('get_editor_capabilities')
  expect(JSON.stringify(tools.find((tool) => tool.name === 'apply_ops')?.inputSchema)).toContain(
    'setFont',
  )
  expect(await client.callTool({ name: 'get_editor_capabilities', arguments: {} })).toMatchObject({
    isError: true,
    content: [{ text: expect.stringContaining('not found') }],
  })
  expect(JSON.stringify(await client.callTool({ name: 'read_document' }))).toContain(
    'Unsaved words',
  )
  const result = await client.callTool({
    name: 'insert_content',
    arguments: { html: '<p>Agent addition</p>' },
  })
  expect(result.isError, JSON.stringify(result)).not.toBe(true)
  expect(active.getText()).toContain('Agent addition')
  active.commands.insertContent('<p>Manual addition</p>')
  expect(JSON.stringify(await client.callTool({ name: 'read_document' }))).toContain(
    'Manual addition',
  )
})

it('authenticates HTTP access and keeps another caller usable after a client terminates its session', async () => {
  const { service, active } = await fixture()
  const first = await connect(service)
  const second = await connect(service)
  const denied = await fetch(service.connection.url, { method: 'POST', body: '{}' })
  expect(denied.status).toBe(401)
  expect(await denied.text()).not.toContain(service.connection.headers.Authorization)
  const origin = await fetch(service.connection.url, {
    method: 'POST',
    headers: { ...service.connection.headers, Origin: 'https://untrusted.test' },
    body: '{}',
  })
  expect(origin.status).toBe(403)
  await first.transport.terminateSession()
  await first.client.close()
  expect(JSON.stringify(await second.client.callTool({ name: 'read_document' }))).toContain(
    'Unsaved words',
  )
  const result = await second.client.callTool({
    name: 'insert_content',
    arguments: { html: '<p>Second caller</p>' },
  })
  expect(result.isError).not.toBe(true)
  expect(active.getText()).toContain('Second caller')
  await service.close()
  await service.close()
  await expect(
    fetch(service.connection.url, { headers: service.connection.headers }),
  ).rejects.toThrow()
})

it('releases an unused document grant when its host ends the document generation', async () => {
  const { router, service } = await fixture()
  // Document replacement releases scopes even before a first client connects.
  const released = vi.fn(() => {
    void service.close()
  })
  const scope = runtime!.capture()
  router.releaseScope('document-grant')
  router.bindScope('document-grant', scope, released)
  router.releaseAll()
  expect(released).toHaveBeenCalledOnce()
  await service.close()
  await expect(connect(service)).rejects.toThrow()
})

it('does not retain sessions after rejected HTTP initialization negotiation', async () => {
  const { service } = await fixture()
  for (let index = 0; index < 35; index++) {
    const response = await fetch(service.connection.url, {
      method: 'POST',
      headers: {
        ...service.connection.headers,
        'content-type': index % 2 ? 'text/plain' : 'application/json',
        ...(index % 2 ? { Accept: 'application/json, text/event-stream' } : {}),
      },
      body: JSON.stringify({
        jsonrpc: '2.0',
        id: index,
        method: 'initialize',
        params: {
          protocolVersion: '2025-11-25',
          capabilities: {},
          clientInfo: { name: 'rejected', version: '1' },
        },
      }),
    })
    expect(response.status).toBe(index % 2 ? 415 : 406)
    await response.text()
  }
  const { client } = await connect(service)
  expect(JSON.stringify(await client.callTool({ name: 'read_document' }))).toContain(
    'Unsaved words',
  )
})

it('closes a pending tool caller and tolerates reentrant grant cleanup', async () => {
  let entered!: () => void
  const started = new Promise<void>((resolve) => {
    entered = resolve
  })
  let operationSignal: AbortSignal | undefined
  service = await createDocsHttpMcpServer({
    version: 'test',
    bindingId: 'pending-grant',
    documentId: 'pending-document',
    releaseScope: () => {
      void service!.close()
    },
    endpoint: {
      request: async (_message, signal) => {
        operationSignal = signal
        entered()
        return new Promise((_resolve, reject) => {
          signal.addEventListener('abort', () => reject(new Error('Document access ended')), {
            once: true,
          })
        })
      },
      closeCaller() {},
    },
  })
  const { client } = await connect(service)
  const pending = client.callTool({ name: 'read_document' }).catch(() => undefined)
  await started
  await service.close()
  expect(operationSignal!.aborted).toBe(true)
  await pending
})
