import { afterEach, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import {
  CallToolResultSchema,
  ToolListChangedNotificationSchema,
} from '@modelcontextprotocol/sdk/types.js'
import { createServer } from 'node:net'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { McpServerService, McpToolExecutionError } from '../../shell/src/main/mcp/mcp-server'
import { DocsOperationError } from '../src/shared/operation-result'
import {
  createDocumentTools,
  documentDriver,
  type DocsControl,
} from '../../shell/src/main/mcp/tools/document-tools'
import { createSessionHost, createSessionTools } from '../../shell/src/main/mcp/tools/session-tools'
import { createDocsRuntime, type DocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { createDocsCommandRouter } from '../src/renderer/runtime/command-router'
import type { McpEditorCommand } from '../src/shared/ipc'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { createDocsMcpServer } from '../src/services/docs-mcp'
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'
import { resolve } from 'node:path'
import { createDocsHostBridge, type DocsHostBridge } from '../src/services/host-bridge'
import { createKnowledgeBaseSkill } from '@genoffice/agent-core'
import { createDocsResourceTools } from '../src/renderer/runtime/resource-tools'

let editor: Editor | undefined
let runtime: DocsRuntime | undefined
let service: McpServerService | undefined
let client: Client | undefined
let boundServer: McpServer | undefined
let hostBridge: DocsHostBridge | undefined
afterEach(async () => {
  await client?.close()
  await service?.stop()
  await boundServer?.close()
  await hostBridge?.close()
  runtime?.dispose()
  editor?.destroy()
})

async function expectStandardToolDiscovery(mcp: Client) {
  const { tools } = await mcp.listTools()
  const names = tools.map((tool) => tool.name)
  expect(names).toEqual(expect.arrayContaining(['query_document', 'read_document', 'apply_ops']))
  expect(names).not.toContain('get_editor_capabilities')
  expect(names).not.toContain('read_attachment')
  const apply = tools.find((tool) => tool.name === 'apply_ops')
  expect(apply?.description).toContain('setFont')
  expect(apply?.inputSchema.properties).toMatchObject({
    ops: {
      type: 'array',
      items: {
        oneOf: expect.arrayContaining([
          expect.objectContaining({
            properties: expect.objectContaining({
              op: { type: 'string', const: 'setFont' },
              target: {
                type: 'object',
                properties: expect.objectContaining({ nodeType: expect.anything() }),
                required: [],
                minProperties: 1,
                description: expect.any(String),
              },
              color: { anyOf: [{ type: 'string' }, { type: 'null' }] },
            }),
            required: ['op', 'target'],
          }),
        ]),
      },
    },
    dryRun: { type: 'boolean', description: expect.any(String) },
  })
  expect(await mcp.callTool({ name: 'get_editor_capabilities', arguments: {} })).toMatchObject({
    isError: true,
    content: [{ text: expect.stringContaining('not found') }],
  })
}

it('keeps optional helper contracts callable as host resources appear and disappear without reconnecting', async () => {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<p>Live text</p>',
  })
  const active = editor
  let sources: ReturnType<typeof createKnowledgeBaseSkill> | undefined
  runtime = createDocsRuntime({
    current: () => ({ editor: active, identity: active, numIds: { bullet: null, ordered: null } }),
    resources: () => sources,
  })
  const router = createDocsCommandRouter(runtime)
  boundServer = createDocsMcpServer({
    version: 'test',
    endpoint: {
      request: (message, signal) => router.handle(message, signal),
      closeCaller: (callerId) => router.disconnect(callerId),
    },
  })
  const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair()
  await boundServer.connect(serverTransport)
  client = new Client({ name: 'dynamic-resources', version: '1' }, { capabilities: {} })
  await client.connect(clientTransport)
  const notifications = vi.fn()
  client.setNotificationHandler(ToolListChangedNotificationSchema, notifications)
  const names = async () => (await client!.listTools()).tools.map((tool) => tool.name)
  expect(await names()).not.toContain('search_knowledge_base')
  expect(await names()).not.toContain('read_attachment')
  sources = createKnowledgeBaseSkill({
    listFiles: () => [
      {
        name: 'later.txt',
        folder: 'granted',
        path: '/granted/later.txt',
        ext: 'txt',
        sizeBytes: 8,
      },
    ],
    search: async () => [
      {
        name: 'later.txt',
        folder: 'granted',
        path: '/granted/later.txt',
        ext: 'txt',
        snippet: 'Late source evidence',
      },
    ],
    read: async () => ({ text: 'Late source evidence', offset: 0, totalChars: 20 }),
  })
  const searched = await client.callTool({
    name: 'search_knowledge_base',
    arguments: { query: 'evidence' },
  })
  expect(searched.isError, JSON.stringify(searched)).not.toBe(true)
  expect(JSON.stringify(searched.content)).toContain('Late source evidence')
  await vi.waitFor(() => expect(notifications).toHaveBeenCalledTimes(1), { timeout: 3500 })
  expect(await names()).toContain('search_knowledge_base')
  expect(await names()).toContain('read_knowledge_file')
  sources = undefined
  await vi.waitFor(() => expect(notifications).toHaveBeenCalledTimes(2), { timeout: 3500 })
  expect(await names()).not.toContain('read_knowledge_file')
  expect(
    (await client.callTool({ name: 'get_document_context', arguments: {} })).structuredContent,
  ).not.toHaveProperty('resourceContext')
  const unavailable = await client.callTool({
    name: 'read_knowledge_file',
    arguments: { index: 0 },
  })
  expect(unavailable).toMatchObject({
    isError: true,
    structuredContent: { committed: false, errorCode: 'unavailable' },
  })
  router.dispose()
})

it('discovers and calls the same authorized attachment and knowledge helpers through ordinary MCP tools', async () => {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<p>Live text</p>',
  })
  const active = editor
  const resources = createDocsResourceTools({
    attachments: {
      list: () => [
        { name: 'attached.txt', path: '/granted/attached.txt', ext: 'txt', sizeBytes: 8 },
      ],
      read: async (path) => {
        expect(path).toBe('/granted/attached.txt')
        return { text: 'Attached evidence', offset: 0, totalChars: 17 }
      },
    },
    knowledge: {
      listFiles: () => [
        {
          name: 'source.txt',
          folder: 'granted',
          path: '/granted/source.txt',
          ext: 'txt',
          sizeBytes: 6,
        },
      ],
      search: async () => [
        {
          name: 'source.txt',
          folder: 'granted',
          path: '/granted/source.txt',
          ext: 'txt',
          snippet: 'Source evidence',
        },
      ],
      read: async (path) => {
        expect(path).toBe('/granted/source.txt')
        return { text: 'Source evidence', offset: 0, totalChars: 15 }
      },
    },
    runCli: async (args, signal) => {
      expect(args).toEqual(['info', '/granted/source.docx'])
      expect(signal?.aborted).toBe(false)
      return { ok: true, code: 0, stdout: 'Office file metadata', stderr: '', committed: false }
    },
  })
  runtime = createDocsRuntime({
    current: () => ({ editor: active, identity: active, numIds: { bullet: null, ordered: null } }),
    resources: () => resources,
  })
  const native = runtime.connect({ id: 'native-sources', resources: () => resources })
  native.context()
  const router = createDocsCommandRouter(runtime)
  boundServer = createDocsMcpServer({
    version: 'test',
    endpoint: {
      request: (message, signal) => router.handle(message, signal),
      closeCaller: (callerId) => router.disconnect(callerId),
    },
  })
  const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair()
  await boundServer.connect(serverTransport)
  client = new Client({ name: 'resource-harness', version: '1' }, { capabilities: {} })
  await client.connect(clientTransport)
  const names = (await client.listTools()).tools.map((tool) => tool.name)
  expect(names).toEqual(
    expect.arrayContaining(['read_attachment', 'search_knowledge_base', 'read_knowledge_file']),
  )
  for (const [name, input] of [
    ['read_attachment', { index: 0 }],
    ['search_knowledge_base', { query: 'evidence' }],
    ['read_knowledge_file', { index: 0 }],
    ['run_cli', { args: ['info', '/granted/source.docx'] }],
  ] as const) {
    const local = await native.execute({ id: name, name, input })
    const external = await client.callTool({ name, arguments: input })
    expect(external.isError).not.toBe(true)
    expect(external.structuredContent).toMatchObject({
      output: local.output,
      committed: false,
      revision: 0,
    })
  }
  const context = await client.callTool({ name: 'get_document_context', arguments: {} })
  expect(context.structuredContent).toMatchObject({
    editable: true,
    output: expect.stringContaining('Live text'),
    resourceContext: expect.stringContaining('attached.txt'),
  })
  expect(context.structuredContent).not.toHaveProperty('tools')
  expect(context.structuredContent).not.toHaveProperty('harnessWorkflows')
  expect(names).toContain('start_document_draft')
  expect(names).not.toContain('write_document')
  expect(active.getText()).toBe('Live text')
  router.dispose()
})

it('edits the captured live document from the packaged stdio MCP process', async () => {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<p>One Two</p>',
  })
  const active = editor
  runtime = createDocsRuntime({
    current: () => ({ editor: active, identity: active, numIds: { bullet: null, ordered: null } }),
  })
  active.commands.setTextSelection({ from: 1, to: 4 })
  const router = createDocsCommandRouter(runtime)
  router.bindScope('host-selected-document', runtime.capture())
  active.commands.setTextSelection({ from: 5, to: 8 })
  hostBridge = await createDocsHostBridge({
    bindingId: 'host-selected-document',
    releaseScope: (binding) => router.releaseScope(binding),
    endpoint: {
      request: (message, signal) => router.handle(message, signal),
      closeCaller: (callerId) => router.disconnect(callerId),
    },
  })
  client = new Client({ name: 'stdio-harness', version: '1' }, { capabilities: {} })
  await client.connect(
    new StdioClientTransport({
      command: process.execPath,
      args: [resolve(__dirname, '../out/services/mcp-stdio.cjs')],
      env: {
        ...Object.fromEntries(
          Object.entries(process.env).filter(
            (entry): entry is [string, string] => entry[1] !== undefined,
          ),
        ),
        FORAOFFICE_DOCS_ENDPOINT: hostBridge.url,
        FORAOFFICE_DOCS_TOKEN: hostBridge.token,
      },
    }),
  )
  await expectStandardToolDiscovery(client)
  const changed = await client.callTool({ name: 'replace_selection', arguments: { html: 'New' } })
  expect(changed.isError).not.toBe(true)
  expect(active.getText()).toBe('New Two')
  router.dispose()
})

it('exposes a host-bound document over a standard SDK transport with no Electron document identifiers', async () => {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<p>One Two</p>',
  })
  const active = editor
  runtime = createDocsRuntime({
    current: () => ({ editor: active, identity: active, numIds: { bullet: null, ordered: null } }),
  })
  active.commands.setTextSelection({ from: 1, to: 4 })
  const scope = runtime.capture()
  active.commands.setTextSelection({ from: 5, to: 8 })
  const router = createDocsCommandRouter(runtime)
  router.bindScope('trusted-action-binding', scope)
  boundServer = createDocsMcpServer({
    version: 'test',
    bindingId: 'trusted-action-binding',
    endpoint: {
      request: (message, signal) => router.handle(message, signal),
      closeCaller: (callerId) => router.disconnect(callerId),
    },
  })
  const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair()
  await boundServer.connect(serverTransport)
  client = new Client({ name: 'ordinary-harness', version: '1' }, { capabilities: {} })
  await client.connect(clientTransport)
  await expectStandardToolDiscovery(client)
  expect((await client.callTool({ name: 'read_document' })).isError).not.toBe(true)
  const result = await client.callTool({ name: 'replace_selection', arguments: { html: 'New' } })
  expect(result.structuredContent).toMatchObject({ committed: true })
  expect(active.getText()).toBe('New Two')
  router.dispose()
})

it('lets only the original write caller stop and discard its own draft when the editor becomes read-only', async () => {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<h1>Heading</h1><p>Original body</p>',
  })
  const active = editor
  runtime = createDocsRuntime({
    current: () => ({ editor: active, identity: active, numIds: { bullet: null, ordered: null } }),
  })
  const router = createDocsCommandRouter(runtime)
  router.bindScope('draft-owner-grant', runtime.capture())
  const connect = async (readOnly = false) => {
    const server = createDocsMcpServer({
      version: 'test',
      bindingId: 'draft-owner-grant',
      readOnly,
      endpoint: {
        request: (message, signal) => router.handle(message, signal),
        closeCaller: (callerId) => router.disconnect(callerId),
      },
    })
    const [upstream, downstream] = InMemoryTransport.createLinkedPair()
    const mcp = new Client({ name: 'draft-test', version: '1' })
    await server.connect(upstream)
    await mcp.connect(downstream)
    return {
      mcp,
      close: async () => {
        await mcp.close()
        await server.close()
      },
    }
  }
  const owner = await connect()
  const unrelated = await connect()
  const reference = await connect(true)
  try {
    expect((await owner.mcp.listTools()).tools.map(({ name }) => name)).toContain('apply_ops')
    const started = await owner.mcp.callTool({
      name: 'start_document_draft',
      arguments: { afterBlockIndex: 0 },
    })
    const { draftId } = JSON.parse(
      String(CallToolResultSchema.parse(started).structuredContent?.output),
    )
    const updated = await owner.mcp.callTool({
      name: 'update_document_draft',
      arguments: { draftId, sequence: 1, html: '<p>Provisional text</p>' },
    })
    expect(updated.isError).not.toBe(true)
    active.setEditable(false)
    const names = (await owner.mcp.listTools()).tools.map(({ name }) => name)
    expect(names).toContain('stop_document_draft')
    expect(names).toContain('finish_document_draft')
    expect(names).not.toContain('start_document_draft')
    expect(names).not.toContain('apply_ops')
    expect((await unrelated.mcp.listTools()).tools.map(({ name }) => name)).not.toContain(
      'stop_document_draft',
    )
    const stoppedByAnother = await unrelated.mcp.callTool({
      name: 'stop_document_draft',
      arguments: { draftId },
    })
    expect(stoppedByAnother).toMatchObject({
      isError: true,
      structuredContent: { output: expect.stringContaining('does not belong') },
    })
    expect((await reference.mcp.listTools()).tools.map(({ name }) => name)).not.toContain(
      'finish_document_draft',
    )
    expect(
      await reference.mcp.callTool({
        name: 'finish_document_draft',
        arguments: { draftId, decision: 'discard' },
      }),
    ).toMatchObject({ isError: true, content: [{ text: expect.stringContaining('not found') }] })
    expect(
      await owner.mcp.callTool({
        name: 'apply_ops',
        arguments: {
          ops: [{ op: 'setFont', target: { nodeType: 'docHeading' }, color: '#FF0000' }],
        },
      }),
    ).toMatchObject({ isError: true, structuredContent: { errorCode: 'permission_denied' } })
    expect(
      (await owner.mcp.callTool({ name: 'stop_document_draft', arguments: { draftId } })).isError,
    ).not.toBe(true)
    expect(
      await owner.mcp.callTool({
        name: 'finish_document_draft',
        arguments: { draftId, decision: 'finalize' },
      }),
    ).toMatchObject({ isError: true, structuredContent: { errorCode: 'permission_denied' } })
    expect(
      (
        await owner.mcp.callTool({
          name: 'finish_document_draft',
          arguments: { draftId, decision: 'discard' },
        })
      ).isError,
    ).not.toBe(true)
    expect(active.getText()).toBe('Heading\n\nOriginal body')
    expect((await owner.mcp.listTools()).tools.map(({ name }) => name)).not.toContain(
      'stop_document_draft',
    )
  } finally {
    await Promise.all([owner.close(), unrelated.close(), reference.close()])
    router.dispose()
  }
})

it.each(['unknown', 'released', 'replaced'] as const)(
  'returns structured stale errors for a %s scope without inspecting the current document',
  async (authority) => {
    const { editorExtensions } = await import('../src/renderer/editor/extensions')
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: '<p>Original document</p>',
    })
    const active = editor
    let identity = {}
    let reads = 0
    let saves = 0
    runtime = createDocsRuntime({
      current: () => {
        reads++
        return { editor: active, identity, numIds: { bullet: null, ordered: null } }
      },
      save: async () => {
        saves++
        return { ok: true, committed: true, pending: false }
      },
    })
    const router = createDocsCommandRouter(runtime)
    if (authority !== 'unknown') router.bindScope('old-action', runtime.capture())
    boundServer = createDocsMcpServer({
      version: 'test',
      bindingId: 'old-action',
      endpoint: {
        request: (message, signal) => router.handle(message, signal),
        closeCaller: (callerId) => router.disconnect(callerId),
      },
    })
    const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair()
    await boundServer.connect(serverTransport)
    client = new Client({ name: 'stale-scope-harness', version: '1' }, { capabilities: {} })
    await client.connect(clientTransport)
    if (authority !== 'unknown') {
      expect((await client.callTool({ name: 'read_document' })).isError).not.toBe(true)
      if (authority === 'released') router.releaseScope('old-action')
      else router.releaseAll()
    }
    active.commands.setContent('<p>Private replacement</p>')
    identity = {}
    // The injected host authority must not be asked to expose the replacement document.
    const readsBefore = reads
    for (const [name, args] of [
      ['read_document', {}],
      ['replace_selection', { html: 'Unauthorized edit' }],
      ['save_document', { path: '/unauthorized.docx' }],
    ] as const) {
      const result = await client.callTool({ name, arguments: args })
      expect(result).toMatchObject({
        isError: true,
        structuredContent: {
          documentId: '',
          revision: 0,
          errorCode: 'stale_target',
          committed: false,
        },
      })
      expect(JSON.stringify(result)).not.toContain('Private replacement')
    }
    expect(reads).toBe(readsBefore)
    expect(saves).toBe(0)
    expect(active.getText()).toBe('Private replacement')
    router.dispose()
  },
)

it.each(['standalone', 'inline', 'table'])(
  'returns %s images through revision-bound references and standard MCP image content',
  async (kind) => {
    const { editorExtensions } = await import('../src/renderer/editor/extensions')
    const data =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6f1cAAAAASUVORK5CYII='
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content:
          kind === 'standalone'
            ? [
                {
                  type: 'docProtected',
                  attrs: { blockType: 'image', imageDataUrl: `data:image/png;base64,${data}` },
                },
              ]
            : kind === 'inline'
              ? [
                  {
                    type: 'docParagraph',
                    content: [
                      {
                        type: 'docInlineImage',
                        attrs: { dataUrl: `data:image/png;base64,${data}` },
                      },
                    ],
                  },
                ]
              : [
                  {
                    type: 'docTable',
                    content: [
                      {
                        type: 'docTableRow',
                        content: [
                          {
                            type: 'docTableCell',
                            content: [
                              {
                                type: 'docParagraph',
                                content: [
                                  {
                                    type: 'docInlineImage',
                                    attrs: { dataUrl: `data:image/png;base64,${data}` },
                                  },
                                ],
                              },
                            ],
                          },
                        ],
                      },
                    ],
                  },
                ],
      },
    })
    const active = editor
    runtime = createDocsRuntime({
      current: () => ({
        editor: active,
        identity: active,
        numIds: { bullet: null, ordered: null },
      }),
    })
    const router = createDocsCommandRouter(runtime)
    boundServer = createDocsMcpServer({
      version: 'test',
      endpoint: {
        request: (message, signal) => router.handle(message, signal),
        closeCaller: (caller) => router.disconnect(caller),
      },
    })
    const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair()
    await boundServer.connect(serverTransport)
    client = new Client({ name: 'media-harness', version: '1' }, { capabilities: {} })
    await client.connect(clientTransport)
    const context = await client.callTool({ name: 'get_document_context', arguments: {} })
    const ref = /media:[a-z0-9-]+:\d+:0\.0/.exec(JSON.stringify(context))?.[0]
    expect(ref).toBeDefined()
    const image = CallToolResultSchema.parse(
      await client.callTool({ name: 'read_document_media', arguments: { mediaRef: ref } }),
    )
    expect(image.isError).not.toBe(true)
    expect(image.content).toContainEqual({ type: 'image', data, mimeType: 'image/png' })
    expect(image.structuredContent).not.toHaveProperty('media')
    expect(JSON.stringify(image.content.filter((item) => item.type === 'text'))).not.toContain(data)
    router.dispose()
  },
)

it('returns mind-map display data to the Harness without claiming an editor chat card was rendered', async () => {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<p>Original</p>',
  })
  const active = editor
  runtime = createDocsRuntime({
    current: () => ({ editor: active, identity: active, numIds: { bullet: null, ordered: null } }),
  })
  const router = createDocsCommandRouter(runtime)
  boundServer = createDocsMcpServer({
    version: 'test',
    endpoint: {
      request: (message, signal) => router.handle(message, signal),
      closeCaller: (caller) => router.disconnect(caller),
    },
  })
  const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair()
  await boundServer.connect(serverTransport)
  client = new Client({ name: 'mindmap-harness', version: '1' }, { capabilities: {} })
  await client.connect(clientTransport)
  expect(
    (await client.listTools()).tools.find((tool) => tool.name === 'generate_mindmap')?.description,
  ).not.toContain('diagram card in the chat')
  const result = CallToolResultSchema.parse(
    await client.callTool({
      name: 'generate_mindmap',
      arguments: {
        root: { label: 'Original', block: 0, children: [{ label: 'Original', block: 0 }] },
      },
    }),
  )
  expect(result.isError, JSON.stringify(result.content)).not.toBe(true)
  expect(result._meta).toMatchObject({
    'io.foraoffice/display': { kind: 'mindmap', tree: { label: 'Original', block: 0 } },
  })
  expect(result.structuredContent).not.toHaveProperty('display')
  expect(JSON.stringify(result.content)).not.toContain('rendered as an interactive card')
  expect(active.getText()).toBe('Original')
  router.dispose()
})

it('reads full live blocks through standard MCP and applies an edit through the same runtime', async () => {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<p>Unsaved manual text</p><p>Keep this paragraph</p>',
  })
  const active = editor
  runtime = createDocsRuntime({
    current: () => ({
      editor: active,
      identity: active,
      numIds: { bullet: null, ordered: null },
    }),
  })
  const router = createDocsCommandRouter(runtime)
  const port = await new Promise<number>((resolve) => {
    const socket = createServer().listen(0, '127.0.0.1', () => {
      const address = socket.address()
      if (!address || typeof address === 'string') throw new Error('No local TCP port')
      socket.close(() => resolve(address.port))
    })
  })
  service = new McpServerService({
    port,
    toolsFactory: () => {
      const host = createSessionHost()
      const docs: DocsControl = {
        openBlankTab: async () => 1,
        runCommand: async (_id, command, payload, invocation) =>
          router
            .handle(
              {
                requestId: crypto.randomUUID(),
                callerId: invocation?.callerId,
                command: command as McpEditorCommand,
                payload,
              },
              invocation?.signal,
            )
            .catch((error: unknown) => {
              if (error instanceof DocsOperationError)
                throw new McpToolExecutionError(error.message, { ...error.result })
              throw error
            }),
      }
      return [
        ...createSessionTools([documentDriver(docs)], host),
        ...createDocumentTools(
          {
            version: 'test',
            defaultSaveDir: () => '.',
            docs,
          },
          host,
        ),
      ]
    },
  })
  await service.start()
  client = new Client({ name: 'docs-integration', version: '1' })
  await client.connect(new StreamableHTTPClientTransport(new URL(`${service.getUrl()}/mcp`)))
  await client.callTool({ name: 'create_session', arguments: { family: 'docx' } })
  await expectStandardToolDiscovery(client)
  const read = await client.callTool({
    name: 'read_blocks',
    arguments: { startBlockIndex: 0, endBlockIndex: 1 },
  })
  expect(read.isError).not.toBe(true)
  expect(JSON.stringify(read.content)).toContain('Unsaved manual text')
  const ops = [
    { op: 'setFont', target: { blockIndexes: [0] }, color: '#FF0000' },
    { op: 'setParagraphFormat', target: { blockIndexes: [1] }, align: 'center' },
  ]
  const before = active.getJSON()
  const preview = await client.callTool({ name: 'apply_ops', arguments: { ops, dryRun: true } })
  expect(preview.structuredContent).toMatchObject({
    committed: false,
    receipt: { outcome: 'preview', revisionBefore: 0, revisionAfter: 0 },
  })
  const invalid = await client.callTool({
    name: 'apply_ops',
    arguments: { ops: [...ops, { op: 'setFont', color: '#FF0000' }] },
  })
  expect(invalid).toMatchObject({
    isError: true,
    structuredContent: { committed: false, output: expect.stringContaining('setFont') },
  })
  expect(active.getJSON()).toEqual(before)
  expect(
    (await client.callTool({ name: 'apply_ops', arguments: { ops } })).structuredContent,
  ).toMatchObject({ committed: true, receipt: { outcome: 'applied' } })
  expect(active.getJSON()).not.toEqual(before)
  expect(
    (await client.callTool({ name: 'document_history', arguments: { action: 'undo' } })).isError,
  ).not.toBe(true)
  expect(active.getJSON()).toEqual(before)
  const edit = await client.callTool({
    name: 'replace_blocks',
    arguments: {
      startBlockIndex: 0,
      endBlockIndex: 0,
      html: '<p>Agent result</p>',
    },
  })
  expect(edit.isError).not.toBe(true)
  expect(edit.structuredContent).toMatchObject({
    committed: true,
    revision: expect.any(Number),
    documentId: expect.any(String),
  })
  expect(active.getText()).toBe('Agent result\n\nKeep this paragraph')
  active.commands.insertContentAt(1, 'Manual ')
  await client.listTools()
  const stale = await client.callTool({
    name: 'replace_blocks',
    arguments: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>Wrong</p>' },
  })
  expect(stale.isError).toBe(true)
  expect(stale.structuredContent).toMatchObject({ committed: false, errorCode: 'stale_target' })
})
