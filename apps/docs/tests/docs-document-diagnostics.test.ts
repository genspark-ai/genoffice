import { afterEach, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createDocsInspection } from '../src/renderer/sdk/inspection'
import { createDocsAnnotations } from '../src/renderer/sdk/annotations'
import { queryDocumentDiagnostics } from '../src/renderer/runtime/document-diagnostics'
import { createDocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { createDocsCommandRouter } from '../src/renderer/runtime/command-router'
import { createDocsMcpServer } from '../src/services/docs-mcp'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import type { DocsOperationResult } from '../src/shared/operation-result'

const cleanup: Array<() => void | Promise<void>> = []
afterEach(async () => {
  for (const dispose of cleanup.splice(0).reverse()) await dispose()
})

function fixture() {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<p>bad bad</p>',
  })
  const inspection = createDocsInspection(() => ({ editor, identity: editor }))
  const annotations = createDocsAnnotations(() => ({ editor, identity: editor }), inspection)
  annotations.activate()
  cleanup.push(
    () => editor.destroy(),
    () => inspection.dispose(),
    () => annotations.dispose(),
  )
  return { editor, inspection, annotations }
}

function runtimeFixture() {
  const base = fixture()
  let available = true
  const runtime = createDocsRuntime({
    current: () => ({
      editor: base.editor,
      identity: base.editor,
      numIds: { bullet: null, ordered: null },
      ...(available ? { annotations: base.annotations } : {}),
    }),
  })
  const caller = runtime.connect({ id: 'diagnostic-reader' })
  caller.context()
  cleanup.push(() => runtime.dispose())
  return {
    ...base,
    runtime,
    caller,
    removeProvider() {
      available = false
    },
  }
}

it('advertises optional runtime diagnostics accurately and does not widen a stale writing baseline', async () => {
  const { editor, annotations, caller, removeProvider } = runtimeFixture()
  annotations.setRules('host', [{ id: 'bad', pattern: 'bad', message: 'Check this word' }])
  const capabilities = await caller.execute({
    id: 'capabilities',
    name: 'get_editor_capabilities',
    input: {},
  })
  expect(JSON.parse(capabilities.output).tools).toContainEqual(
    expect.objectContaining({ name: 'read_document_diagnostics', available: true }),
  )
  editor.commands.insertContentAt(1, 'Changed ')
  const before = editor.getJSON()
  const diagnosticRead = await caller.execute({
    id: 'diagnostics',
    name: 'read_document_diagnostics',
    input: {},
  })
  expect(diagnosticRead).toMatchObject({
    committed: false,
    mutated: false,
    receipt: { outcome: 'read' },
  })
  expect(JSON.parse(diagnosticRead.output)).toMatchObject({
    documentId: diagnosticRead.documentId,
    revision: diagnosticRead.revision,
    checksComplete: true,
  })
  const write = await caller.execute({
    id: 'stale-write',
    name: 'apply_ops',
    input: {
      ops: [
        {
          op: 'findReplace',
          target: { blockIndexes: [0] },
          find: 'Changed',
          replace: 'Unexpected',
        },
      ],
    },
  })
  expect(write).toMatchObject({ isError: true, committed: false, errorCode: 'stale_target' })
  expect(editor.getJSON()).toEqual(before)
  removeProvider()
  const unavailable = await caller.execute({
    id: 'unavailable',
    name: 'read_document_diagnostics',
    input: {},
  })
  expect(unavailable).toMatchObject({ isError: true, committed: false, errorCode: 'unavailable' })
  const absentCapabilities = await caller.execute({
    id: 'absent-capabilities',
    name: 'get_editor_capabilities',
    input: {},
  })
  expect(JSON.parse(absentCapabilities.output).tools).toContainEqual(
    expect.objectContaining({
      name: 'read_document_diagnostics',
      available: false,
      errorCode: 'unavailable',
    }),
  )
})

it('exposes current and pending host checks through independent standard MCP read-only connections', async () => {
  const { editor, annotations, inspection, runtime } = runtimeFixture()
  let calls = 0
  let complete!: (
    values: Array<{ from: number; to: number; message: string; code: string }>,
  ) => void
  annotations.registerChecker('host', () => {
    calls++
    return new Promise((resolve) => {
      complete = resolve
    })
  })
  const router = createDocsCommandRouter(runtime)
  cleanup.push(() => router.dispose())
  const connect = async (readOnly: boolean) => {
    const server = createDocsMcpServer({
      version: 'diagnostic-contract',
      readOnly,
      endpoint: {
        request: (message, signal) => router.handle(message, signal),
        closeCaller: (id) => router.disconnect(id),
      },
    })
    const client = new Client({ name: readOnly ? 'reviewer' : 'editor', version: '1' })
    const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair()
    await server.connect(serverTransport)
    await client.connect(clientTransport)
    cleanup.push(
      () => server.close(),
      () => client.close(),
    )
    return client
  }
  const writer = await connect(false)
  const reviewer = await connect(true)
  const tools = (await reviewer.listTools()).tools
  expect(tools).toContainEqual(
    expect.objectContaining({
      name: 'read_document_diagnostics',
      annotations: expect.objectContaining({ readOnlyHint: true }),
    }),
  )
  expect(tools.map((tool) => tool.name)).not.toContain('apply_ops')
  const read = async (client: Client) => {
    const result = await client.callTool({ name: 'read_document_diagnostics', arguments: {} })
    expect(result.isError).not.toBe(true)
    return JSON.parse((result.structuredContent as unknown as DocsOperationResult).output)
  }
  const pending = await read(reviewer)
  expect(pending).toMatchObject({
    checksComplete: false,
    items: [{ kind: 'source', source: 'host', status: 'pending' }],
  })
  expect(calls).toBe(1)
  complete([{ from: 1, to: 4, message: 'Current host diagnostic', code: 'HOST_CODE' }])
  await Promise.resolve()
  const approved = await read(writer)
  const reviewed = await read(reviewer)
  expect(approved).toEqual(reviewed)
  expect(reviewed).toMatchObject({
    checksComplete: true,
    items: [
      expect.objectContaining({ kind: 'source', status: 'complete' }),
      expect.objectContaining({ kind: 'diagnostic', code: 'HOST_CODE', stale: false }),
    ],
  })
  const before = editor.getJSON()
  const denied = await reviewer
    .callTool({
      name: 'apply_ops',
      arguments: {
        ops: [
          {
            op: 'findReplace',
            target: { blockIndexes: [0] },
            find: 'bad',
            replace: 'forbidden',
            match: 'all',
          },
        ],
      },
    })
    .catch((error: unknown) => ({
      isError: true,
      content: [{ type: 'text', text: String(error) }],
    }))
  expect(denied.isError).toBe(true)
  expect(JSON.stringify(denied.content)).toMatch(/not found|unknown tool|not registered/i)
  expect(editor.getJSON()).toEqual(before)
  annotations.setRanges('host', [
    { ...inspection.snapshot(), from: 1, to: 4, message: 'Host manual result' },
  ])
  editor.commands.insertContentAt(1, 'new ')
  const stale = await read(reviewer)
  expect(stale.checksComplete).toBe(false)
  const staleFinding = stale.items.find((item: { kind: string }) => item.kind === 'diagnostic')
  expect(staleFinding).toMatchObject({ stale: true })
  expect(staleFinding).not.toHaveProperty('targets')
})

it('returns bounded revision-bound current findings with source coverage and executable carrier targets', () => {
  const { editor, annotations } = fixture()
  annotations.setRules('host', [
    { id: 'HOST_BAD', pattern: 'bad', message: 'Check value', severity: 'error' },
  ])
  const before = editor.getJSON()
  const first = queryDocumentDiagnostics(editor, annotations, 'live-document', 7, { limit: 2 })
  expect(first).toMatchObject({
    documentId: 'live-document',
    revision: 7,
    checksComplete: true,
    page: { complete: false },
  })
  expect(first.items[0]).toMatchObject({
    kind: 'source',
    source: 'host',
    status: 'complete',
    coverage: { kind: 'document-text' },
  })
  expect(first.items[1]).toMatchObject({
    kind: 'diagnostic',
    source: 'host',
    code: 'HOST_BAD',
    severity: 'error',
    targets: [{ documentId: 'live-document', revision: 7, expectedType: 'docParagraph' }],
  })
  const second = queryDocumentDiagnostics(editor, annotations, 'live-document', 7, {
    cursor: first.page.nextCursor,
  })
  expect(second.items).toHaveLength(1)
  expect(second.page.complete).toBe(true)
  expect(editor.getJSON()).toEqual(before)
  expect(editor.commands.undo()).toBe(false)
})

it('expires diagnostic continuations after checker completion even without a document revision change', async () => {
  const { editor, annotations } = fixture()
  let complete!: (value: Array<{ from: number; to: number; message: string }>) => void
  annotations.setRules('first', [{ id: 'bad', pattern: 'bad', message: 'Check' }])
  annotations.registerChecker(
    'async',
    () =>
      new Promise((resolve) => {
        complete = resolve
      }),
  )
  const first = queryDocumentDiagnostics(editor, annotations, 'live', 3, { limit: 1 })
  expect(first.checksComplete).toBe(false)
  complete([])
  await Promise.resolve()
  expect(() =>
    queryDocumentDiagnostics(editor, annotations, 'live', 3, { cursor: first.page.nextCursor }),
  ).toThrow(/stale/i)
})

it('does not return executable targets for stale manual findings and refuses stale/filter-modified cursors', async () => {
  const { editor, inspection, annotations } = fixture()
  annotations.setRanges('manual', [
    { ...inspection.snapshot(), from: 1, to: 4, message: 'Old finding' },
  ])
  editor.commands.insertContentAt(1, 'new ')
  await annotations.refresh()
  const current = queryDocumentDiagnostics(editor, annotations, 'live', 8, {})
  const finding = current.items.find((item) => item.kind === 'diagnostic')
  expect(finding).toMatchObject({ stale: true, source: 'manual' })
  expect(finding).not.toHaveProperty('targets')
  const first = queryDocumentDiagnostics(editor, annotations, 'live', 8, { limit: 1 })
  expect(() =>
    queryDocumentDiagnostics(editor, annotations, 'live', 9, { cursor: first.page.nextCursor }),
  ).toThrow(/stale/i)
  expect(() =>
    queryDocumentDiagnostics(editor, annotations, 'live', 8, {
      cursor: first.page.nextCursor,
      source: 'manual',
    }),
  ).toThrow(/cannot change/i)
})

it('does not expose host diagnostic messages anchored in withheld content', () => {
  const { editor, inspection, annotations } = fixture()
  editor.commands.setContent({
    type: 'doc',
    content: [
      {
        type: 'docParagraph',
        content: [
          {
            type: 'text',
            text: 'private',
            marks: [{ type: 'redaction', attrs: { label: 'private' } }],
          },
        ],
      },
    ],
  })
  annotations.setRanges('host', [
    { ...inspection.snapshot(), from: 1, to: 8, message: 'private value copied by checker' },
  ])
  const result = queryDocumentDiagnostics(editor, annotations, 'live', 1, {})
  expect(JSON.stringify(result)).not.toContain('private value')
  expect(result.withheldDiagnostics).toBe(1)
})
