import { afterEach, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createDocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { createDocsMcpServer } from '../src/services/docs-mcp'
import { createDocsCommandRouter } from '../src/renderer/runtime/command-router'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import type { DocsOperationResult } from '../src/shared/operation-result'
import JSZip from 'jszip'

const cleanup: Array<() => void> = []
afterEach(() => {
  cleanup
    .splice(0)
    .reverse()
    .forEach((dispose) => dispose())
  vi.useRealTimers()
})
const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6f1cAAAAASUVORK5CYII='

function open(table = false) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: table
      ? '<table><tr><td><p>KEEP CELL</p></td></tr></table><p>AFTER</p>'
      : {
          type: 'doc',
          content: [
            {
              type: 'docParagraph',
              content: [
                { type: 'text', text: 'KEEP ' },
                { type: 'docInlineImage', attrs: { dataUrl: PNG } },
              ],
            },
          ],
        },
  })
  const runtime = createDocsRuntime({
    current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
  })
  const client = runtime.connect({ id: 'draft-safety' })
  cleanup.push(
    () => editor.destroy(),
    () => runtime.dispose(),
  )
  const context = client.context()
  return { editor, client, context }
}

it.each(['unbound', 'missing', 'wrapped', 'mixed', 'wrong-document', 'expired'])(
  'rejects %s draft image references before accepting or rendering the update',
  async (kind) => {
    vi.useFakeTimers()
    const { editor, client, context } = open()
    const reference = `media:${context.documentId}:${context.revision}:0.0`
    const image =
      kind === 'unbound'
        ? 'doc:0.0'
        : kind === 'missing'
          ? `media:${context.documentId}:${context.revision}:99.0`
          : kind === 'wrong-document'
            ? `media:other:${context.revision}:0.0`
            : reference
    const html =
      kind === 'wrapped'
        ? `<p>PROBE</p><ul><img src="${reference}"></ul>`
        : `<p>PROBE<img src="${image}">${kind === 'mixed' ? '<img src="doc:99.0">' : ''}</p>`
    if (kind === 'expired') {
      editor.commands.insertContent('EDIT')
      client.context()
    }
    const start = await client.execute({
      id: 'start',
      name: 'start_document_draft',
      input: { afterBlockIndex: 0 },
    })
    const { draftId } = JSON.parse(start.output)
    const before = client.context()
    const update = await client.execute({
      id: 'update',
      name: 'update_document_draft',
      input: { draftId, sequence: 1, html },
    })
    expect(update).toMatchObject({
      isError: true,
      committed: false,
      mutated: false,
      revision: before.revision,
    })
    await vi.advanceTimersByTimeAsync(400)
    expect(client.context().text).toBe(before.text)
    expect(
      await client.execute({
        id: 'finish',
        name: 'finish_document_draft',
        input: { draftId, decision: 'retain_partial' },
      }),
    ).toMatchObject({ isError: true, committed: false })
    expect(client.context().text).toBe(before.text)
  },
)

it('keeps validated source images across preview revisions and commits their exact bytes', async () => {
  vi.useFakeTimers()
  const { client, context } = open()
  const reference = `media:${context.documentId}:${context.revision}:0.0`
  const start = await client.execute({
    id: 'start',
    name: 'start_document_draft',
    input: { afterBlockIndex: 0 },
  })
  const { draftId, referenceSnapshot } = JSON.parse(start.output)
  expect(referenceSnapshot).toEqual({ documentId: context.documentId, revision: context.revision })
  for (const sequence of [1, 2]) {
    expect(
      await client.execute({
        id: 'update',
        name: 'update_document_draft',
        input: { draftId, sequence, html: `<p>DRAFT ${sequence}<img src="${reference}"></p>` },
      }),
    ).not.toHaveProperty('isError', true)
    await vi.advanceTimersByTimeAsync(400)
  }
  expect(
    await client.execute({
      id: 'finish',
      name: 'finish_document_draft',
      input: { draftId, decision: 'finalize' },
    }),
  ).toMatchObject({ committed: true, mutated: true })
  const images = await client.execute({
    id: 'images',
    name: 'query_document',
    input: { domain: 'images' },
  })
  const items = JSON.parse(images.output).items
  expect(items).toHaveLength(2)
  for (const item of items)
    expect(
      await client.execute({
        id: 'bytes',
        name: 'read_document_media',
        input: { mediaRef: item.mediaRef },
      }),
    ).toMatchObject({ media: { data: PNG.split(',')[1], mimeType: 'image/png' } })
})

it('accepts exact whole-document structural intent captured before the preview', async () => {
  vi.useFakeTimers()
  const { client } = open(true)
  const query = await client.execute({ id: 'query', name: 'query_document', input: {} })
  const ids = JSON.parse(query.output).items.flatMap(
    (item: { structures: Array<{ id: string }> }) =>
      item.structures.map((structure) => structure.id),
  )
  const start = await client.execute({
    id: 'start',
    name: 'start_document_draft',
    input: { replaceDocument: true },
  })
  const { draftId } = JSON.parse(start.output)
  expect(
    await client.execute({
      id: 'duplicate-intent',
      name: 'update_document_draft',
      input: { draftId, sequence: 1, html: '<p>NEW DOC</p>', deleteStructures: [...ids, ...ids] },
    }),
  ).toMatchObject({ isError: true, committed: false, revision: start.revision })
  expect(
    await client.execute({
      id: 'update',
      name: 'update_document_draft',
      input: { draftId, sequence: 1, html: '<p>NEW DOC</p>', deleteStructures: ids },
    }),
  ).not.toHaveProperty('isError', true)
  await vi.advanceTimersByTimeAsync(400)
  expect(
    await client.execute({
      id: 'finish',
      name: 'finish_document_draft',
      input: { draftId, decision: 'retain_partial' },
    }),
  ).toMatchObject({ committed: true })
  expect(client.context().text).toContain('NEW DOC')
  expect(client.context().text).not.toContain('KEEP CELL')
})

it('refuses finalization after the captured document-side source changes without publishing a pending preview', async () => {
  vi.useFakeTimers()
  const { editor } = open()
  let version = 0
  const runtime = createDocsRuntime({
    current: () => ({
      editor,
      identity: editor,
      stateVersion: version,
      numIds: { bullet: null, ordered: null },
    }),
  })
  cleanup.push(() => runtime.dispose())
  const client = runtime.connect({ id: 'side-state' })
  const context = client.context()
  const start = await client.execute({
    id: 'start',
    name: 'start_document_draft',
    input: { afterBlockIndex: 0 },
  })
  const { draftId } = JSON.parse(start.output)
  await client.execute({
    id: 'update',
    name: 'update_document_draft',
    input: {
      draftId,
      sequence: 1,
      html: `<p>PREVIEW<img src="media:${context.documentId}:${context.revision}:0.0"></p>`,
    },
  })
  version++
  await vi.advanceTimersByTimeAsync(400)
  const before = client.context()
  expect(before.text).not.toContain('PREVIEW')
  expect(before.revision).toBe(start.revision + 1)
  expect(
    await client.execute({
      id: 'finish',
      name: 'finish_document_draft',
      input: { draftId, decision: 'finalize' },
    }),
  ).toMatchObject({
    isError: true,
    errorCode: 'stale_target',
    committed: false,
    revision: before.revision,
  })
  await vi.advanceTimersByTimeAsync(400)
  expect(client.context()).toMatchObject({ text: before.text, revision: before.revision })
  expect(
    await client.execute({
      id: 'discard',
      name: 'finish_document_draft',
      input: { draftId, decision: 'discard' },
    }),
  ).toMatchObject({ committed: false })
})

it('does not publish a pending reference preview after its document binding is replaced', async () => {
  vi.useFakeTimers()
  const { editor } = open()
  let identity = {}
  const runtime = createDocsRuntime({
    current: () => ({ editor, identity, numIds: { bullet: null, ordered: null } }),
  })
  cleanup.push(() => runtime.dispose())
  const client = runtime.connect({ id: 'old-draft' })
  const context = client.context()
  const start = await client.execute({
    id: 'start',
    name: 'start_document_draft',
    input: { afterBlockIndex: 0 },
  })
  const { draftId } = JSON.parse(start.output)
  await client.execute({
    id: 'update',
    name: 'update_document_draft',
    input: {
      draftId,
      sequence: 1,
      html: `<p>STALE PREVIEW<img src="media:${context.documentId}:${context.revision}:0.0"></p>`,
    },
  })
  identity = {}
  await vi.advanceTimersByTimeAsync(400)
  const next = runtime.connect({ id: 'replacement' }).context()
  expect(next.text).toContain('KEEP')
  expect(next.text).not.toContain('STALE PREVIEW')
})

it('refuses stale reference finalization after native body edits without removing or republishing its preview', async () => {
  vi.useFakeTimers()
  const { editor, client, context } = open()
  const start = await client.execute({
    id: 'start',
    name: 'start_document_draft',
    input: { afterBlockIndex: 0 },
  })
  const { draftId } = JSON.parse(start.output)
  await client.execute({
    id: 'update',
    name: 'update_document_draft',
    input: {
      draftId,
      sequence: 1,
      html: `<p>PREVIEW<img src="media:${context.documentId}:${context.revision}:0.0"></p>`,
    },
  })
  await vi.advanceTimersByTimeAsync(400)
  editor.commands.insertContentAt(1, 'USER')
  const before = client.context()
  expect(
    await client.execute({
      id: 'finish',
      name: 'finish_document_draft',
      input: { draftId, decision: 'retain_partial' },
    }),
  ).toMatchObject({
    committed: false,
    mutated: false,
    errorCode: 'stale_target',
    revision: before.revision,
  })
  await vi.advanceTimersByTimeAsync(400)
  expect(client.context()).toMatchObject({ text: before.text, revision: before.revision })
  expect(
    await client.execute({
      id: 'discard',
      name: 'finish_document_draft',
      input: { draftId, decision: 'discard' },
    }),
  ).toMatchObject({ committed: false, mutated: false })
  expect(client.context().text).toContain('USER')
  expect(client.context().text).not.toContain('PREVIEW')
})

it('does not split typing history for starting, refusing and discarding an empty draft', async () => {
  const { editor, client } = open()
  editor.commands.setTextSelection(1)
  editor.commands.insertContent('a')
  client.context()
  const start = await client.execute({
    id: 'start',
    name: 'start_document_draft',
    input: { afterBlockIndex: 0 },
  })
  const { draftId } = JSON.parse(start.output)
  await client.execute({
    id: 'reject',
    name: 'update_document_draft',
    input: { draftId, sequence: 1, html: '<p>WRONG<img src="doc:99.0"></p>' },
  })
  editor.commands.insertContent('b')
  await client.execute({
    id: 'discard',
    name: 'finish_document_draft',
    input: { draftId, decision: 'discard' },
  })
  client.context()
  await client.execute({ id: 'undo', name: 'document_history', input: { action: 'undo' } })
  expect(editor.state.doc.textContent).toBe('KEEP ')
})

it.each(['retain', 'complete'])(
  'buffers a split canonical image tag and handles %s without publishing an unvalidated image',
  async (decision) => {
    vi.useFakeTimers()
    const { client, context } = open()
    const start = await client.execute({
      id: 'start',
      name: 'start_document_draft',
      input: { afterBlockIndex: 0 },
    })
    const { draftId } = JSON.parse(start.output)
    await client.execute({
      id: 'safe',
      name: 'update_document_draft',
      input: { draftId, sequence: 1, html: '<p>SAFE</p>' },
    })
    await vi.advanceTimersByTimeAsync(400)
    const before = client.context()
    const split = await client.execute({
      id: 'split',
      name: 'update_document_draft',
      input: {
        draftId,
        sequence: 2,
        html: `<p>SAFE<img src="media:${context.documentId}:${context.revision}:0.`,
      },
    })
    expect(split.isError, split.output).not.toBe(true)
    expect(JSON.parse(split.output)).toMatchObject({
      buffering: true,
      validatedSequence: 1,
      sequence: 2,
    })
    await vi.advanceTimersByTimeAsync(400)
    expect(client.context().text).toBe(before.text)
    expect(
      await client.execute({
        id: 'premature',
        name: 'finish_document_draft',
        input: { draftId, decision: 'finalize' },
      }),
    ).toMatchObject({ isError: true, committed: false })
    if (decision === 'complete') {
      expect(
        await client.execute({
          id: 'completed-image',
          name: 'update_document_draft',
          input: {
            draftId,
            sequence: 3,
            html: `<p>SAFE<img src="media:${context.documentId}:${context.revision}:0.0"></p>`,
          },
        }),
      ).not.toHaveProperty('isError', true)
      const committed = await client.execute({
        id: 'commit',
        name: 'finish_document_draft',
        input: { draftId, decision: 'finalize' },
      })
      expect(committed).toMatchObject({ committed: true })
      expect(JSON.parse(committed.output).result).toMatchObject({ partial: false })
      expect(JSON.parse(committed.output).result.omittedTail).toBeUndefined()
      const images = await client.execute({
        id: 'images',
        name: 'query_document',
        input: { domain: 'images' },
      })
      expect(JSON.parse(images.output).items).toHaveLength(2)
      return
    }
    const retained = await client.execute({
      id: 'retain',
      name: 'finish_document_draft',
      input: { draftId, decision: 'retain_partial' },
    })
    expect(retained).toMatchObject({ committed: true })
    expect(JSON.parse(retained.output).result).toMatchObject({
      partial: true,
      validatedSequence: 1,
      omittedTail: true,
    })
    expect(client.context().text).toContain('SAFE')
  },
)

it('checks whole-document table/field deletion through standard MCP and actual DOCX save/reopen', async () => {
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml:
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>KEEP CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:fldSimple w:instr="DATE"><w:r><w:t>OLD FIELD</w:t></w:r></w:fldSimple></w:p>',
    }),
  )
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks),
  })
  let saved: Uint8Array | undefined
  const runtime = createDocsRuntime({
    current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    save: async () => {
      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      saved = await saveDocx(parsed, plan.saveBlocks)
      return { ok: true, committed: true, pending: false }
    },
  })
  cleanup.push(
    () => editor.destroy(),
    () => runtime.dispose(),
  )
  const router = createDocsCommandRouter(runtime)
  const server = createDocsMcpServer({
    version: 'test',
    endpoint: {
      request: (message, signal) => router.handle(message, signal),
      closeCaller: (id) => router.disconnect(id),
    },
  })
  const client = new Client({ name: 'draft-test', version: 'test' })
  const [a, b] = InMemoryTransport.createLinkedPair()
  await server.connect(a)
  await client.connect(b)
  const call = async (name: string, input: Record<string, unknown> = {}) =>
    (await client.callTool({ name, arguments: input }))
      .structuredContent as unknown as DocsOperationResult
  try {
    await call('read_document')
    const query = await call('query_document')
    const ids = JSON.parse(query.output).items.flatMap(
      (item: { structures: Array<{ id: string }> }) =>
        item.structures.map((structure) => structure.id),
    )
    const start = await call('start_document_draft', { replaceDocument: true })
    const { draftId } = JSON.parse(start.output)
    expect(
      await call('update_document_draft', { draftId, sequence: 1, html: '<p>NEW DOC</p>' }),
    ).toMatchObject({ isError: true, committed: false, revision: start.revision })
    expect(
      await call('update_document_draft', {
        draftId,
        sequence: 1,
        html: '<p>NEW DOC</p>',
        deleteStructures: ids,
      }),
    ).not.toHaveProperty('isError', true)
    expect(await call('finish_document_draft', { draftId, decision: 'finalize' })).toMatchObject({
      committed: true,
    })
    expect(await call('save_document', { path: '/draft.docx', overwrite: true })).toMatchObject({
      committed: true,
      pending: false,
    })
    const reopened = await parseDocx(saved!)
    expect(
      reopened.blocks.map((block) => ({
        type: block.type,
        text: block.runs?.map((run) => run.text).join(''),
      })),
    ).toEqual([
      { type: 'paragraph', text: 'NEW DOC' },
      { type: 'passthrough', text: undefined },
    ])
    expect(reopened.blocks[1].originalXml).toMatch(/^<w:sectPr[\s>]/)
    expect(reopened.blocks[0].runs?.map((run) => run.text).join('')).toBe('NEW DOC')
    const xml = await (await JSZip.loadAsync(saved!)).file('word/document.xml')!.async('string')
    expect(xml).not.toMatch(/<w:tbl[\s>]|<w:fldSimple[\s>]|<w:instrText[\s>]/)
    expect(xml).not.toContain('KEEP CELL')
    expect(xml).not.toContain('OLD FIELD')
  } finally {
    await client.close()
    await server.close()
    router.dispose()
  }
})

it('refuses protected whole-document deletion without exact intent before any preview is published', async () => {
  vi.useFakeTimers()
  const { client } = open(true)
  const start = await client.execute({
    id: 'start',
    name: 'start_document_draft',
    input: { replaceDocument: true },
  })
  const { draftId } = JSON.parse(start.output)
  const before = client.context()
  expect(
    await client.execute({
      id: 'update',
      name: 'update_document_draft',
      input: { draftId, sequence: 1, html: '<p>REPLACED</p>' },
    }),
  ).toMatchObject({ isError: true, committed: false, revision: before.revision })
  await vi.advanceTimersByTimeAsync(400)
  expect(client.context().text).toBe(before.text)
  expect(
    await client.execute({
      id: 'finish',
      name: 'finish_document_draft',
      input: { draftId, decision: 'finalize' },
    }),
  ).toMatchObject({ isError: true, committed: false })
  expect(client.context().text).toBe(before.text)
})
