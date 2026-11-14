// Modified for ForaOffice: explicit image writes are validated before mutation.
import { afterEach, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import JSZip from 'jszip'
import { inlinePictureDrawingXml, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { createDocsMcpServer } from '../src/services/docs-mcp'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createDocsRuntime, type DocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { createDocsCommandRouter } from '../src/renderer/runtime/command-router'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'

const editors: Editor[] = []
const runtimes: DocsRuntime[] = []
afterEach(() => {
  runtimes.splice(0).forEach((runtime) => runtime.dispose())
  editors.splice(0).forEach((editor) => editor.destroy())
})

const paragraph = (text: string) => ({
  type: 'docParagraph',
  content: [{ type: 'text', text }],
})
const imageParagraph = (data = 'AQ==') => ({
  type: 'docParagraph',
  content: [
    { type: 'text', text: 'BEFORE ' },
    { type: 'docInlineImage', attrs: { dataUrl: `data:image/png;base64,${data}` } },
  ],
})

function fixture(content: object[] = [paragraph('ORIGINAL')]) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  const runtime = createDocsRuntime({
    current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
  })
  editors.push(editor)
  runtimes.push(runtime)
  const caller = runtime.connect({ id: 'image-writer' })
  return { editor, runtime, caller }
}

it('refuses an unresolved explicit image without changing the document or revision', async () => {
  const { caller } = fixture()
  const before = caller.context()
  const result = await caller.execute({
    id: 'invalid-image',
    name: 'replace_blocks',
    input: {
      startBlockIndex: 0,
      endBlockIndex: 0,
      html: '<p>REFERENCE_PROBE <img src="doc:99999.0"></p>',
    },
  })
  expect(result).toMatchObject({
    isError: true,
    committed: false,
    mutated: false,
    errorCode: 'invalid_input',
    revision: before.revision,
  })
  const after = await caller.execute({
    id: 'read',
    name: 'read_blocks',
    input: { startBlockIndex: 0, endBlockIndex: 0 },
  })
  expect(after.output).toContain('<p>ORIGINAL</p>')
  expect(after.revision).toBe(before.revision)
})

it('returns writable revision-bound HTML and preserves the requested image', async () => {
  const { caller } = fixture([imageParagraph()])
  const before = caller.context()
  const read = await caller.execute({
    id: 'read',
    name: 'read_blocks',
    input: { startBlockIndex: 0, endBlockIndex: 0 },
  })
  const reference = `media:${before.documentId}:${before.revision}:0.0`
  expect(read.output).toContain(`<img src="${reference}">`)
  const html = read.output.slice(0, read.output.indexOf('\nAuthorized document images'))
  expect(
    await caller.execute({
      id: 'replace',
      name: 'replace_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0, html: html.replace('BEFORE', 'AFTER') },
    }),
  ).toMatchObject({ committed: true, mutated: true })
  const fresh = caller.context()
  expect(
    await caller.execute({
      id: 'image',
      name: 'read_document_media',
      input: { mediaRef: `media:${fresh.documentId}:${fresh.revision}:0.0` },
    }),
  ).toMatchObject({ media: { data: 'AQ==', mimeType: 'image/png' } })
})

it('does not confuse a visible cell image with a nested-table image omitted from its HTML', async () => {
  const nested = {
    type: 'docTable',
    content: [
      { type: 'docTableRow', content: [{ type: 'docTableCell', content: [imageParagraph()] }] },
    ],
  }
  const { caller } = fixture([
    {
      type: 'docTable',
      content: [
        {
          type: 'docTableRow',
          content: [{ type: 'docTableCell', content: [nested, imageParagraph('Ag==')] }],
        },
      ],
    },
  ])
  caller.context()
  const read = await caller.execute({
    id: 'nested-read',
    name: 'read_blocks',
    input: { startBlockIndex: 0, endBlockIndex: 0 },
  })
  const reference = /<img src="([^"]+)">/.exec(read.output)![1]
  expect(
    await caller.execute({
      id: 'visible-image',
      name: 'read_document_media',
      input: { mediaRef: reference },
    }),
  ).toMatchObject({ media: { data: 'Ag==' } })
})

it('refuses an image in an unsupported HTML position instead of silently dropping it', async () => {
  const { caller } = fixture([imageParagraph()])
  const before = caller.context()
  expect(
    await caller.execute({
      id: 'unsupported-position',
      name: 'replace_blocks',
      input: {
        startBlockIndex: 0,
        endBlockIndex: 0,
        html: `<p>WRONG</p><ul><img src="media:${before.documentId}:${before.revision}:0.0"></ul>`,
      },
    }),
  ).toMatchObject({ isError: true, committed: false, mutated: false, revision: before.revision })
  expect(caller.context().text).not.toContain('WRONG')
})

it('does not split existing undo history when an explicit image cannot be represented', async () => {
  const { editor, caller } = fixture([imageParagraph()])
  editor.commands.insertContent('a')
  const before = caller.context()
  await caller.execute({
    id: 'unsupported-image',
    name: 'replace_blocks',
    input: {
      startBlockIndex: 0,
      endBlockIndex: 0,
      html: `<p>WRONG</p><ul><img src="media:${before.documentId}:${before.revision}:0.0"></ul>`,
    },
  })
  editor.commands.insertContent('b')
  caller.context()
  expect(
    await caller.execute({ id: 'undo', name: 'document_history', input: { action: 'undo' } }),
  ).toMatchObject({ committed: true })
  expect(editor.getText()).toBe('BEFORE ')
})

it('rejects an old reference even after refresh when another image occupies its coordinate', async () => {
  const { caller, editor } = fixture([imageParagraph(), imageParagraph('Ag==')])
  const old = caller.context()
  editor.commands.setContent({ type: 'doc', content: [imageParagraph('Ag==')] })
  const fresh = caller.context()
  expect(fresh.revision).toBeGreaterThan(old.revision)
  expect(
    await caller.execute({
      id: 'stale-image',
      name: 'replace_blocks',
      input: {
        startBlockIndex: 0,
        endBlockIndex: 0,
        html: `<p>WRONG <img src="media:${old.documentId}:${old.revision}:0.0"></p>`,
      },
    }),
  ).toMatchObject({
    isError: true,
    committed: false,
    revision: fresh.revision,
    errorCode: 'stale_target',
    imageReferenceErrors: [{ reason: 'expired' }],
  })
  const after = await caller.execute({
    id: 'read-after',
    name: 'read_blocks',
    input: { startBlockIndex: 0, endBlockIndex: 0 },
  })
  expect(after.output).toContain('BEFORE')
  expect(after.output).not.toContain('WRONG')
})

it.each(['insert_content', 'replace_blocks', 'replace_selection'])(
  'refuses mixed valid and invalid references before %s changes anything',
  async (name) => {
    const { editor, caller } = fixture([imageParagraph()])
    editor.commands.setTextSelection({ from: 1, to: 7 })
    const before = caller.context()
    const result = await caller.execute({
      id: 'mixed',
      name,
      input: {
        html: `CHANGED <img src="media:${before.documentId}:${before.revision}:0.0"><img src="media:${before.documentId}:${before.revision}:0.99">`,
        startBlockIndex: 0,
        endBlockIndex: 0,
        afterBlockIndex: 0,
      },
    })
    expect(result).toMatchObject({
      isError: true,
      committed: false,
      mutated: false,
      revision: before.revision,
      imageReferenceErrors: [{ reason: 'missing' }],
    })
    expect(caller.context().text).not.toContain('CHANGED')
    expect(editor.getText()).toBe('BEFORE ')
    expect(editor.commands.undo()).toBe(false)
  },
)

it.each([
  ['doc:0.0', 'unbound'],
  ['doc:not-an-ordinal', 'unbound'],
  ['media:malformed', 'malformed'],
  ['https://example.test/image.png', 'unsupported'],
  ['data:image/png;base64,AQ==', 'unsupported'],
  ['', 'unsupported'],
])('refuses unsupported explicit image src %s', async (reference, reason) => {
  const { caller } = fixture([imageParagraph()])
  const before = caller.context()
  expect(
    await caller.execute({
      id: 'invalid-src',
      name: 'replace_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0, html: `<p>WRONG<img src="${reference}"></p>` },
    }),
  ).toMatchObject({
    isError: true,
    committed: false,
    revision: before.revision,
    imageReferenceErrors: [{ reference, reason }],
  })
})

it('refuses references issued for another document', async () => {
  const first = fixture([imageParagraph()]).caller.context()
  const { caller } = fixture([imageParagraph('Ag==')])
  const current = caller.context()
  expect(
    await caller.execute({
      id: 'other-document',
      name: 'insert_content',
      input: { html: `<p><img src="media:${first.documentId}:${first.revision}:0.0"></p>` },
    }),
  ).toMatchObject({
    isError: true,
    committed: false,
    errorCode: 'stale_target',
    revision: current.revision,
  })
})

it('refuses a missing block in a current bound reference', async () => {
  const { caller } = fixture([imageParagraph()])
  const before = caller.context()
  expect(
    await caller.execute({
      id: 'missing-block',
      name: 'insert_content',
      input: { html: `<p><img src="media:${before.documentId}:${before.revision}:99999.0"></p>` },
    }),
  ).toMatchObject({
    isError: true,
    committed: false,
    revision: before.revision,
    imageReferenceErrors: [{ reason: 'missing' }],
  })
})

it('validates the public input schema before attempting an image write', async () => {
  const { caller } = fixture([imageParagraph()])
  const before = caller.context()
  expect(
    await caller.execute({
      id: 'invalid-schema',
      name: 'replace_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0, html: 123 },
    }),
  ).toMatchObject({
    isError: true,
    committed: false,
    errorCode: 'invalid_input',
    revision: before.revision,
  })
})

it('keeps bound HTML references intact across character-page continuation', async () => {
  const content = imageParagraph()
  content.content[0].text = 'a'.repeat(23990)
  const { caller } = fixture([content])
  const before = caller.context()
  let html = ''
  let offset = 0
  for (;;) {
    const read = await caller.execute({
      id: `page-${offset}`,
      name: 'read_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0, offset },
    })
    expect(read.isError, read.output).not.toBe(true)
    html += read.output
      .split('\n…')[0]
      .split('\n(end of range:')[0]
      .split('\nAuthorized document images')[0]
    const next = /offset=(\d+)/.exec(read.output)
    if (!next) break
    offset = Number(next[1])
  }
  expect(offset).toBeGreaterThan(0)
  expect(html).toBe(
    `<p>${'a'.repeat(23990)}<img src="media:${before.documentId}:${before.revision}:0.0"></p>`,
  )
  expect(
    await caller.execute({
      id: 'paged-write',
      name: 'replace_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0, html },
    }),
  ).toMatchObject({ committed: true })
})

it.each(['redaction', 'del'])('does not authorize an image hidden by %s', async (mark) => {
  const content = imageParagraph()
  Object.assign(content.content[1], {
    marks: [{ type: mark, attrs: { label: 'private', author: 'test' } }],
  })
  const { caller } = fixture([content])
  const before = caller.context()
  expect(before.text).not.toContain(`media:${before.documentId}:${before.revision}:0.0`)
  expect(
    await caller.execute({
      id: 'hidden',
      name: 'replace_blocks',
      input: {
        startBlockIndex: 0,
        endBlockIndex: 0,
        html: `<p>GUESSED<img src="media:${before.documentId}:${before.revision}:0.0"></p>`,
      },
    }),
  ).toMatchObject({ isError: true, committed: false, revision: before.revision })
})

it('does not present a protected image or OLE preview as a writable inline object', async () => {
  const { caller } = fixture([
    {
      type: 'docProtected',
      attrs: {
        blockType: 'image',
        imageDataUrl: 'data:image/png;base64,AQ==',
        oleProgId: 'Equation.3',
      },
    },
  ])
  const before = caller.context()
  expect(
    await caller.execute({
      id: 'object-preview',
      name: 'replace_blocks',
      input: {
        startBlockIndex: 0,
        endBlockIndex: 0,
        html: `<p><img src="media:${before.documentId}:${before.revision}:0.0"></p>`,
      },
    }),
  ).toMatchObject({
    isError: true,
    committed: false,
    imageReferenceErrors: [{ reason: 'unsupported' }],
  })
})

it('returns structured reference refusal through standard MCP', async () => {
  const { runtime } = fixture([imageParagraph()])
  const router = createDocsCommandRouter(runtime)
  const server = createDocsMcpServer({
    version: 'test',
    endpoint: {
      request: (message, signal) => router.handle(message, signal),
      closeCaller: (id) => router.disconnect(id),
    },
  })
  const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair()
  await server.connect(serverTransport)
  const client = new Client({ name: 'reference-regression', version: '1' }, { capabilities: {} })
  try {
    await client.connect(clientTransport)
    const read = await client.callTool({ name: 'read_document', arguments: {} })
    const before = read.structuredContent as { revision: number }
    const result = await client.callTool({
      name: 'replace_blocks',
      arguments: {
        startBlockIndex: 0,
        endBlockIndex: 0,
        html: '<p>WRONG<img src="doc:99999.0"></p>',
      },
    })
    expect(result).toMatchObject({
      isError: true,
      structuredContent: {
        committed: false,
        revision: before.revision,
        imageReferenceErrors: [{ reference: 'doc:99999.0', reason: 'unbound' }],
      },
    })
  } finally {
    await client.close()
    await server.close()
  }
})

it('preserves exact active image bytes through a valid write and DOCX save/reopen', async () => {
  const parsed = await parseDocx(
    await buildDocx({
      withImage: true,
      bodyXml: `<w:p><w:r><w:t>BEFORE </w:t></w:r><w:r>${inlinePictureDrawingXml({ rId: 'rId10', docPrId: 1, widthPx: 100, heightPx: 80 })}</w:r></w:p>`,
    }),
  )
  const { editor, caller } = fixture(blocksToPmDoc(parsed.blocks).content!)
  caller.context()
  const read = await caller.execute({
    id: 'read',
    name: 'read_blocks',
    input: { startBlockIndex: 0, endBlockIndex: 0 },
  })
  const reference = /<img src="([^"]+)">/.exec(read.output)![1]
  const beforeImage = await caller.execute({
    id: 'image',
    name: 'read_document_media',
    input: { mediaRef: reference },
  })
  expect(beforeImage.isError, beforeImage.output).not.toBe(true)
  expect(beforeImage.media?.mimeType).toBe('image/png')
  expect(beforeImage.media?.data.length).toBeGreaterThan(0)
  expect(
    await caller.execute({
      id: 'write',
      name: 'replace_blocks',
      input: {
        startBlockIndex: 0,
        endBlockIndex: 0,
        html: `<p>AFTER <img src="${reference}"></p>`,
      },
    }),
  ).toMatchObject({ committed: true })
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const saved = await saveDocx(parsed, plan.saveBlocks)
  const reopened = await parseDocx(saved)
  const fresh = fixture(blocksToPmDoc(reopened.blocks).content!).caller
  const context = fresh.context()
  const afterImage = await fresh.execute({
    id: 'image-after',
    name: 'read_document_media',
    input: { mediaRef: `media:${context.documentId}:${context.revision}:0.0` },
  })
  expect(afterImage.isError, afterImage.output).not.toBe(true)
  expect(afterImage.media?.data.length).toBeGreaterThan(0)
  expect(afterImage.media).toEqual(beforeImage.media)
  expect(context.text).toContain('AFTER')
  const zip = await JSZip.loadAsync(saved)
  expect(await zip.file('word/document.xml')!.async('string')).toContain('r:embed="rId10"')
})
