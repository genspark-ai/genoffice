import { expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  parseDocx,
  readSections,
  saveDocx,
  PAGE_MARK,
  TOTAL_PAGES_MARK,
} from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { createHeaderFooterAccess } from '../src/renderer/ai/header-footer-access'
import { hfSaveOptions, resolveHf, type HfSectionState } from '../src/renderer/hf-sections'

const variants = ['default', 'first', 'even'] as const
const rel = (id: string) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="${id}.xml"/>`
const refs = (id?: string) =>
  id ? variants.map((v) => `<w:footerReference w:type="${v}" r:id="${id}"/>`).join('') : ''
async function fixture() {
  const ids = ['shared', undefined, 'shared', 'own3', 'own4', undefined, 'own6', undefined, 'own8']
  return buildDocx({
    bodyXml:
      ids
        .slice(0, -1)
        .map(
          (id, i) =>
            `<w:p><w:pPr><w:sectPr>${refs(id)}</w:sectPr></w:pPr><w:r><w:t>section ${i}</w:t></w:r></w:p>`,
        )
        .join('') + '<w:p><w:r><w:t>last</w:t></w:r></w:p>',
    sectPrExtra: refs(ids.at(-1)),
    extraRels: [...new Set(ids.filter((x): x is string => !!x))].map(rel).join(''),
    extraParts: [...new Set(ids.filter((x): x is string => !!x))].map((id) => ({
      path: `word/${id}.xml`,
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
      xml: `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>${id}</w:t></w:r></w:p></w:ftr>`,
    })),
  })
}
function state(parsed: Awaited<ReturnType<typeof parseDocx>>): HfSectionState {
  return { sections: readSections(parsed), edits: {}, links: {}, parsed }
}

it.each([0, 1, 4, 8].flatMap((section) => variants.map((view) => ({ section, view }))))(
  'isolates section $section $view across save/reopen with live page fields',
  async ({ section, view }) => {
    const parsed = await parseDocx(await fixture())
    let current = state(parsed)
    const before = current.sections.map((_, i) =>
      variants.map((v) => resolveHf(current, i, 'footer', v).value?.text),
    )
    const access = createHeaderFooterAccess({
      current: () => current,
      commit: (next) => {
        current = next
      },
      locked: () => false,
      evenOdd: () => false,
      enableFirst: () => null,
      enableEven: () => {},
    })
    expect(access.read().sections).toHaveLength(9)
    expect(access.set('footer', view, 'Changed {PAGE} / {NUMPAGES}', section)).toBeNull()
    current.sections.forEach((_, i) => {
      if (i !== section)
        expect(variants.map((v) => resolveHf(current, i, 'footer', v).value?.text)).toEqual(
          before[i],
        )
    })
    const saved = await saveDocx(
      parsed,
      parsed.blocks
        .filter((b) => !b.hidden)
        .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
      hfSaveOptions(current),
    )
    const reopened = await parseDocx(saved)
    const after = state(reopened)
    after.sections.forEach((_, i) =>
      variants.forEach((v, vi) => {
        const value = resolveHf(after, i, 'footer', v).value
        if (i === section && v === view) expect(value?.text).toContain('Changed')
        else expect(value?.text).toBe(before[i][vi])
      }),
    )
    const rId = after.sections[section].footerRefs[view]!
    const part = reopened.hfParts![rId]
    expect(part.hasPageNumber).toBe(true)
    expect(part.text).toContain(PAGE_MARK)
    expect(part.text).toContain(TOTAL_PAGES_MARK)
    const zip = await JSZip.loadAsync(saved)
    expect(Object.keys(zip.files).some((p) => /footer\d+\.xml$/.test(p))).toBe(true)
  },
)

it('coalesces shared aliases and preserves an isolated variant through later shared edits', async () => {
  const parsed = await parseDocx(await fixture())
  let current = state(parsed)
  const access = createHeaderFooterAccess({
    current: () => current,
    commit: (next) => {
      current = next
    },
    locked: () => false,
    evenOdd: () => false,
    enableFirst: () => null,
    enableEven: () => {},
  })
  expect(access.set('footer', 'default', 'A', 0, 'shared')).toBeNull()
  expect(access.set('footer', 'first', 'B', 0, 'shared')).toBeNull()
  for (const view of variants) expect(resolveHf(current, 0, 'footer', view).value?.text).toBe('B')
  expect(access.set('footer', 'first', 'C', 0, 'section')).toBeNull()
  expect(access.set('footer', 'default', 'D', 0, 'shared')).toBeNull()
  const saved = await saveDocx(
    parsed,
    parsed.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    hfSaveOptions(current),
  )
  const reopened = state(await parseDocx(saved))
  for (const view of variants) {
    expect(resolveHf(current, 0, 'footer', view).value?.text).toBe(view === 'first' ? 'C' : 'D')
    expect(resolveHf(reopened, 0, 'footer', view).value?.text).toBe(view === 'first' ? 'C' : 'D')
  }
})

it('targets an inserted unsaved section by its emitted section ordinal', async () => {
  const parsed = await parseDocx(await fixture())
  let current = state(parsed)
  current.sections = [
    { ...current.sections[0], lastBlockIndex: -1, pendingBreak: true },
    ...current.sections,
  ]
  const access = createHeaderFooterAccess({
    current: () => current,
    commit: (next) => {
      current = next
    },
    locked: () => false,
    evenOdd: () => false,
    enableFirst: () => null,
    enableEven: () => {},
  })
  expect(access.set('footer', 'default', 'New section {PAGE}', 0)).toBeNull()
  const blocks = [
    {
      kind: 'xml' as const,
      xml: `<w:p><w:pPr><w:sectPr>${refs('shared')}</w:sectPr></w:pPr></w:p>`,
    },
    ...parsed.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
  ]
  const reopened = state(await parseDocx(await saveDocx(parsed, blocks, hfSaveOptions(current))))
  expect(reopened.sections).toHaveLength(10)
  expect(resolveHf(reopened, 0, 'footer', 'default').value?.text).toContain('New section')
  expect(resolveHf(reopened, 1, 'footer', 'default').value?.text).toBe('shared')
})

it('clones relationships from the effective owner after a pending Link to Previous', async () => {
  const bytes = await fixture()
  const zip = await JSZip.loadAsync(bytes)
  const rels = (target: string) =>
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${target}" TargetMode="External"/></Relationships>`
  zip.file('word/_rels/own3.xml.rels', rels('https://example.com/source-a'))
  zip.file('word/_rels/own4.xml.rels', rels('https://example.com/source-b'))
  const parsed = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
  let current = state(parsed)
  const key = `${current.sections[4].lastBlockIndex}:footer`
  current.links[key] = true
  const access = createHeaderFooterAccess({
    current: () => current,
    commit: (next) => {
      current = next
    },
    locked: () => false,
    evenOdd: () => false,
    enableFirst: () => null,
    enableEven: () => {},
  })
  expect(access.set('footer', 'default', 'Isolated {PAGE}', 4)).toBeNull()
  expect(resolveHf(current, 5, 'footer', 'default').value?.text).toBe('own3')
  const saved = await saveDocx(
    parsed,
    parsed.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    hfSaveOptions(current),
  )
  const savedZip = await JSZip.loadAsync(saved)
  const cloneRels = await savedZip.file('word/_rels/footer1.xml.rels')?.async('string')
  expect(cloneRels).toContain('source-a')
  expect(cloneRels).not.toContain('source-b')
  const reopened = state(await parseDocx(saved))
  expect(resolveHf(reopened, 5, 'footer', 'default').value?.text).toBe('own3')
})

it('retains isolation when the native page-number gallery replaces the strip', async () => {
  const { withHfEdit } = await import('../src/renderer/hf-sections')
  const { hfPageNumberOnly } = await import('../src/renderer/editor/hf-dom')
  const parsed = await parseDocx(await fixture())
  let current = state(parsed)
  const access = createHeaderFooterAccess({
    current: () => current,
    commit: (next) => {
      current = next
    },
    locked: () => false,
    evenOdd: () => false,
    enableFirst: () => null,
    enableEven: () => {},
  })
  access.set('footer', 'default', 'Scoped', 0)
  current = {
    ...current,
    ...withHfEdit(current, 0, 'footer', 'default', hfPageNumberOnly('center'))!,
  }
  expect(resolveHf(current, 1, 'footer', 'default').value?.text).toBe('shared')
  const reopened = state(
    await parseDocx(
      await saveDocx(
        parsed,
        parsed.blocks
          .filter((b) => !b.hidden)
          .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
        hfSaveOptions(current),
      ),
    ),
  )
  expect(resolveHf(reopened, 0, 'footer', 'default').value?.pageNumber).toBe(true)
  expect(resolveHf(reopened, 1, 'footer', 'default').value?.text).toBe('shared')
})

it('uses the section contract through a real MCP client and saves the same target', async () => {
  const { Editor } = await import('@tiptap/core')
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const { createDocsRuntime } = await import('../src/renderer/runtime/docs-runtime')
  const { createDocsCommandRouter } = await import('../src/renderer/runtime/command-router')
  const { createDocsMcpServer } = await import('../src/services/docs-mcp')
  const { InMemoryTransport } = await import('@modelcontextprotocol/sdk/inMemory.js')
  const { Client } = await import('@modelcontextprotocol/sdk/client/index.js')
  const parsed = await parseDocx(await fixture())
  let current = state(parsed)
  let saved: Uint8Array | undefined
  const access = createHeaderFooterAccess({
    current: () => current,
    commit: (next) => {
      current = next
    },
    locked: () => false,
    evenOdd: () => false,
    enableFirst: () => null,
    enableEven: () => {},
  })
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [{ type: 'text', text: 'Body' }] },
        {
          type: 'docTable',
          content: [
            {
              type: 'docTableRow',
              content: [
                {
                  type: 'docTableCell',
                  attrs: { colspan: 2 },
                  content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'Merged' }] }],
                },
              ],
            },
          ],
        },
      ],
    },
  })
  const runtime = createDocsRuntime({
    current: () => ({
      editor,
      identity: editor,
      stateVersion: current,
      headerFooter: access,
      numIds: { bullet: null, ordered: null },
    }),
    save: async ({ path }) => {
      saved = await saveDocx(
        parsed,
        parsed.blocks
          .filter((b) => !b.hidden)
          .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
        hfSaveOptions(current),
      )
      return { ok: true, path, committed: true, pending: false }
    },
  })
  const router = createDocsCommandRouter(runtime)
  const server = createDocsMcpServer({
    version: 'test',
    endpoint: {
      request: (message, signal) => router.handle(message, signal),
      closeCaller: (id) => router.disconnect(id),
    },
  })
  const [hostTransport, clientTransport] = InMemoryTransport.createLinkedPair()
  const client = new Client({ name: 'section-contract', version: '1' }, { capabilities: {} })
  try {
    await server.connect(hostTransport)
    await client.connect(clientTransport)
    const context = await client.callTool({ name: 'get_document_context', arguments: {} })
    expect(JSON.stringify(context)).toContain('section 4')
    const blocks = await client.callTool({
      name: 'read_blocks',
      arguments: { startBlockIndex: 1, endBlockIndex: 1 },
    })
    expect(blocks).toMatchObject({
      structuredContent: { output: expect.stringContaining('colspan="2"') },
    })
    const before = editor.getJSON()
    for (const dryRun of [true, false]) {
      const rejected = await client.callTool({
        name: 'apply_ops',
        arguments: {
          dryRun,
          ops: [
            { op: 'setList', target: { blockIndexes: [0] } },
            { op: 'setList', target: { blockIndexes: [1] } },
          ],
        },
      })
      expect(rejected).toMatchObject({
        isError: true,
        structuredContent: { errorCode: 'invalid_input', committed: false },
      })
      expect(JSON.stringify(rejected)).not.toContain('Invalid content for node')
      expect(editor.getJSON()).toEqual(before)
    }

    const edit = await client.callTool({
      name: 'set_header_footer',
      arguments: { kind: 'footer', section: 4, view: 'default', text: 'MCP {PAGE}' },
    })
    expect(edit.isError, JSON.stringify(edit)).not.toBe(true)
    expect(JSON.stringify(edit)).toContain('affected targets')
    const result = await client.callTool({
      name: 'save_document',
      arguments: { path: 'saved.docx', overwrite: false },
    })
    expect(result.isError, JSON.stringify(result)).not.toBe(true)
    expect(saved).toBeDefined()
    const reopened = state(await parseDocx(saved!))
    expect(resolveHf(reopened, 4, 'footer', 'default').value?.text).toContain('MCP')
    expect(resolveHf(reopened, 5, 'footer', 'default').value?.text).toBe('own4')
    expect(resolveHf(reopened, 0, 'footer', 'default').value?.text).toBe('shared')
  } finally {
    await client.close()
    await server.close()
    router.dispose()
    runtime.dispose()
    editor.destroy()
  }
})
