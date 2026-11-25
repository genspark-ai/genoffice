import { afterEach, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createDocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { createDocsMcpServer } from '../src/services/docs-mcp'
import { createDocsCommandRouter } from '../src/renderer/runtime/command-router'
import { docsToolCatalog } from '../src/shared/tool-catalog'

const cleanup: Array<() => void> = []
afterEach(() =>
  cleanup
    .splice(0)
    .reverse()
    .forEach((dispose) => dispose()),
)

function open(html = '<p>Alpha</p><p>Beta</p>') {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: html,
  })
  const runtime = createDocsRuntime({
    current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
  })
  cleanup.push(
    () => editor.destroy(),
    () => runtime.dispose(),
  )
  const client = runtime.connect({ id: 'query' })
  client.context()
  return { editor, client }
}

it.each([
  { type: 'docInlineMath', attrs: { latex: 'x+y' } },
  {
    type: 'docInlineImage',
    attrs: { dataUrl: 'data:image/png;base64,AA==', widthPx: 20, heightPx: 20 },
  },
  { type: 'text', text: 'FIELD', marks: [{ type: 'instrField', attrs: { instr: 'DATE' } }] },
  {
    type: 'text',
    text: 'INSERTED',
    marks: [{ type: 'ins', attrs: { author: 'Author', id: '1' } }],
  },
  { type: 'text', text: 'DELETED', marks: [{ type: 'del', attrs: { author: 'Author', id: '2' } }] },
])('does not advertise ordinary cell text writes over $type payloads', async (inline) => {
  const { editor, client } = open()
  editor.commands.setContent({
    type: 'doc',
    content: [
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
                      { type: 'text', text: 'Before ' },
                      inline,
                      { type: 'text', text: ' after' },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
      { type: 'docParagraph' },
    ],
  })
  client.context()
  const result = await client.execute({
    id: 'query-mixed-cell',
    name: 'query_document',
    input: { domain: 'cells' },
  })
  const item = JSON.parse(result.output).items[0]
  expect(item.supportedOperations).not.toContain('setTableCellText')
  expect(item.limitations).toContainEqual(
    expect.objectContaining({ operation: 'setTableCellText', code: 'protected_content' }),
  )
  expect(item.projection).toMatchObject({ lossless: false })
  const before = editor.getJSON()
  const refused = await client.execute({
    id: 'write-mixed-cell',
    name: 'apply_ops',
    input: {
      ops: [{ op: 'setTableCellText', target: { engine: item.target }, text: 'REPLACEMENT' }],
    },
  })
  expect(refused).toMatchObject({ isError: true, committed: false, revision: result.revision })
  expect(editor.getJSON()).toEqual(before)
})

it('describes nested cell objects without permitting their implicit replacement', async () => {
  const { editor, client } = open()
  editor.commands.setContent({
    type: 'doc',
    content: [
      {
        type: 'docTable',
        content: [
          {
            type: 'docTableRow',
            content: [
              {
                type: 'docTableCell',
                content: [
                  { type: 'docParagraph', content: [{ type: 'text', text: 'Before' }] },
                  { type: 'docNestedTable', attrs: { model: { rows: [[{ paras: ['Nested'] }]] } } },
                ],
              },
            ],
          },
        ],
      },
      { type: 'docParagraph' },
    ],
  })
  client.context()
  const result = await client.execute({
    id: 'nested-query',
    name: 'query_document',
    input: { domain: 'cells' },
  })
  const item = JSON.parse(result.output).items[0]
  expect(item.projection).toMatchObject({ lossless: false, contentKinds: ['docNestedTable'] })
  expect(item.structures).toContainEqual(
    expect.objectContaining({
      kind: 'docNestedTable',
      id: expect.stringContaining(`structure:${result.documentId}:${result.revision}:`),
    }),
  )
  expect(item.supportedOperations).toEqual([])
  expect(item.complete).toBe(true)
  const before = editor.getJSON()
  expect(
    await client.execute({
      id: 'nested-write',
      name: 'apply_ops',
      input: {
        ops: [
          { op: 'setTableCellText', target: { engine: item.target }, text: 'Lost nested content' },
        ],
      },
    }),
  ).toMatchObject({ isError: true, committed: false, revision: result.revision })
  expect(editor.getJSON()).toEqual(before)
})

it('shares parent content-control protection between cell discovery and execution', async () => {
  const { editor, client } = open('<table><tr><td>Cell</td></tr></table><p>After</p>')
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(0, undefined, {
      ...editor.state.doc.firstChild!.attrs,
      sdtShell: '{}',
    }),
  )
  client.context()
  const result = await client.execute({
    id: 'controlled-query',
    name: 'query_document',
    input: { domain: 'cells' },
  })
  const item = JSON.parse(result.output).items[0]
  expect(item.supportedOperations).toEqual([])
  expect(item.projection.contentKinds).toContain('content-control')
  expect(
    await client.execute({
      id: 'controlled-write',
      name: 'apply_ops',
      input: {
        ops: [{ op: 'setTableCellText', target: { engine: item.target }, text: 'Replacement' }],
      },
    }),
  ).toMatchObject({ isError: true, committed: false, revision: result.revision })
})

it('continues a cell object inventory on the same owner when earlier cells also match', async () => {
  const { editor, client } = open()
  editor.commands.setContent({
    type: 'doc',
    content: [
      {
        type: 'docTable',
        content: [
          {
            type: 'docTableRow',
            content: [
              {
                type: 'docTableCell',
                content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'Plain' }] }],
              },
              {
                type: 'docTableCell',
                content: [
                  {
                    type: 'docParagraph',
                    content: Array.from({ length: 40 }, (_, i) => ({
                      type: 'text',
                      text: `FIELD_${i} `,
                      marks: [
                        { type: 'instrField', attrs: { instr: ` DOCPROPERTY property${i} ` } },
                      ],
                    })),
                  },
                ],
              },
            ],
          },
        ],
      },
      { type: 'docParagraph' },
    ],
  })
  client.context()
  const first = await client.execute({
    id: 'cell-objects',
    name: 'query_document',
    input: { domain: 'cells' },
  })
  const item = JSON.parse(first.output).items[1]
  expect(item.structures).toHaveLength(32)
  expect(item.structuresComplete).toBe(false)
  const next = await client.execute({
    id: 'cell-objects-next',
    name: 'query_document',
    input: { cursor: item.structuresPage.nextCursor },
  })
  const nextItems = JSON.parse(next.output).items
  expect(nextItems).toHaveLength(1)
  expect(nextItems[0].target).toEqual(item.target)
  expect(nextItems[0].structures).toHaveLength(8)
  expect(nextItems[0].structuresPage.complete).toBe(true)
  expect(
    new Set([...item.structures, ...nextItems[0].structures].map((structure) => structure.id)).size,
  ).toBe(40)
})

it('publishes merged table and explicit row capabilities that agree with executable operations', async () => {
  const { client } = open('<p>Anchor</p>')
  const insert = docsToolCatalog().find((tool) => tool.name === 'insert_content')
  const definition = JSON.stringify(insert?.inputSchema)
  expect(definition).toContain('support horizontal/vertical merges')
  expect(definition).not.toContain('merged cells are not supported')
  expect(definition).not.toContain('only cell text remains editable')
  expect(
    await client.execute({
      id: 'merged-insert',
      name: 'insert_content',
      input: {
        afterBlockIndex: 0,
        html: '<table><tr><td colspan="2">Merged</td></tr><tr><td>Left</td><td>Right</td></tr></table>',
      },
    }),
  ).toMatchObject({ committed: true })
  const before = await client.execute({
    id: 'merged-read',
    name: 'query_document',
    input: { domain: 'cells' },
  })
  expect(JSON.parse(before.output).items[0].html).toContain('colspan="2"')
  expect(
    await client.execute({
      id: 'row-delete',
      name: 'apply_ops',
      input: { ops: [{ op: 'deleteTableRow', target: { blockIndexes: [1] }, row: 1 }] },
    }),
  ).toMatchObject({ committed: true })
  const after = await client.execute({
    id: 'row-read',
    name: 'query_document',
    input: { domain: 'cells' },
  })
  expect(JSON.parse(after.output).items).toHaveLength(1)
})

it.each(['insert_content', 'replace_blocks'])(
  'refuses unsupported table formula HTML atomically through %s',
  async (name) => {
    const { editor, client } = open('<p>Anchor</p><p>Keep</p>')
    const before = editor.getJSON()
    const context = client.context()
    const html =
      '<p>Must not be inserted</p><table><tr><td>Before <formula>x+y</formula> after</td></tr></table>'
    const result = await client.execute({
      id: 'unsupported-table-formula',
      name,
      input:
        name === 'insert_content'
          ? { afterBlockIndex: 0, html }
          : { startBlockIndex: 0, endBlockIndex: 1, html },
    })
    expect(result).toMatchObject({
      isError: true,
      committed: false,
      mutated: false,
      revision: context.revision,
    })
    expect(result.output).toMatch(/formula.*table cell.*not supported/i)
    expect(editor.getJSON()).toEqual(before)
    expect(editor.commands.undo()).toBe(false)
  },
)

it('preserves existing native cell equations through a neighboring public text edit and save/reopen', async () => {
  const equation =
    '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x+y</m:t></m:r></m:oMath>'
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml:
        '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2880"/></w:tblGrid><w:tr>' +
        `<w:tc><w:p>${equation}</w:p></w:tc><w:tc><w:p><w:r><w:t>Neighbor</w:t></w:r></w:p></w:tc>` +
        '</w:tr></w:tbl><w:p/>',
    }),
  )
  const { editor, client } = open()
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  client.context()
  const query = await client.execute({
    id: 'native-equation-neighbor',
    name: 'query_document',
    input: { domain: 'cells' },
  })
  const [mathCell, neighbor] = JSON.parse(query.output).items
  expect(mathCell.structures).toContainEqual(expect.objectContaining({ kind: 'math' }))
  expect(neighbor.supportedOperations).toContain('setTableCellText')
  expect(
    await client.execute({
      id: 'native-equation-edit',
      name: 'apply_ops',
      input: {
        ops: [{ op: 'setTableCellText', target: { engine: neighbor.target }, text: 'Updated' }],
      },
    }),
  ).toMatchObject({ committed: true })
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
  const beforeMath = parsed.blocks[0].table?.rows[0][0].richParas
    ?.flatMap((p) => p.runs)
    .find((run) => run.math)?.math
  const afterMath = reopened.blocks[0].table?.rows[0][0].richParas
    ?.flatMap((p) => p.runs)
    .find((run) => run.math)?.math
  expect(beforeMath).toBeDefined()
  expect(afterMath).toEqual(beforeMath)
  expect(reopened.blocks[0].table?.rows[0][1].paras).toEqual(['Updated'])
})

it('returns bounded current targets and coherent continuation', async () => {
  const { client } = open()
  const first = await client.execute({ id: 'q', name: 'query_document', input: { limit: 1 } })
  expect(first.isError).not.toBe(true)
  const page = JSON.parse(first.output)
  expect(page.page).toMatchObject({ complete: false })
  expect(page.items[0]).toMatchObject({
    type: 'docParagraph',
    html: '<p>Alpha</p>',
    complete: true,
    target: {
      documentId: first.documentId,
      revision: first.revision,
      expectedType: 'docParagraph',
    },
  })
  const second = await client.execute({
    id: 'q2',
    name: 'query_document',
    input: { cursor: page.page.nextCursor },
  })
  expect(JSON.parse(second.output)).toMatchObject({
    page: { complete: true },
    items: [{ html: '<p>Beta</p>' }],
  })
})

it('distinguishes an accepted no-op from a mutation without expiring targets', async () => {
  const { client } = open()
  const before = client.context()
  const result = await client.execute({
    id: 'same',
    name: 'apply_ops',
    input: {
      ops: [{ op: 'findReplace', target: { blockIndexes: [0] }, find: 'Alpha', replace: 'Alpha' }],
    },
  })
  expect(result).toMatchObject({
    committed: false,
    mutated: false,
    revision: before.revision,
    receipt: {
      outcome: 'no-op',
      revisionBefore: before.revision,
      revisionAfter: before.revision,
      persistence: 'unchanged',
    },
  })
})

it('uses query targets directly for precise edits and refuses the expired address', async () => {
  const { client } = open()
  const query = await client.execute({
    id: 'query',
    name: 'query_document',
    input: { containsText: 'Beta' },
  })
  const target = JSON.parse(query.output).items[0].target
  const preview = await client.execute({
    id: 'preview',
    name: 'apply_ops',
    input: {
      dryRun: true,
      ops: [{ op: 'findReplace', target: { engine: target }, find: 'Beta', replace: 'Gamma' }],
    },
  })
  expect(preview).toMatchObject({
    committed: false,
    mutated: false,
    matchCount: 1,
    changeCount: 1,
    revision: query.revision,
    receipt: { outcome: 'preview' },
  })
  const edit = await client.execute({
    id: 'edit',
    name: 'apply_ops',
    input: {
      ops: [{ op: 'findReplace', target: { engine: target }, find: 'Beta', replace: 'Gamma' }],
    },
  })
  expect(edit).toMatchObject({
    committed: true,
    matchCount: 1,
    changeCount: 1,
    receipt: { outcome: 'applied', persistence: 'pending', resolvedTargets: [target] },
  })
  const stale = await client.execute({
    id: 'again',
    name: 'apply_ops',
    input: {
      ops: [{ op: 'findReplace', target: { engine: target }, find: 'Gamma', replace: 'Delta' }],
    },
  })
  expect(stale).toMatchObject({
    isError: true,
    committed: false,
    errorCode: 'stale_target',
    receipt: { outcome: 'rejected' },
  })
  expect(client.context().text).toContain('Gamma')
})

it('maps original targets through preceding deletes and refuses targets deleted earlier in an atomic batch', async () => {
  const { client } = open('<p>First</p><p>Second</p><p>Third</p>')
  const result = await client.execute({ id: 'query', name: 'query_document', input: {} })
  const [first, second, third] = JSON.parse(result.output).items
  const refused = await client.execute({
    id: 'refused',
    name: 'apply_ops',
    input: {
      ops: [
        { op: 'deleteBlocks', target: { engine: first.target } },
        { op: 'findReplace', target: { engine: first.target }, find: 'Second', replace: 'Wrong' },
      ],
    },
  })
  expect(refused).toMatchObject({ isError: true, committed: false, revision: result.revision })
  const applied = await client.execute({
    id: 'mapped',
    name: 'apply_ops',
    input: {
      ops: [
        { op: 'deleteBlocks', target: { engine: first.target } },
        { op: 'findReplace', target: { engine: third.target }, find: 'Third', replace: 'Final' },
      ],
    },
  })
  expect(applied.isError).not.toBe(true)
  const context = client.context().text
  expect(context).toContain('Second')
  expect(context).toContain('Final')
  expect(context).not.toContain('First')
  expect(second.target.range.from).toBeGreaterThan(0)
})

it.each(['type', 'content', 'document'])(
  'rejects a mismatched %s precondition before an atomic edit',
  async (mismatch) => {
    const { client } = open()
    const q = await client.execute({ id: 'q', name: 'query_document', input: { limit: 1 } })
    const target = JSON.parse(q.output).items[0].target
    if (mismatch === 'type') target.expectedType = 'docTable'
    if (mismatch === 'content') target.expectedHtml = '<p>Other</p>'
    if (mismatch === 'document') target.documentId = 'another-document'
    const result = await client.execute({
      id: 'edit',
      name: 'apply_ops',
      input: {
        ops: [{ op: 'findReplace', target: { engine: target }, find: 'Alpha', replace: 'Changed' }],
      },
    })
    expect(result).toMatchObject({
      isError: true,
      committed: false,
      mutated: false,
      revision: q.revision,
    })
    expect(client.context().text).toContain('Alpha')
  },
)

it('counts actual detached header effects and preserves separate side-state freshness', async () => {
  const { editor } = open()
  let header = 'Original'
  const headerFooter = {
    read: () => ({
      header,
      footer: '',
      headerFirst: null,
      footerFirst: null,
      headerEven: null,
      footerEven: null,
      titlePg: false,
      evenOddHf: false,
      multiSection: false,
    }),
    set: (_kind: 'header' | 'footer', _view: 'default' | 'first' | 'even', text: string) => {
      header = text
      return null
    },
  }
  const runtime = createDocsRuntime({
    current: () => ({
      editor,
      identity: editor,
      headerFooter,
      numIds: { bullet: null, ordered: null },
    }),
  })
  cleanup.push(() => runtime.dispose())
  const a = runtime.connect({ id: 'a' })
  const b = runtime.connect({ id: 'b' })
  editor.commands.setTextSelection(1)
  editor.commands.insertContent('a')
  const before = a.context()
  b.context()
  const same = await a.execute({
    id: 'same',
    name: 'set_header_footer',
    input: { kind: 'header', text: 'Original' },
  })
  expect(same).toMatchObject({
    committed: false,
    mutated: false,
    revision: before.revision,
    receipt: { outcome: 'no-op' },
  })
  editor.commands.insertContent('b')
  a.context()
  await a.execute({ id: 'undo', name: 'document_history', input: { action: 'undo' } })
  expect(editor.getText()).toBe('Alpha\n\nBeta')
  a.context()
  const changed = await a.execute({
    id: 'change',
    name: 'set_header_footer',
    input: { kind: 'header', text: 'Updated' },
  })
  expect(changed.committed).toBe(true)
  expect(changed.revision).toBeGreaterThan(before.revision)
  await b.execute({ id: 'partial', name: 'query_document', input: {} })
  const stale = await b.execute({
    id: 'stale',
    name: 'set_header_footer',
    input: { kind: 'header', text: 'Lost update' },
  })
  expect(stale).toMatchObject({ committed: false, errorCode: 'stale_target' })
  expect(stale.output).toContain('get_document_context')
  expect(header).toBe('Updated')
  const headers = await a.execute({
    id: 'headers',
    name: 'query_document',
    input: { domain: 'headers' },
  })
  const item = JSON.parse(headers.output).items[0]
  expect(item).toMatchObject({ type: 'header', supportedOperations: ['set_header_footer'] })
  expect(
    await a.execute({
      id: 'header-target',
      name: 'set_header_footer',
      input: { engineTarget: item.target, text: 'Targeted' },
    }),
  ).toMatchObject({ committed: true })
  expect(header).toBe('Targeted')
})

it('reports an effect committed before a host publication error instead of inviting a replay', async () => {
  const { editor } = open()
  const runtime = createDocsRuntime({
    current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    apply: (operation) => {
      operation()
      throw new Error('Host publication failed after commit')
    },
  })
  cleanup.push(() => runtime.dispose())
  const caller = runtime.connect({ id: 'failed-publication' })
  const before = caller.context()
  const result = await caller.execute({
    id: 'write',
    name: 'apply_ops',
    input: {
      ops: [
        { op: 'findReplace', target: { blockIndexes: [0] }, find: 'Alpha', replace: 'Committed' },
      ],
    },
  })
  expect(result).toMatchObject({
    isError: true,
    committed: true,
    mutated: true,
    receipt: { outcome: 'failure-after-commit', persistence: 'pending' },
  })
  expect(result.revision).toBeGreaterThan(before.revision)
  expect(caller.context().text).toContain('Committed')
})

it('queries a logical cell and edits it without manual table coordinates', async () => {
  const { client } = open(
    '<table><tr><td><p>Left</p></td><td><p>Right</p></td></tr></table><p>After</p>',
  )
  const q = await client.execute({
    id: 'cells',
    name: 'query_document',
    input: { domain: 'cells', containsText: 'Right' },
  })
  expect(q.isError).not.toBe(true)
  const item = JSON.parse(q.output).items[0]
  expect(item).toMatchObject({ type: 'docTableCell', supportedOperations: ['setTableCellText'] })
  const incompatible = await client.execute({
    id: 'wrong-kind',
    name: 'apply_ops',
    input: { ops: [{ op: 'deleteBlocks', target: { engine: item.target } }] },
  })
  expect(incompatible).toMatchObject({ isError: true, committed: false, revision: q.revision })
  item.target.expectedHtml = item.html
  const edited = await client.execute({
    id: 'cell',
    name: 'apply_ops',
    input: { ops: [{ op: 'setTableCellText', target: { engine: item.target }, text: 'Filled' }] },
  })
  expect(edited).toMatchObject({ committed: true, changeCount: 1 })
  const read = await client.execute({
    id: 'read',
    name: 'query_document',
    input: { domain: 'cells' },
  })
  expect(
    JSON.parse(read.output)
      .items.map((cell: { html: string }) => cell.html)
      .join(''),
  ).toContain('Left')
  expect(read.output).toContain('Filled')
  expect(read.output).not.toContain('Right')
})

it('keeps private and deleted text out of both query filtering and projections', async () => {
  const { editor, client } = open()
  editor.commands.setContent({
    type: 'doc',
    content: [
      {
        type: 'docParagraph',
        content: [
          { type: 'text', text: 'VISIBLE ' },
          {
            type: 'text',
            text: 'PRIVATE_VALUE',
            marks: [{ type: 'redaction', attrs: { label: 'secret' } }],
          },
          {
            type: 'text',
            text: 'DELETED_VALUE',
            marks: [{ type: 'del', attrs: { author: 'test' } }],
          },
        ],
      },
    ],
  })
  client.context()
  for (const containsText of ['PRIVATE_VALUE', 'DELETED_VALUE']) {
    const result = await client.execute({
      id: 'secret',
      name: 'query_document',
      input: { containsText },
    })
    expect(JSON.parse(result.output).items).toEqual([])
    expect(result.output).not.toContain(containsText)
  }
  const all = await client.execute({ id: 'all', name: 'query_document', input: {} })
  expect(all.output).toContain('VISIBLE')
  expect(all.output).not.toContain('PRIVATE_VALUE')
  expect(all.output).not.toContain('DELETED_VALUE')
})

it('requires an authoritative epoch for stores whose getters omit authored data', async () => {
  const { editor } = open()
  let watermark = 'Original'
  let version = 0
  let publishEpoch = false
  const runtime = createDocsRuntime({
    current: () => ({
      editor,
      identity: editor,
      stateVersion: publishEpoch ? version : undefined,
      numIds: { bullet: null, ordered: null },
      extras: {
        watermark: {
          current: () => watermark,
          set: (spec) => {
            const next = spec && 'text' in spec ? spec.text : ''
            if (next !== watermark) version++
            watermark = next
            return null
          },
        },
      },
    }),
  })
  cleanup.push(() => runtime.dispose())
  const client = runtime.connect({ id: 'watermark' })
  client.context()
  const refused = await client.execute({
    id: 'unobservable',
    name: 'set_watermark',
    input: { text: 'Changed' },
  })
  expect(refused).toMatchObject({ committed: false, errorCode: 'unavailable' })
  expect(watermark).toBe('Original')
  expect(client.availableTools()).not.toContain('set_watermark')
  publishEpoch = true
  client.context()
  const changed = await client.execute({
    id: 'observable',
    name: 'set_watermark',
    input: { text: 'Changed' },
  })
  expect(changed.committed).toBe(true)
  expect(watermark).toBe('Changed')
  const same = await client.execute({
    id: 'same',
    name: 'set_watermark',
    input: { text: 'Changed' },
  })
  expect(same).toMatchObject({
    committed: false,
    revision: changed.revision,
    receipt: { outcome: 'no-op' },
  })
})

it('does not attribute concurrent native typing to a no-op request', async () => {
  const { editor, client } = open()
  const pending = client.execute({
    id: 'same',
    name: 'apply_ops',
    input: {
      ops: [{ op: 'findReplace', target: { blockIndexes: [0] }, find: 'Alpha', replace: 'Alpha' }],
    },
  })
  await Promise.resolve()
  editor.commands.insertContent('USER')
  const result = await pending
  expect(result).toMatchObject({
    committed: false,
    mutated: false,
    receipt: { outcome: 'no-op', persistence: 'unchanged' },
  })
  expect(client.context().text).toContain('USER')
})

it('discovers exact structural intent and validates it after earlier batch mapping', async () => {
  const { client } = open('<p>First</p><table><tr><td><p>Cell</p></td></tr></table><p>Last</p>')
  const q = await client.execute({ id: 'q', name: 'query_document', input: {} })
  const [first, table] = JSON.parse(q.output).items
  expect(table.structuresComplete).toBe(true)
  expect(table.structures).toEqual([
    expect.objectContaining({
      kind: 'table',
      id: expect.stringContaining(`structure:${q.documentId}:${q.revision}:table:`),
    }),
  ])
  const refused = await client.execute({
    id: 'refused',
    name: 'apply_ops',
    input: { ops: [{ op: 'deleteBlocks', target: { engine: table.target } }] },
  })
  expect(refused).toMatchObject({ isError: true, committed: false, revision: q.revision })
  const removed = await client.execute({
    id: 'remove',
    name: 'apply_ops',
    input: {
      ops: [
        { op: 'deleteBlocks', target: { engine: first.target } },
        {
          op: 'deleteBlocks',
          target: { engine: table.target },
          deleteStructures: table.structures.map((item: { id: string }) => item.id),
        },
      ],
    },
  })
  expect(removed).toMatchObject({ committed: true, receipt: { outcome: 'applied' } })
  expect(client.context().text).not.toContain('Cell')
  expect(client.context().text).toContain('Last')
})

it('refuses an engine-target batch after a full-document move before committing either operation', async () => {
  const { client } = open('<p>First</p><p>Second</p><p>Third</p>')
  const q = await client.execute({ id: 'q', name: 'query_document', input: {} })
  const target = JSON.parse(q.output).items[1].target
  const result = await client.execute({
    id: 'move',
    name: 'apply_ops',
    input: {
      ops: [
        { op: 'moveBlocks', blockIndexes: [0], afterBlockIndex: 2 },
        { op: 'findReplace', target: { engine: target }, find: 'Second', replace: 'Changed' },
      ],
    },
  })
  expect(result).toMatchObject({
    committed: false,
    mutated: false,
    isError: true,
    revision: q.revision,
  })
  expect(result.output).toMatch(/target.*removed|target.*resolve/i)
  const after = await client.execute({ id: 'after', name: 'query_document', input: {} })
  expect(JSON.parse(after.output).items.map((item: { html: string }) => item.html)).toEqual([
    '<p>First</p>',
    '<p>Second</p>',
    '<p>Third</p>',
  ])
})

it('continues a large individual structure inventory coherently before exact deletion', async () => {
  const { editor, client } = open()
  editor.commands.setContent({
    type: 'doc',
    content: [
      {
        type: 'docParagraph',
        content: Array.from({ length: 40 }, (_, i) => ({
          type: 'text',
          text: `FIELD_${i} `,
          marks: [{ type: 'instrField', attrs: { instr: ` DOCPROPERTY property${i} ` } }],
        })),
      },
    ],
  })
  client.context()
  const q = await client.execute({ id: 'q', name: 'query_document', input: { limit: 1 } })
  let item = JSON.parse(q.output).items[0]
  expect(item.structuresComplete).toBe(false)
  const ids: string[] = item.structures.map((structure: { id: string }) => structure.id)
  expect(item.structuresPage.nextCursor).toEqual(expect.any(String))
  const firstCursor = item.structuresPage.nextCursor
  while (item.structuresPage.nextCursor) {
    const next = await client.execute({
      id: 'next',
      name: 'query_document',
      input: { cursor: item.structuresPage.nextCursor },
    })
    expect(next.revision).toBe(q.revision)
    item = JSON.parse(next.output).items[0]
    ids.push(...item.structures.map((structure: { id: string }) => structure.id))
  }
  expect(item.structuresPage.complete).toBe(true)
  expect(new Set(ids).size).toBe(40)
  expect(
    await client.execute({
      id: 'remove',
      name: 'apply_ops',
      input: {
        ops: [{ op: 'deleteBlocks', target: { engine: item.target }, deleteStructures: ids }],
      },
    }),
  ).toMatchObject({ committed: true })
  expect(
    await client.execute({
      id: 'stale-inventory',
      name: 'query_document',
      input: { cursor: firstCursor },
    }),
  ).toMatchObject({ isError: true, errorCode: 'stale_target' })
})

it('discovers standalone images as readable media without exposing asset URLs or write capabilities', async () => {
  const { editor, client } = open()
  editor.commands.setContent({
    type: 'doc',
    content: [
      {
        type: 'docProtected',
        attrs: { blockType: 'image', imageDataUrl: 'asset://private-host/image.png' },
      },
    ],
  })
  client.context()
  const result = await client.execute({
    id: 'images',
    name: 'query_document',
    input: { domain: 'images' },
  })
  const item = JSON.parse(result.output).items[0]
  expect(item).toMatchObject({
    type: 'image',
    target: { kind: 'image' },
    carrierTarget: {
      kind: 'block',
      documentId: result.documentId,
      revision: result.revision,
      expectedType: 'docProtected',
    },
    mediaRef: `media:${result.documentId}:${result.revision}:0.0`,
    supportedOperations: ['read_document_media'],
  })
  expect(result.output).not.toContain('asset:')
  const refused = await client.execute({
    id: 'wrong-kind',
    name: 'apply_ops',
    input: { ops: [{ op: 'deleteBlocks', target: { engine: item.target } }] },
  })
  expect(refused).toMatchObject({ isError: true, committed: false, revision: result.revision })
})

it('refuses contradictory and bulk-widened ranges on a queried cell target', async () => {
  const { client } = open('<table><tr><td><p>Left</p></td><td><p>Right</p></td></tr></table>')
  const q = await client.execute({
    id: 'q',
    name: 'query_document',
    input: { domain: 'cells', containsText: 'Left' },
  })
  const target = JSON.parse(q.output).items[0].target
  for (const range of [
    { row: 0, col: 1 },
    { row: 0, col: 0, colEnd: 1 },
  ]) {
    expect(
      await client.execute({
        id: 'wrong',
        name: 'apply_ops',
        input: {
          ops: [
            {
              op: 'setTableCellText',
              target: { engine: target },
              range,
              match: 'all',
              text: 'WRONG',
            },
          ],
        },
      }),
    ).toMatchObject({ isError: true, committed: false, revision: q.revision })
  }
  expect(client.context().text).toContain('Left')
  expect(client.context().text).toContain('Right')
})

it.each(['column', 'row'])(
  'keeps a queried physical cell when a preceding %s insertion shifts its grid coordinates',
  async (axis) => {
    const { client } = open('<table><tr><td><p>Left</p></td><td><p>Right</p></td></tr></table>')
    const q = await client.execute({
      id: 'q',
      name: 'query_document',
      input: { domain: 'cells', containsText: 'Right' },
    })
    const target = JSON.parse(q.output).items[0].target
    const result = await client.execute({
      id: 'shift',
      name: 'apply_ops',
      input: {
        ops: [
          {
            op: axis === 'column' ? 'insertTableColumn' : 'insertTableRow',
            target: { blockIndexes: [0] },
            ...(axis === 'column' ? { col: 0 } : { row: 0 }),
            position: 'before',
          },
          { op: 'setTableCellText', target: { engine: target }, text: 'Updated' },
        ],
      },
    })
    expect(result.isError, result.output).not.toBe(true)
    const read = await client.execute({
      id: 'read',
      name: 'query_document',
      input: { domain: 'cells' },
    })
    expect(JSON.parse(read.output).items.map((item: { html: string }) => item.html)).toEqual([
      ...(axis === 'row' ? [expect.not.stringContaining('Updated')] : []),
      expect.not.stringContaining('Updated'),
      expect.stringContaining('Left'),
      expect.stringContaining('Updated'),
    ])
  },
)

it('projects live cell coordinates after deleted rows and writes the live merged owner', async () => {
  const { editor, client } = open()
  editor.commands.setContent({
    type: 'doc',
    content: [
      {
        type: 'docTable',
        content: [
          {
            type: 'docTableRow',
            attrs: { rowRevision: { kind: 'del', author: 'test' } },
            content: [
              {
                type: 'docTableCell',
                attrs: { colspan: 2 },
                content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'DELETED' }] }],
              },
            ],
          },
          {
            type: 'docTableRow',
            content: [
              {
                type: 'docTableCell',
                attrs: { colspan: 2 },
                content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'LIVE' }] }],
              },
            ],
          },
        ],
      },
    ],
  })
  client.context()
  const q = await client.execute({
    id: 'q',
    name: 'query_document',
    input: { domain: 'cells', containsText: 'LIVE' },
  })
  const item = JSON.parse(q.output).items[0]
  expect(item.target.cell).toEqual({ row: 1, col: 0 })
  item.target.expectedHtml = item.html
  expect(
    await client.execute({
      id: 'write',
      name: 'apply_ops',
      input: {
        ops: [
          {
            op: 'setTableCellText',
            target: { engine: item.target },
            range: { row: 1, col: 1 },
            text: 'Updated',
          },
        ],
      },
    }),
  ).toMatchObject({ committed: true })
  expect(editor.state.doc.textContent).toBe('DELETEDUpdated')
})

it('rejects an atomic batch when a preceding column deletion removes the queried owner', async () => {
  const { client } = open(
    '<table><tr><td><p>Left</p></td><td><p>Middle</p></td><td><p>Right</p></td></tr></table>',
  )
  const query = await client.execute({
    id: 'q',
    name: 'query_document',
    input: { domain: 'cells', containsText: 'Middle' },
  })
  const target = JSON.parse(query.output).items[0].target
  const result = await client.execute({
    id: 'removed',
    name: 'apply_ops',
    input: {
      ops: [
        { op: 'deleteTableColumn', target: { blockIndexes: [0] }, col: 1 },
        { op: 'setTableCellText', target: { engine: target }, text: 'WRONG' },
      ],
    },
  })
  expect(result).toMatchObject({
    isError: true,
    committed: false,
    mutated: false,
    revision: query.revision,
  })
  const after = await client.execute({
    id: 'after',
    name: 'query_document',
    input: { domain: 'cells' },
  })
  expect(JSON.parse(after.output).items.map((item: { html: string }) => item.html)).toEqual([
    expect.stringContaining('Left'),
    expect.stringContaining('Middle'),
    expect.stringContaining('Right'),
  ])
})

it('publishes a completed-save receipt and reopens the exact query-targeted effect', async () => {
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Alpha</w:t></w:r></w:p><w:p><w:r><w:t>KEEP</w:t></w:r></w:p>',
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
  const client = runtime.connect({ id: 'save' })
  client.context()
  const q = await client.execute({
    id: 'q',
    name: 'query_document',
    input: { containsText: 'Alpha' },
  })
  await client.execute({
    id: 'w',
    name: 'apply_ops',
    input: {
      ops: [
        {
          op: 'findReplace',
          target: { engine: JSON.parse(q.output).items[0].target },
          find: 'Alpha',
          replace: 'Saved',
        },
      ],
    },
  })
  const result = await client.save({ path: '/query.docx', overwrite: true })
  expect(result).toMatchObject({
    ok: true,
    committed: true,
    pending: false,
    receipt: { persistence: 'completed' },
  })
  const immediate = await client.execute({
    id: 'immediate',
    name: 'query_document',
    input: { containsText: 'Saved' },
  })
  expect(immediate.isError).not.toBe(true)
  expect(JSON.parse(immediate.output).items[0].html).toContain('<strong>Saved</strong>')
  const reopened = await parseDocx(saved!)
  expect(reopened.blocks[0].runs?.map((run) => run.text).join('')).toBe('Saved')
  expect(reopened.blocks[0].runs?.[0].bold).toBe(true)
  expect(reopened.blocks[1].runs?.map((run) => run.text).join('')).toBe('KEEP')
})

it('advertises only granted operations through read-only MCP queries and capabilities', async () => {
  const { editor } = open()
  const runtime = createDocsRuntime({
    current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
  })
  const router = createDocsCommandRouter(runtime)
  const server = createDocsMcpServer({
    version: 'test',
    readOnly: true,
    endpoint: {
      request: (message, signal) => router.handle(message, signal),
      closeCaller: (id) => router.disconnect(id),
    },
  })
  const client = new Client({ name: 'readonly', version: 'test' })
  const [a, b] = InMemoryTransport.createLinkedPair()
  await server.connect(a)
  await client.connect(b)
  try {
    const tools = await client.listTools()
    expect(tools.tools.map((tool) => tool.name)).toContain('query_document')
    expect(tools.tools.map((tool) => tool.name)).not.toContain('apply_ops')
    const result = await client.callTool({ name: 'query_document', arguments: {} })
    const operation = result.structuredContent as { output: string }
    expect(JSON.parse(operation.output).items[0].supportedOperations).toEqual([])
    const context = await client.callTool({ name: 'get_document_context', arguments: {} })
    expect(context.structuredContent).toMatchObject({ editable: false })
    const draftControls = [
      'start_document_draft',
      'inspect_document_draft',
      'update_document_draft',
      'stop_document_draft',
      'finish_document_draft',
    ]
    for (const name of draftControls) {
      expect(tools.tools.some((tool) => tool.name === name)).toBe(false)
      expect(
        await client.callTool({ name, arguments: { draftId: 'cached', decision: 'discard' } }),
      ).toMatchObject({ isError: true, content: [{ text: expect.stringContaining('not found') }] })
    }
    expect(await client.callTool({ name: 'apply_ops', arguments: { ops: [] } })).toMatchObject({
      isError: true,
      content: [{ text: expect.stringContaining('not found') }],
    })
  } finally {
    await client.close()
    await server.close()
    router.dispose()
    runtime.dispose()
  }
})

it('bounds output, marks item truncation and rejects changed or stale continuation', async () => {
  const { client } = open(
    Array.from({ length: 20 }, (_, i) => `<p>${i} ${'x'.repeat(i === 19 ? 10 : 3000)}</p>`).join(
      '',
    ),
  )
  const first = await client.execute({
    id: 'q',
    name: 'query_document',
    input: { limit: 50, maxItemChars: 4000 },
  })
  expect(first.output.length).toBeLessThan(24000)
  const page = JSON.parse(first.output)
  expect(page.page.complete).toBe(false)
  const indexes: number[] = page.items.map((item: { blockIndex: number }) => item.blockIndex)
  let next = page.page.nextCursor
  while (next) {
    const result = await client.execute({
      id: 'continue',
      name: 'query_document',
      input: { cursor: next },
    })
    const continued = JSON.parse(result.output)
    indexes.push(...continued.items.map((item: { blockIndex: number }) => item.blockIndex))
    next = continued.page.nextCursor
  }
  expect(indexes).toEqual(Array.from({ length: 20 }, (_, i) => i))
  const changed = await client.execute({
    id: 'changed',
    name: 'query_document',
    input: { cursor: page.page.nextCursor, containsText: 'elsewhere' },
  })
  expect(changed.isError).toBe(true)
  const short = await client.execute({
    id: 'short',
    name: 'query_document',
    input: { limit: 1, maxItemChars: 2 },
  })
  expect(JSON.parse(short.output).items[0].complete).toBe(false)
  await client.execute({
    id: 'edit',
    name: 'apply_ops',
    input: {
      ops: [{ op: 'findReplace', target: { blockIndexes: [0] }, find: '0 ', replace: 'Changed ' }],
    },
  })
  const stale = await client.execute({
    id: 'stale',
    name: 'query_document',
    input: { cursor: page.page.nextCursor },
  })
  expect(stale).toMatchObject({ isError: true, errorCode: 'stale_target' })
})
