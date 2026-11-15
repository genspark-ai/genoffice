import { afterEach, expect, it } from 'vitest'
import { Editor, type JSONContent } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createDocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { executeTool, markDocSeen } from '../src/renderer/ai/tools'

const cleanup: Array<() => void> = []
afterEach(() =>
  cleanup
    .splice(0)
    .reverse()
    .forEach((dispose) => dispose()),
)

const mathTable: JSONContent = {
  type: 'docTable',
  content: [
    {
      type: 'docTableRow',
      content: [
        {
          type: 'docTableCell',
          content: [
            { type: 'docParagraph', content: [{ type: 'docInlineMath', attrs: { latex: 'x+y' } }] },
          ],
        },
      ],
    },
  ],
}
function open(html: string | JSONContent = '<h2>Heading</h2><p>Body</p><h2>Next</h2>') {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: html,
  })
  const runtime = createDocsRuntime({
    current: () => ({
      editor,
      identity: editor,
      numIds: { bullet: null, ordered: null },
      autoAcceptUntracked: true,
    }),
  })
  cleanup.push(
    () => editor.destroy(),
    () => runtime.dispose(),
  )
  const client = runtime.connect({ id: 'preflight' })
  client.context()
  const query = async () =>
    JSON.parse(
      (await client.execute({ id: crypto.randomUUID(), name: 'query_document', input: {} })).output,
    )
  const replace = (input: Record<string, unknown>) =>
    client.execute({
      id: crypto.randomUUID(),
      name: 'replace_blocks',
      input: { startBlockIndex: 1, endBlockIndex: 1, html: '<p>Updated</p>', ...input },
    })
  return { editor, client, query, replace }
}

it('rejects a range containing a heading when the caller allows body paragraphs only', async () => {
  const { editor, client, replace } = open()
  const before = client.context()
  const result = await replace({
    endBlockIndex: 2,
    preconditions: { allowedTypes: ['docParagraph'] },
  })
  expect(result).toMatchObject({ isError: true, committed: false, revision: before.revision })
  expect(editor.state.doc.textContent).toBe('HeadingBodyNext')
  expect(editor.commands.undo()).toBe(false)
})

it('never silently ignores runtime-only guard or preview parameters on the native tool entry point', async () => {
  const { editor } = open()
  markDocSeen(editor)
  const before = editor.getJSON()
  for (const input of [{ dryRun: true }, { preconditions: { allowedTypes: ['docParagraph'] } }]) {
    const result = await executeTool(
      editor,
      {
        id: 'direct',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0, html: '<h2>Unexpected</h2>', ...input },
      },
      { bullet: null, ordered: null },
    )
    expect(result).toMatchObject({ isError: true, mutated: false })
    expect(editor.getJSON()).toEqual(before)
  }
})

it('rejects mismatched target extents/types, incomplete sequences and cell targets without widening authority', async () => {
  const { editor, client, query, replace } = open()
  const source = await query()
  const target = source.items[1].target
  const before = client.context()
  for (const bad of [
    { ...target, expectedType: 'docHeading' },
    { ...target, range: { ...target.range, to: target.range.to - 1 } },
    { ...target, kind: 'cell', cell: { row: 0, col: 0 } },
  ]) {
    expect(await replace({ preconditions: { targets: [bad] } })).toMatchObject({
      isError: true,
      committed: false,
      revision: before.revision,
    })
  }
  expect(await replace({ startBlockIndex: 0, preconditions: { targets: [target] } })).toMatchObject(
    { isError: true, committed: false, revision: before.revision },
  )
  expect(editor.commands.undo()).toBe(false)
})

it('checks revision-bound targets, expected content and preserved block identities before replacement', async () => {
  const { editor, client, query, replace } = open()
  const source = await query()
  const heading = source.items[2].target
  const body = source.items[1].target
  const before = client.context()
  const wrongContent = await replace({
    preconditions: { targets: [{ ...body, expectedHtml: '<p>Different</p>' }] },
  })
  expect(wrongContent).toMatchObject({ isError: true, committed: false, revision: before.revision })
  const overlaps = await replace({ endBlockIndex: 2, preconditions: { preserve: [heading] } })
  expect(overlaps).toMatchObject({ isError: true, committed: false, revision: before.revision })
  const success = await replace({
    preconditions: { targets: [body], allowedTypes: ['docParagraph'], preserve: [heading] },
  })
  expect(success.isError, success.output).not.toBe(true)
  expect(editor.state.doc.textContent).toBe('HeadingUpdatedNext')
  const stale = await replace({ preconditions: { targets: [body] } })
  expect(stale).toMatchObject({ isError: true, committed: false, errorCode: 'stale_target' })
})

it('permits explicitly targeted heading edits and previews without changing content, history or revision', async () => {
  const { editor, client, query, replace } = open()
  const source = await query()
  const input = {
    startBlockIndex: 0,
    endBlockIndex: 0,
    html: '<h2>New heading</h2>',
    preconditions: { targets: [source.items[0].target], allowedTypes: ['docHeading'] },
  }
  const before = client.context()
  const preview = await replace({ ...input, dryRun: true })
  expect(preview).toMatchObject({
    committed: false,
    revision: before.revision,
    receipt: { outcome: 'preview' },
  })
  expect(editor.state.doc.textContent).toBe('HeadingBodyNext')
  expect(editor.commands.undo()).toBe(false)
  const result = await replace(input)
  expect(result.isError, result.output).not.toBe(true)
  expect(editor.state.doc.firstChild?.textContent).toBe('New heading')
})

it('returns missing and extra structure declarations for a rejected HTML edit and corresponding preview/commit effects', async () => {
  const { editor, client, query, replace } = open({
    type: 'doc',
    content: [mathTable, { type: 'docParagraph', content: [{ type: 'text', text: 'Tail' }] }],
  })
  const source = await query()
  const structures = source.items[0].structures
  const table = structures.find((item: { kind: string }) => item.kind === 'table').id
  const math = structures.find((item: { kind: string }) => item.kind === 'math').id
  const input = {
    startBlockIndex: 0,
    endBlockIndex: 0,
    html: '<table><tr><td>Changed</td></tr></table>',
  }
  const before = client.context()
  const rejected = await replace({ ...input, deleteStructures: [table] })
  expect(rejected).toMatchObject({
    isError: true,
    committed: false,
    revision: before.revision,
    structurePlan: { required: [math], missing: [math], extra: [table], complete: true },
  })
  expect(rejected.structurePlan?.wouldPreserve).toContainEqual(
    expect.objectContaining({ id: table, kind: 'table' }),
  )
  const preview = await replace({ ...input, deleteStructures: [math], dryRun: true })
  expect(preview).toMatchObject({
    committed: false,
    revision: before.revision,
    structurePlan: { required: [math] },
    receipt: {
      outcome: 'preview',
      operations: [
        {
          op: 'replace_blocks',
          structuralEffects: [{ reference: math, kind: 'math', effect: 'removed' }],
        },
      ],
    },
  })
  const result = await replace({ ...input, deleteStructures: [math] })
  expect(result.isError, result.output).not.toBe(true)
  expect(result.receipt?.operations).toEqual(preview.receipt?.operations)
  expect(editor.state.doc.textContent).toBe('ChangedTail')
})

it('explains native structural rejection atomically, using source-bound identities after preceding edits', async () => {
  const { editor, client, query } = open({
    type: 'doc',
    content: [
      { type: 'docParagraph', content: [{ type: 'text', text: 'Lead' }] },
      mathTable,
      { type: 'docParagraph', content: [{ type: 'text', text: 'Tail' }] },
    ],
  })
  const source = await query()
  const required = source.items[1].structures.map((item: { id: string }) => item.id)
  const before = client.context()
  const result = await client.execute({
    id: 'native-refusal',
    name: 'apply_ops',
    input: {
      ops: [
        {
          op: 'findReplace',
          target: { blockIndexes: [0] },
          find: 'Lead',
          replace: 'A longer lead',
        },
        { op: 'deleteBlocks', target: { blockIndexes: [1] } },
      ],
    },
  })
  expect(result).toMatchObject({
    isError: true,
    committed: false,
    revision: before.revision,
    structurePlan: {
      required: expect.arrayContaining(required),
      missing: expect.arrayContaining(required),
      extra: [],
    },
  })
  expect(editor.state.doc.textContent).toBe('LeadTail')
  expect(editor.commands.undo()).toBe(false)
})

it('requires a fresh observation instead of inventing source IDs after an unmappable move', async () => {
  const { editor, client } = open({
    type: 'doc',
    content: [
      { type: 'docParagraph', content: [{ type: 'text', text: 'Lead' }] },
      mathTable,
      { type: 'docParagraph', content: [{ type: 'text', text: 'Tail' }] },
    ],
  })
  const before = client.context()
  const result = await client.execute({
    id: 'unmapped',
    name: 'apply_ops',
    input: {
      ops: [
        { op: 'moveBlocks', blockIndexes: [1], afterBlockIndex: -1 },
        { op: 'deleteBlocks', target: { blockIndexes: [0] } },
      ],
    },
  })
  expect(result).toMatchObject({ isError: true, committed: false, revision: before.revision })
  expect(result.output).toMatch(/identit|query/i)
  expect(result.structurePlan).toBeUndefined()
  expect(editor.state.doc.textContent).toBe('LeadTail')
  expect(editor.commands.undo()).toBe(false)
})

it('bounds structure diagnostics without weakening validation of a large deletion', async () => {
  const many = {
    ...mathTable,
    content: [
      {
        type: 'docTableRow',
        content: [
          {
            type: 'docTableCell',
            content: [
              {
                type: 'docParagraph',
                content: Array.from({ length: 70 }, () => ({
                  type: 'docInlineMath',
                  attrs: { latex: 'x' },
                })),
              },
            ],
          },
        ],
      },
    ],
  }
  const { client } = open({ type: 'doc', content: [many, { type: 'docParagraph' }] })
  const before = client.context()
  const result = await client.execute({
    id: 'large-intent',
    name: 'apply_ops',
    input: { ops: [{ op: 'deleteBlocks', target: { blockIndexes: [0] } }] },
  })
  expect(result).toMatchObject({
    isError: true,
    committed: false,
    revision: before.revision,
    structurePlan: { complete: false, counts: { required: 71, missing: 71 } },
  })
  expect(result.structurePlan?.required).toHaveLength(64)
  expect(result.structurePlan?.missing).toHaveLength(64)
  const ids: string[] = []
  let input: Record<string, unknown> = { domain: 'body', startBlockIndex: 0, endBlockIndex: 0 }
  for (;;) {
    const response = await client.execute({
      id: crypto.randomUUID(),
      name: 'query_document',
      input,
    })
    const item = JSON.parse(response.output).items[0]
    ids.push(...item.structures.map((structure: { id: string }) => structure.id))
    if (item.structuresPage.complete) break
    input = { cursor: item.structuresPage.nextCursor }
  }
  const preview = await client.execute({
    id: 'large-preview',
    name: 'apply_ops',
    input: {
      dryRun: true,
      ops: [{ op: 'deleteBlocks', target: { blockIndexes: [0] }, deleteStructures: ids }],
    },
  })
  expect(preview.isError, preview.output).not.toBe(true)
  expect(preview.receipt?.operations?.[0]).toMatchObject({
    structuralEffectCount: 71,
    structuralEffectsComplete: false,
  })
  expect(preview.receipt?.operations?.[0].structuralEffects).toHaveLength(64)
  expect(preview).toMatchObject({ committed: false, revision: before.revision })
})
