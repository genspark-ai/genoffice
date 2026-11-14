// Modified for ForaOffice: precise logical-text writes share the native search engine.
import { afterEach, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { createDocsRuntime, type DocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { executeOps } from '../src/renderer/ai/ops'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { createDocsCommandRouter } from '../src/renderer/runtime/command-router'
import { createDocsMcpServer } from '../src/services/docs-mcp'

const editors: Editor[] = []
const runtimes: DocsRuntime[] = []
afterEach(() => {
  runtimes.splice(0).forEach((runtime) => runtime.dispose())
  editors.splice(0).forEach((editor) => editor.destroy())
})

async function fixture(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
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
  editors.push(editor)
  runtimes.push(runtime)
  const caller = runtime.connect({ id: 'logical-text-writer' })
  caller.context()
  const read = async (start = 0, end = 0) =>
    caller.execute({
      id: 'read',
      name: 'read_blocks',
      input: { startBlockIndex: start, endBlockIndex: end },
    })
  const write = (ops: object[]) =>
    caller.execute({ id: 'write', name: 'apply_ops', input: { ops } })
  return { editor, runtime, caller, read, write, saved: () => saved! }
}

it('replaces a five-run date once, inherits match-start formatting and preserves neighbors through save/reopen', async () => {
  const { caller, read, write, saved } = await fixture(
    '<w:p><w:r><w:t xml:space="preserve">KEEP_BEFORE </w:t></w:r>' +
      '<w:r><w:rPr><w:b/></w:rPr><w:t>二〇</w:t></w:r>' +
      '<w:r><w:rPr><w:i/></w:rPr><w:t>**</w:t></w:r>' +
      '<w:r><w:t>年</w:t></w:r><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>**</w:t></w:r>' +
      '<w:r><w:t>月</w:t></w:r><w:r><w:t xml:space="preserve"> KEEP_AFTER</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>KEEP_NEIGHBOR</w:t></w:r></w:p>',
  )
  expect(
    await write([
      {
        op: 'findReplace',
        find: '二〇**年**月',
        replace: 'NEW_DATE',
        target: { blockIndexes: [0] },
      },
    ]),
  ).toMatchObject({ committed: true, mutated: true })
  const after = await read(0, 1)
  expect(after.output).toContain('KEEP_BEFORE <strong>NEW_DATE</strong> KEEP_AFTER')
  expect(after.output).toContain('KEEP_NEIGHBOR')
  expect(after.output).not.toContain('二〇')
  expect(await caller.save({ path: '/logical-text.docx', overwrite: true })).toMatchObject({
    ok: true,
    pending: false,
  })
  const reopened = await parseDocx(saved())
  expect(reopened.blocks[0].runs?.map((run) => run.text).join('')).toBe(
    'KEEP_BEFORE NEW_DATE KEEP_AFTER',
  )
  expect(reopened.blocks[0].runs?.find((run) => run.text === 'NEW_DATE')?.bold).toBe(true)
  const xml = await (await JSZip.loadAsync(saved())).file('word/document.xml')!.async('string')
  expect(xml).toContain('KEEP_NEIGHBOR')
})

it.each(['absent', 'same same'])(
  'refuses a non-unique precise match in %s and rolls back preceding operations',
  async (text) => {
    const { caller, read, write } = await fixture(`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)
    const before = await read()
    const result = await write([
      { op: 'setFont', target: { blockIndexes: [0] }, bold: true },
      { op: 'findReplace', find: 'same', replace: 'changed', target: { blockIndexes: [0] } },
    ])
    expect(result).toMatchObject({
      isError: true,
      committed: false,
      mutated: false,
      revision: before.revision,
    })
    expect(result.output).toMatch(/exactly one/i)
    expect((await read()).output).toBe(before.output)
    expect(caller.context().revision).toBe(before.revision)
  },
)

it('does not rewrite or match through a Word field', async () => {
  const { read, write } = await fixture(
    '<w:p><w:r><w:t>A</w:t></w:r>' +
      '<w:fldSimple w:instr="DATE"><w:r><w:t>DATE</w:t></w:r></w:fldSimple>' +
      '<w:r><w:t>B</w:t></w:r></w:p>',
  )
  const before = await read()
  for (const find of ['DATE', 'ADATEB']) {
    expect(
      await write([{ op: 'findReplace', find, replace: 'BROKEN', target: { blockIndexes: [0] } }]),
    ).toMatchObject({ isError: true, committed: false, revision: before.revision })
    expect((await read()).output).toBe(before.output)
  }
})

it.each([
  { kind: 'symbol', glyph: '\u2212', xml: '<w:r><w:sym w:font="Symbol" w:char="F02D"/></w:r>' },
  {
    kind: 'checkbox',
    glyph: '☐',
    xml:
      '<w:sdt xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wordx="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sdtPr><wordx:alias wordx:val="Agree"/><w14:checkbox>' +
      '<w14:checked w14:val="0"/><w14:checkedState w14:val="2611" w14:font="MS Gothic"/>' +
      '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>' +
      '<w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>',
  },
])(
  'does not erase, copy or match through semantic $kind payload',
  async ({ kind, glyph, xml: inline }) => {
    const { caller, read, write, saved } = await fixture(
      '<w:p><w:r><w:t>A</w:t></w:r>' + inline + '<w:r><w:t>B</w:t></w:r></w:p>',
    )
    const before = await read()
    for (const find of [glyph, `A${glyph}B`]) {
      for (const replace of ['', 'BROKEN']) {
        expect(await write([{ op: 'findReplace', find, replace }])).toMatchObject({
          isError: true,
          committed: false,
          revision: before.revision,
        })
        expect((await read()).output).toBe(before.output)
      }
    }
    // Regenerate the same paragraph with a supported neighboring edit before saving.
    expect(await write([{ op: 'findReplace', find: 'B', replace: 'NEIGHBOR' }])).toMatchObject({
      committed: true,
      matchCount: 1,
      changeCount: 1,
    })
    expect(await caller.save({ path: '/semantic-payload.docx', overwrite: true })).toMatchObject({
      ok: true,
      pending: false,
    })
    const reopened = await parseDocx(saved())
    expect(reopened.blocks[0].runs?.map((run) => run.text).join('')).toBe(`A${glyph}NEIGHBOR`)
    const xml = new DOMParser().parseFromString(
      await (await JSZip.loadAsync(saved())).file('word/document.xml')!.async('string'),
      'application/xml',
    )
    expect(
      xml.getElementsByTagName('parsererror'),
      xml.documentElement.textContent ?? '',
    ).toHaveLength(0)
    if (kind === 'symbol') {
      const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
      const symbols = xml.getElementsByTagNameNS(ns, 'sym')
      expect(symbols).toHaveLength(1)
      expect(symbols[0].getAttributeNS(ns, 'char')).toBe('F02D')
    } else {
      const ns = 'http://schemas.microsoft.com/office/word/2010/wordml'
      expect(xml.getElementsByTagNameNS(ns, 'checkbox')).toHaveLength(1)
      expect(xml.getElementsByTagNameNS(ns, 'checked')[0].getAttributeNS(ns, 'val')).toBe('0')
      const wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
      expect(xml.getElementsByTagNameNS(wordNs, 'alias')[0].getAttributeNS(wordNs, 'val')).toBe(
        'Agree',
      )
    }
  },
)

it.each(['refusal', 'no-op'] as const)(
  'does not split user undo history on a precise %s',
  async (decision) => {
    const { editor, caller, read, write } = await fixture('<w:p><w:r><w:t>BASE</w:t></w:r></w:p>')
    editor.commands.insertContent('a')
    editor.commands.insertContent('b')
    const before = caller.context()
    const result = await write([
      {
        op: 'findReplace',
        find: decision === 'refusal' ? 'missing' : 'ab',
        replace: decision === 'refusal' ? 'changed' : 'ab',
        target: { blockIndexes: [0] },
      },
    ])
    expect(result).toMatchObject({ committed: false, mutated: false, revision: before.revision })
    if (decision === 'no-op') expect(result).toMatchObject({ matchCount: 1, changeCount: 0 })
    editor.commands.insertContent('c')
    caller.context()
    expect(
      await caller.execute({ id: 'undo', name: 'document_history', input: { action: 'undo' } }),
    ).toMatchObject({ committed: true })
    expect((await read()).output).toBe('<p>BASE</p>')
  },
)

it('requires a scope for bulk writes and reports occurrence counts within that scope', async () => {
  const { read, write } = await fixture(
    '<w:p><w:r><w:t>same same</w:t></w:r></w:p><w:p><w:r><w:t>same</w:t></w:r></w:p>',
  )
  expect(
    await write([{ op: 'findReplace', find: 'same', replace: 'NEW', match: 'all' }]),
  ).toMatchObject({ isError: true, committed: false, revision: 0 })
  expect(
    await write([
      {
        op: 'findReplace',
        find: 'same',
        replace: 'NEW',
        match: 'all',
        target: { blockIndexes: [0] },
      },
    ]),
  ).toMatchObject({ committed: true, matchCount: 2, changeCount: 2 })
  expect((await read(0, 1)).output).toBe('<p>NEW NEW</p>\n<p>same</p>')
})

it('keeps authored caret codes and replacement expressions literal', async () => {
  const { read, write } = await fixture('<w:p><w:r><w:t>^p ^&amp; $1 \\1</w:t></w:r></w:p>')
  expect(
    await write([{ op: 'findReplace', find: '^p ^& $1 \\1', replace: '^& $1 \\1 ^p' }]),
  ).toMatchObject({ committed: true, matchCount: 1, changeCount: 1 })
  expect((await read()).output).toBe('<p>^&amp; $1 \\1 ^p</p>')
})

it('does not join matches across paragraphs or table cells', async () => {
  const { read, write } = await fixture(
    '<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>' +
      '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  )
  const before = await read(0, 2)
  for (const find of ['AB', 'CD']) {
    expect(await write([{ op: 'findReplace', find, replace: 'BAD' }])).toMatchObject({
      isError: true,
      committed: false,
      revision: before.revision,
    })
  }
  expect((await read(0, 2)).output).toBe(before.output)
})

it.each([
  {
    type: 'docInlineImage',
    attrs: { dataUrl: 'data:image/png;base64,AA==', widthPx: 1, heightPx: 1 },
  },
  { type: 'docInlineMath', attrs: { latex: 'x', text: 'x', omml: '<m:oMath/>' } },
  {
    type: 'text',
    text: 'gone',
    marks: [{ type: 'del', attrs: { author: 'User', date: '2026-01-01T00:00:00Z' } }],
  },
])('does not match through a $type boundary', async (boundary) => {
  const { editor, caller, read, write } = await fixture('<w:p/>')
  editor.commands.setContent({
    type: 'doc',
    content: [
      {
        type: 'docParagraph',
        content: [{ type: 'text', text: 'A' }, boundary, { type: 'text', text: 'B' }],
      },
    ],
  })
  caller.context()
  const before = await read()
  for (const find of ['AB', ...(boundary.type === 'text' ? ['AgoneB', 'gone'] : [])]) {
    expect(await write([{ op: 'findReplace', find, replace: 'BAD' }])).toMatchObject({
      isError: true,
      committed: false,
      revision: before.revision,
    })
  }
  expect((await read()).output).toBe(before.output)
  expect(await write([{ op: 'findReplace', find: 'A', replace: 'NEW' }])).toMatchObject({
    committed: true,
    matchCount: 1,
    changeCount: 1,
  })
})

it('preserves intentional native UI bulk semantics', async () => {
  const { editor } = await fixture('<w:p><w:r><w:t>same same</w:t></w:r></w:p>')
  expect(
    executeOps(editor, [{ op: 'findReplace', find: 'same', replace: 'NEW' }], { source: 'ui' }),
  ).toMatchObject({ ok: true, results: [{ matchCount: 2, changeCount: 2 }] })
  expect(editor.state.doc.textContent).toBe('NEW NEW')
})

it('keeps full-width and ASCII literals distinct', async () => {
  const { read, write } = await fixture('<w:p><w:r><w:t>ABC ＡＢＣ</w:t></w:r></w:p>')
  expect(await write([{ op: 'findReplace', find: 'ABC', replace: 'NEW' }])).toMatchObject({
    committed: true,
    matchCount: 1,
    changeCount: 1,
  })
  expect((await read()).output).toBe('<p>NEW ＡＢＣ</p>')
})

it('returns truthful exact counts and atomic refusal through standard MCP', async () => {
  const { runtime } = await fixture('<w:p><w:r><w:t>same same</w:t></w:r></w:p>')
  const router = createDocsCommandRouter(runtime)
  const server = createDocsMcpServer({
    version: 'test',
    endpoint: {
      request: (message, signal) => router.handle(message, signal),
      closeCaller: (id) => router.disconnect(id),
    },
  })
  const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair()
  const client = new Client({ name: 'logical-text-mcp', version: '1' }, { capabilities: {} })
  try {
    await server.connect(serverTransport)
    await client.connect(clientTransport)
    await client.callTool({ name: 'get_document_context', arguments: {} })
    const ops = [{ op: 'findReplace', find: 'same', replace: 'NEW' }]
    expect(await client.callTool({ name: 'apply_ops', arguments: { ops } })).toMatchObject({
      isError: true,
      structuredContent: { committed: false, mutated: false, revision: 0 },
    })
    expect(
      await client.callTool({
        name: 'apply_ops',
        arguments: { ops: [{ ...ops[0], match: 'all', target: { blockIndexes: [0] } }] },
      }),
    ).toMatchObject({
      structuredContent: { committed: true, matchCount: 2, changeCount: 2, revision: 1 },
    })
  } finally {
    await client.close()
    await server.close()
  }
})
