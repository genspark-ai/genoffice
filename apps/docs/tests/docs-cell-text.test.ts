import { afterEach, expect, it } from 'vitest'
import { Editor, type JSONContent } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { createDocsRuntime, type DocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { t } from '../src/renderer/i18n/locale'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { createDocsMcpServer } from '../src/services/docs-mcp'
import { createDocsCommandRouter } from '../src/renderer/runtime/command-router'

const editors: Editor[] = []
const runtimes: DocsRuntime[] = []
afterEach(() => {
  runtimes.splice(0).forEach((runtime) => runtime.dispose())
  editors.splice(0).forEach((editor) => editor.destroy())
})

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const tc = (text: string, props = '') => `<w:tc><w:tcPr>${props}</w:tcPr>${p(text)}</w:tc>`

async function fixture(bodyXml?: string) {
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml:
        bodyXml ??
        p('INTRO') +
          '<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/></w:tblPr>' +
          '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
          '<w:tr>' +
          tc('MERGED', '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>') +
          tc(
            '',
            '<w:shd w:fill="ABCDEF"/><w:tcBorders><w:left w:val="single" w:sz="8" w:color="123456"/></w:tcBorders>',
          ) +
          '</w:tr><w:tr>' +
          tc('', '<w:gridSpan w:val="2"/><w:vMerge/>') +
          tc('KEEP') +
          '</w:tr></w:tbl>' +
          p('AFTER_TABLE'),
    }),
  )
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.push(editor)
  const runtime = createDocsRuntime({
    current: () => ({
      editor,
      identity: editor,
      numIds: { bullet: null, ordered: null },
      autoAcceptUntracked: true,
    }),
  })
  runtimes.push(runtime)
  const caller = runtime.connect({ id: 'cell-text' })
  caller.context()
  const read = async () =>
    (
      await caller.execute({
        id: crypto.randomUUID(),
        name: 'read_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 2 },
      })
    ).output
  const saved = async () => {
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const bytes = await saveDocx(parsed, plan.saveBlocks)
    const reopened = await parseDocx(bytes)
    const zip = await JSZip.loadAsync(bytes)
    return {
      reopened,
      xml: new DOMParser().parseFromString(
        await zip.file('word/document.xml')!.async('string'),
        'application/xml',
      ),
    }
  }
  const apply = (ops: Record<string, unknown>[]) =>
    caller.execute({
      id: crypto.randomUUID(),
      name: 'apply_ops',
      input: { ops },
    })
  return { caller, read, saved, apply, editor, runtime }
}

it('fills an addressed empty cell without rebuilding merged geometry or neighboring content', async () => {
  const { caller, read, saved } = await fixture()
  const result = await caller.execute({
    id: 'fill',
    name: 'apply_ops',
    input: {
      ops: [
        {
          op: 'setTableCellText',
          target: { blockIndexes: [1] },
          range: { row: 0, col: 2 },
          text: 'FILLED',
        },
      ],
    },
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(result.committed).toBe(true)
  expect(result).toMatchObject({ matchCount: 1, changeCount: 1 })
  const html = new DOMParser().parseFromString(await read(), 'text/html')
  expect(html.querySelector('[data-row="0"][data-col="2"]')?.textContent).toBe('FILLED')
  expect(html.querySelector('[colspan="2"][rowspan="2"]')?.textContent).toBe('MERGED')
  expect(html.querySelector('[data-row="1"][data-col="2"]')?.textContent).toBe('KEEP')
  expect(html.body.textContent).toContain('AFTER_TABLE')
  const { xml, reopened } = await saved()
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  expect(
    [...xml.getElementsByTagNameNS(ns, 'gridCol')].map((n) => n.getAttributeNS(ns, 'w')),
  ).toEqual(['2400', '2400', '2400'])
  expect(xml.getElementsByTagNameNS(ns, 'vMerge')).toHaveLength(2)
  expect(xml.getElementsByTagNameNS(ns, 'gridSpan')).toHaveLength(2)
  expect(xml.getElementsByTagNameNS(ns, 'shd')[0].getAttributeNS(ns, 'fill')).toBe('ABCDEF')
  expect(xml.getElementsByTagNameNS(ns, 'tblW')[0].getAttributeNS(ns, 'w')).toBe('7200')
  expect(xml.getElementsByTagNameNS(ns, 'left')[0].getAttributeNS(ns, 'color')).toBe('123456')
  expect(reopened.blocks.find((block) => block.type === 'table')?.table?.rows[0][1].paras).toEqual([
    'FILLED',
  ])
  const tableIndex = reopened.blocks.findIndex((block) => block.type === 'table')
  expect(reopened.blocks[tableIndex + 1]?.runs?.map((run) => run.text).join('')).toBe('AFTER_TABLE')
})

it('rejects a target that includes a table and another block without changing the document', async () => {
  const { caller, read, apply } = await fixture()
  const before = await read()
  const context = caller.context()
  const result = await apply([
    {
      op: 'setTableCellText',
      target: { blockIndexes: [0, 1] },
      range: { row: 0, col: 2 },
      text: 'UNSAFE',
    },
  ])
  expect(result).toMatchObject({ isError: true, committed: false, revision: context.revision })
  expect(await read()).toBe(before)
})

it('refuses omitted grid slots rather than reporting an edit that disappears on save', async () => {
  const { caller, read, apply, saved } = await fixture(
    p('INTRO') +
      '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>' +
      tc('REAL') +
      '</w:tr></w:tbl>' +
      p('AFTER_TABLE'),
  )
  const before = await read()
  const context = caller.context()
  const result = await apply([
    {
      op: 'setTableCellText',
      target: { blockIndexes: [1] },
      range: { row: 0, col: 0 },
      text: 'VANISHES',
    },
  ])
  expect(result).toMatchObject({ isError: true, committed: false, revision: context.revision })
  expect(await read()).toBe(before)
  expect((await saved()).xml.documentElement.textContent).not.toContain('VANISHES')
})

it('keeps required paragraph properties on clear and reports repeated clear as zero changes', async () => {
  const { apply, saved } = await fixture(
    p('INTRO') +
      '<w:tbl><w:tr><w:tc>' +
      '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>FIRST</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:spacing w:after="160"/></w:pPr><w:r><w:t>SECOND</w:t></w:r></w:p>' +
      '</w:tc></w:tr></w:tbl>' +
      p('AFTER_TABLE'),
  )
  const clear = {
    op: 'setTableCellText',
    target: { blockIndexes: [1] },
    range: { row: 0, col: 0 },
    text: '',
  }
  expect((await apply([clear])).isError).not.toBe(true)
  const again = await apply([clear])
  expect(again.isError).not.toBe(true)
  expect(again).toMatchObject({ committed: false, matchCount: 1, changeCount: 0 })
  expect(again.output).toBe(t('aiCmdNoneUnchanged', { count: 1 }))
  const { xml } = await saved()
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  const cell = xml.getElementsByTagNameNS(ns, 'tc')[0]
  expect(cell.getElementsByTagNameNS(ns, 'p')).toHaveLength(2)
  expect(cell.getElementsByTagNameNS(ns, 'jc')[0].getAttributeNS(ns, 'val')).toBe('right')
  expect(cell.getElementsByTagNameNS(ns, 'spacing')[0].getAttributeNS(ns, 'after')).toBe('160')
  expect(cell.textContent).toBe('')
})

it('clears and refills a covered combined-span cell once, including literal authored tokens', async () => {
  const { apply, read, saved } = await fixture()
  const target = { blockIndexes: [1] }
  expect(
    (await apply([{ op: 'setTableCellText', target, range: { row: 1, col: 1 }, text: '' }]))
      .isError,
  ).not.toBe(true)
  expect(
    (
      await apply([
        {
          op: 'setTableCellText',
          target,
          range: { row: 0, col: 0, rowEnd: 1, colEnd: 1 },
          text: '{{AUTHORED}}\nLine',
        },
      ])
    ).isError,
  ).not.toBe(true)
  const html = new DOMParser().parseFromString(await read(), 'text/html')
  expect(html.querySelector('[colspan="2"][rowspan="2"]')?.textContent).toContain('{{AUTHORED}}')
  const { xml } = await saved()
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  const texts = [...xml.getElementsByTagNameNS(ns, 't')].map((n) => n.textContent)
  expect(texts.filter((text) => text === '{{AUTHORED}}')).toHaveLength(1)
  expect(texts).toContain('Line')
  expect(texts).toContain('KEEP')
  expect(xml.getElementsByTagNameNS(ns, 'br')).toHaveLength(1)
  expect(xml.getElementsByTagNameNS(ns, 'vMerge')).toHaveLength(2)
})

it('deduplicates physical owners in an explicit bounded bulk edit and maps subsequent operations', async () => {
  const { apply, saved } = await fixture()
  const result = await apply([
    {
      op: 'findReplace',
      target: { blockIndexes: [0] },
      find: 'INTRO',
      replace: 'A much longer introduction',
    },
    {
      op: 'setTableCellText',
      target: { nodeType: 'table' },
      range: { row: 0, col: 0, rowEnd: 1, colEnd: 2 },
      text: 'BULK',
      match: 'all',
    },
    {
      op: 'setTableCellText',
      target: { nodeType: 'table' },
      range: { row: 1, col: 2 },
      text: 'LAST',
    },
  ])
  expect(result.isError, result.output).not.toBe(true)
  expect(result.output).toContain('setTableCellText: 1')
  expect(result).toMatchObject({ matchCount: 5, changeCount: 5 })
  const { xml } = await saved()
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  const cells = [...xml.getElementsByTagNameNS(ns, 'tc')]
  expect(cells.map((cell) => cell.textContent)).toEqual(['BULK', 'BULK', '', 'LAST'])
  expect(xml.documentElement.textContent).toContain('A much longer introduction')
  expect(xml.documentElement.textContent).toContain('AFTER_TABLE')
})

it.each([
  [{ row: 0, col: 0, rowEnd: 1, colEnd: 2 }, undefined],
  [{ row: 0, col: 1, colEnd: 2 }, 'all'],
  [{ row: 2, col: 0 }, undefined],
  [{ row: 0, col: 3 }, undefined],
  [{ row: -1, col: 0 }, undefined],
  [{ row: 1, col: 0, rowEnd: 0 }, undefined],
] as const)(
  'refuses ambiguous, partial-merged or invalid range %j atomically',
  async (range, match) => {
    const { caller, apply, read } = await fixture()
    const before = await read()
    const context = caller.context()
    const result = await apply([
      {
        op: 'setTableCellText',
        target: { blockIndexes: [1] },
        range: { row: 0, col: 2 },
        text: 'ROLLBACK',
      },
      {
        op: 'setTableCellText',
        target: { blockIndexes: [1] },
        range,
        text: 'INVALID',
        ...(match ? { match } : {}),
      },
    ])
    expect(result).toMatchObject({ isError: true, committed: false, revision: context.revision })
    expect(await read()).toBe(before)
  },
)

it.each([
  { type: 'text', text: 'PRIVATE', marks: [{ type: 'redaction', attrs: { label: 'private' } }] },
  { type: 'text', text: 'DELETED', marks: [{ type: 'del', attrs: { author: 'A', id: '1' } }] },
  { type: 'text', text: 'FIELD', marks: [{ type: 'instrField', attrs: { instr: 'DATE' } }] },
  { type: 'text', text: 'REF', marks: [{ type: 'refField', attrs: { name: 'Target' } }] },
  {
    type: 'docInlineImage',
    attrs: { dataUrl: 'data:image/png;base64,AA==', widthPx: 20, heightPx: 20 },
  },
  { type: 'docInlineMath', attrs: { latex: 'x+y' } },
  { type: 'docNoteRef', attrs: { kind: 'footnote', id: '1', num: 1 } },
])(
  'refuses whole-cell text replacement across protected or structured inline content %#',
  async (inline) => {
    const { caller, apply, read, editor } = await fixture()
    const doc: JSONContent = editor.getJSON()
    doc.content![1].content![0].content![1].content![0].content = [
      { type: 'text', text: 'Before ' },
      inline,
      { type: 'text', text: ' after' },
    ]
    editor.commands.setContent(doc)
    const context = caller.context()
    const before = await read()
    const result = await apply([
      {
        op: 'setTableCellText',
        target: { blockIndexes: [1] },
        range: { row: 1, col: 2 },
        text: 'ROLLBACK',
      },
      {
        op: 'setTableCellText',
        target: { blockIndexes: [1] },
        range: { row: 0, col: 2 },
        text: '',
      },
    ])
    expect(result).toMatchObject({ isError: true, committed: false, revision: context.revision })
    expect(result.output).not.toContain('PRIVATE')
    expect(await read()).toBe(before)
  },
)

it('refuses a cell containing a nested table and preserves the table through save', async () => {
  const nested = '<w:tbl><w:tr>' + tc('NESTED') + '</w:tr></w:tbl>'
  const { apply, saved } = await fixture(
    p('INTRO') +
      '<w:tbl><w:tr><w:tc>' +
      p('Ordinary') +
      nested +
      p('Tail') +
      '</w:tc></w:tr></w:tbl>' +
      p('AFTER_TABLE'),
  )
  expect(
    (
      await apply([
        {
          op: 'setTableCellText',
          target: { blockIndexes: [1] },
          range: { row: 0, col: 0 },
          text: '',
        },
      ])
    ).isError,
  ).toBe(true)
  const { xml } = await saved()
  expect(
    xml.getElementsByTagNameNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'tbl',
    ),
  ).toHaveLength(2)
  expect(xml.documentElement.textContent).toContain('NESTED')
})

it('discovers and uses the cell text operation through standard MCP and persists a clear', async () => {
  const { runtime, saved } = await fixture()
  const router = createDocsCommandRouter(runtime)
  const server = createDocsMcpServer({
    version: 'test',
    endpoint: {
      request: (message, signal) => router.handle(message, signal),
      closeCaller: (callerId) => router.disconnect(callerId),
    },
  })
  const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair()
  const client = new Client({ name: 'cell-text', version: '1' })
  try {
    await server.connect(serverTransport)
    await client.connect(clientTransport)
    await client.callTool({ name: 'get_document_context', arguments: {} })
    const caps = await client.callTool({ name: 'get_editor_capabilities', arguments: {} })
    expect(JSON.stringify(caps)).toContain('setTableCellText')
    for (const text of ['MCP', '']) {
      const result = await client.callTool({
        name: 'apply_ops',
        arguments: {
          ops: [
            {
              op: 'setTableCellText',
              target: { blockIndexes: [1] },
              range: { row: 1, col: 2 },
              text,
            },
          ],
        },
      })
      expect(result.isError, JSON.stringify(result)).not.toBe(true)
      expect(result.structuredContent).toMatchObject({
        committed: true,
        matchCount: 1,
        changeCount: 1,
      })
    }
    const { reopened, xml } = await saved()
    expect(
      reopened.blocks.find((block) => block.type === 'table')?.table?.rows[1][1].paras,
    ).toEqual([''])
    expect(xml.documentElement.textContent).not.toContain('MCP')
    expect(xml.documentElement.textContent).toContain('MERGED')
  } finally {
    await client.close()
    await server.close()
    router.dispose()
  }
})

it('does not split user typing history when a cell batch is rejected', async () => {
  const { caller, editor, apply, read } = await fixture()
  const before = await read()
  editor.commands.insertContent('a')
  caller.context()
  const result = await apply([
    {
      op: 'setTableCellText',
      target: { blockIndexes: [1] },
      range: { row: 0, col: 2 },
      text: 'ROLLBACK',
    },
    { op: 'setTableCellText', target: { blockIndexes: [1] }, range: { row: 9, col: 0 }, text: '' },
  ])
  expect(result).toMatchObject({ isError: true, committed: false })
  editor.commands.insertContent('b')
  caller.context()
  expect(
    (await caller.execute({ id: 'undo', name: 'document_history', input: { action: 'undo' } }))
      .isError,
  ).not.toBe(true)
  expect(await read()).toBe(before)
})

it.each([
  {
    name: 'horizontal',
    rows: '<w:tr>' + tc('SPAN', '<w:gridSpan w:val="2"/>') + '</w:tr>',
    range: { row: 0, col: 1 },
    attr: 'gridSpan',
  },
  {
    name: 'vertical',
    rows:
      '<w:tr>' +
      tc('SPAN', '<w:vMerge w:val="restart"/>') +
      '</w:tr><w:tr>' +
      tc('', '<w:vMerge/>') +
      '</w:tr>',
    range: { row: 1, col: 0 },
    attr: 'vMerge',
  },
])('fills a covered $name-only span without splitting it', async ({ rows, range, attr }) => {
  const { apply, saved } = await fixture(
    p('INTRO') + '<w:tbl>' + rows + '</w:tbl>' + p('AFTER_TABLE'),
  )
  const result = await apply([
    { op: 'setTableCellText', target: { blockIndexes: [1] }, range, text: 'FILLED' },
  ])
  expect(result.isError, result.output).not.toBe(true)
  const { xml } = await saved()
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  expect(xml.getElementsByTagNameNS(ns, attr).length).toBeGreaterThan(0)
  expect(
    [...xml.getElementsByTagNameNS(ns, 't')].filter((t) => t.textContent === 'FILLED'),
  ).toHaveLength(1)
  expect(xml.documentElement.textContent).not.toContain('SPAN')
})

it('refuses a stale caller before editing addressed cell text', async () => {
  const { runtime, caller, editor, apply, read } = await fixture()
  const stale = runtime.connect({ id: 'stale-cell' })
  stale.context()
  editor.commands.insertContent('User edit ')
  caller.context()
  const before = await read()
  const result = await stale.execute({
    id: 'stale',
    name: 'apply_ops',
    input: {
      ops: [
        {
          op: 'setTableCellText',
          target: { blockIndexes: [1] },
          range: { row: 0, col: 2 },
          text: 'STALE',
        },
      ],
    },
  })
  expect(result).toMatchObject({ isError: true, committed: false, errorCode: 'stale_target' })
  expect(await read()).toBe(before)
  expect(
    (
      await apply([
        {
          op: 'setTableCellText',
          target: { blockIndexes: [1] },
          range: { row: 0, col: 2 },
          text: 'FRESH',
        },
      ])
    ).isError,
  ).not.toBe(true)
})

it.each([
  { kind: 'symbol', xml: '<w:r><w:sym w:font="Symbol" w:char="F02D"/></w:r>', text: '' },
  { kind: 'symbol', xml: '<w:r><w:sym w:font="Symbol" w:char="F02D"/></w:r>', text: 'FILL' },
  ...['', 'FILL'].map((text) => ({
    kind: 'checkbox',
    text,
    xml:
      '<w:sdt xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:sdtPr><w14:checkbox>' +
      '<w14:checked w14:val="0"/><w14:checkedState w14:val="2611" w14:font="MS Gothic"/>' +
      '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>' +
      '<w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>',
  })),
])(
  'refuses erasing or copying semantic $kind payload with replacement "$text"',
  async ({ kind, xml: inline, text }) => {
    const { apply, read, saved } = await fixture(
      p('INTRO') +
        '<w:tbl><w:tr><w:tc><w:p>' +
        inline +
        '<w:r><w:t> tail</w:t></w:r></w:p></w:tc>' +
        tc('NEIGHBOR') +
        '</w:tr></w:tbl>' +
        p('AFTER_TABLE'),
    )
    const before = await read()
    const result = await apply([
      { op: 'setTableCellText', target: { blockIndexes: [1] }, range: { row: 0, col: 0 }, text },
    ])
    expect(result).toMatchObject({ isError: true, committed: false })
    expect(await read()).toBe(before)
    // A supported neighboring edit regenerates the table, exercising preserved payload on save.
    expect(
      (
        await apply([
          {
            op: 'setTableCellText',
            target: { blockIndexes: [1] },
            range: { row: 0, col: 1 },
            text: 'CHANGED',
          },
        ])
      ).isError,
    ).not.toBe(true)
    const { xml } = await saved()
    if (kind === 'symbol') {
      const symbols = xml.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'sym',
      )
      expect(symbols).toHaveLength(1)
      expect(
        symbols[0].getAttributeNS(
          'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
          'char',
        ),
      ).toBe('F02D')
    } else {
      const ns = 'http://schemas.microsoft.com/office/word/2010/wordml'
      expect(xml.getElementsByTagNameNS(ns, 'checkbox')).toHaveLength(1)
      expect(xml.getElementsByTagNameNS(ns, 'checked')[0].getAttributeNS(ns, 'val')).toBe('0')
      expect(xml.documentElement.textContent).toContain('☐')
    }
    expect(xml.documentElement.textContent).toContain(' tail')
    expect(xml.documentElement.textContent).toContain('CHANGED')
  },
)
