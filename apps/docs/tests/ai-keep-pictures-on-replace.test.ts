// Modified for ForaOffice from GenOffice v0.12.852.
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TableMap } from '@tiptap/pm/tables'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  buildDocContext,
  countInlineImages,
  serializeRangeToHtml,
} from '../src/renderer/ai/protocol'
import { executeTool } from '../src/renderer/ai/tools'
import { acceptAllRevisions, rejectAllRevisions } from '../src/renderer/editor/revisions'
import { createDocsRuntime } from '../src/renderer/runtime/docs-runtime'
import type { DocsQueryResult } from '../src/shared/document-query'
import { setModuleLang } from '../src/renderer/i18n/locale'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
const picture = (marks?: Array<{ type: string; attrs?: Record<string, unknown> }>): JsonNode => ({
  type: 'docInlineImage',
  attrs: { dataUrl: PNG, widthPx: 120, heightPx: 80 },
  ...(marks ? { marks } : {}),
})
const text = (t: string): JsonNode => ({ type: 'text', text: t })
const para = (content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
})
const NUM_IDS = { bullet: null, ordered: null }
const cell = (content: JsonNode[], attrs?: Record<string, unknown>): JsonNode => ({
  type: 'docTableCell',
  attrs,
  content: [para(content)],
})
const table = (rows: JsonNode[][]): JsonNode => ({
  type: 'docTable',
  content: rows.map((content) => ({ type: 'docTableRow', content })),
})
const parameterTable = (rows: number, firstImage = '') =>
  `<table>${Array.from(
    { length: rows },
    (_, i) => `<tr><td>Parameter ${i}</td><td>${i === 0 ? firstImage : ''}</td></tr>`,
  ).join('')}</table>`

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.add(editor)
  return editor
}

function pictureCount(editor: Editor): number {
  let n = 0
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'docInlineImage') n++
    return true
  })
  return n
}

async function replace(
  editor: Editor,
  start: number,
  end: number,
  html: string | ((source: DocsQueryResult) => string),
  intent: { kinds?: string[]; imageIndexes?: number[] } = {},
  track?: { author: string },
) {
  const runtime = createDocsRuntime({
    current: () => ({ editor, identity: editor, numIds: NUM_IDS, track }),
  })
  const caller = runtime.connect({ id: 'picture-replacement' })
  try {
    caller.context()
    const read = await caller.execute({
      id: 'read',
      name: 'query_document',
      input: { domain: 'body' },
    })
    expect(read.isError, read.output).not.toBe(true)
    const source = JSON.parse(read.output) as DocsQueryResult
    expect(source.page.complete).toBe(true)
    const structures = source.items
      .filter((item) => item.blockIndex >= start && item.blockIndex <= end)
      .flatMap((item) => {
        expect(item.structuresComplete).toBe(true)
        return item.structures ?? []
      })
    const images = structures.filter((item) => item.kind === 'image')
    return await caller.execute({
      id: 'replace',
      name: 'replace_blocks',
      input: {
        startBlockIndex: start,
        endBlockIndex: end,
        html: typeof html === 'function' ? html(source) : html,
        deleteStructures: [
          ...structures.filter((item) => intent.kinds?.includes(item.kind)).map((item) => item.id),
          ...(intent.imageIndexes ?? []).map((index) => images[index].id),
        ],
      },
    })
  } finally {
    caller.close()
    runtime.dispose()
  }
}

function refs(source: DocsQueryResult): string[] {
  return source.items.flatMap((item) =>
    [...item.html.matchAll(/<img src="([^"]+)"/g)].map((match) => match[1]),
  )
}

setModuleLang('en')

describe('pictures are visible to the model and survive a block rewrite when echoed back', () => {
  it('read_blocks shows a picture-only paragraph as an img ref, not as an empty paragraph', () => {
    const editor = createEditor([
      para([picture()]),
      para([text('Caption: '), picture(), text(' end')]),
    ])
    expect(serializeRangeToHtml(editor, 0, 1)).toBe(
      '<p><img src="doc:0.0"></p>\n<p>Caption: <img src="doc:1.0"> end</p>',
    )
  })

  it('the block list names a picture-only paragraph instead of leaving the preview blank', () => {
    const editor = createEditor([para([picture()]), para([text('Body')])])
    expect(buildDocContext(editor)).toContain('(1 picture, no text)')
  })

  it('a rewrite that omits the picture deletes it and the result says so', async () => {
    const editor = createEditor([para([picture()]), para([text('Body')])])
    const r = await replace(editor, 0, 1, '<p>New intro</p><p>New body</p>', { imageIndexes: [0] })
    expect(r.isError).toBeFalsy()
    expect(r.output).toContain(
      '1 picture(s) that were in the replaced range are not in the new content',
    )
    expect(pictureCount(editor)).toBe(0)
  })

  it('echoing the img ref keeps the picture, placed where the tag sits, with no warning', async () => {
    const editor = createEditor([para([picture()]), para([text('Body')])])
    const r = await replace(
      editor,
      0,
      1,
      (source) => `<p>Figure 1 <img src="${refs(source)[0]}"> shows it</p><p>New body</p>`,
    )
    expect(r.isError).toBeFalsy()
    expect(r.output).not.toContain('picture(s)')
    expect(pictureCount(editor)).toBe(1)
    expect(editor.state.doc.child(0).textContent).toBe('Figure 1  shows it')
    expect(editor.state.doc.child(0).content.child(1).type.name).toBe('docInlineImage')
    expect(editor.state.doc.child(0).content.child(1).attrs.dataUrl).toBe(PNG)
  })

  it('refuses an unknown ref or foreign image instead of silently omitting it', async () => {
    const editor = createEditor([para([text('Body')])])
    const r = await replace(
      editor,
      0,
      0,
      '<p>A <img src="doc:9.9"> B <img src="https://x/y.png"> C</p>',
    )
    expect(r).toMatchObject({ isError: true, mutated: false })
    expect(pictureCount(editor)).toBe(0)
    expect(editor.state.doc.child(0).textContent).toBe('Body')
  })

  it('a ref from outside the replaced range copies that picture in without counting as a loss', async () => {
    const editor = createEditor([para([picture()]), para([text('Body')])])
    const r = await replace(
      editor,
      1,
      1,
      (source) => `<p>Body with <img src="${refs(source)[0]}"></p>`,
    )
    expect(r.isError).toBeFalsy()
    expect(r.output).not.toContain('picture(s)')
    expect(pictureCount(editor)).toBe(2)
  })

  it('a withheld picture gets no ref, and a guessed ref is refused without mutation', async () => {
    const editor = createEditor([
      para([picture([{ type: 'redaction', attrs: { label: 'private' } }]), text(' after')]),
    ])
    expect(serializeRangeToHtml(editor, 0, 0)).not.toContain('doc:')
    const r = await replace(editor, 0, 0, '<p>x <img src="doc:0.0"></p>')
    expect(r).toMatchObject({ isError: true, mutated: false })
    expect(pictureCount(editor)).toBe(1)
  })

  it('a picture pending deletion is neither numbered nor counted as lost', async () => {
    const editor = createEditor([
      para([picture([{ type: 'del', attrs: { author: 'a' } }]), picture(), text(' tail')]),
    ])
    expect(serializeRangeToHtml(editor, 0, 0)).toBe('<p><img src="doc:0.0"> tail</p>')
    const r = await replace(editor, 0, 0, (source) => `<p>kept <img src="${refs(source)[0]}"></p>`)
    expect(r.output).not.toContain('picture(s)')
    const live = editor.state.doc.child(0)
    expect(live.content.child(1).type.name).toBe('docInlineImage')
    expect(live.content.child(1).marks.some((m) => m.type.name === 'del')).toBe(false)
  })

  it('replace_selection accepts an echoed ref instead of throwing', async () => {
    const editor = createEditor([para([text('Hello '), picture(), text(' world')])])
    // select "Hello <pic>" (positions 1..8 inside the paragraph)
    editor.commands.setTextSelection({ from: 1, to: 8 })
    const r = await executeTool(
      editor,
      { id: 't', name: 'replace_selection', input: { html: 'Hi <img src="doc:0.0">' } },
      NUM_IDS,
    )
    expect(r.isError, r.output).toBeFalsy()
    expect(pictureCount(editor)).toBe(1)
    expect(editor.state.doc.child(0).textContent).toBe('Hi  world')
  })

  it('echoing one picture twice does not hide that another was dropped', async () => {
    const editor = createEditor([para([picture(), picture()])])
    const r = await replace(
      editor,
      0,
      0,
      (source) => `<p><img src="${refs(source)[0]}"><img src="${refs(source)[0]}"></p>`,
      { imageIndexes: [1] },
    )
    expect(r.output).toContain('1 picture(s) that were in the replaced range')
    expect(pictureCount(editor)).toBe(2)
  })

  it("a carried-over picture does not strip the selection's inherited run marks from the new text", async () => {
    const editor = createEditor([
      para([
        { type: 'text', text: 'Bold ', marks: [{ type: 'bold' }] },
        picture([{ type: 'docTextStyle', attrs: { color: 'FF0000' } }]),
        { type: 'text', text: ' tail', marks: [{ type: 'bold' }] },
      ]),
    ])
    editor.commands.setTextSelection({ from: 1, to: 7 })
    const r = await executeTool(
      editor,
      { id: 't', name: 'replace_selection', input: { html: 'Strong <img src="doc:0.0">' } },
      NUM_IDS,
    )
    expect(r.isError, r.output).toBeFalsy()
    const first = editor.state.doc.child(0).content.child(0)
    expect(first.text).toBe('Strong ')
    expect(first.marks.some((m) => m.type.name === 'bold')).toBe(true)
  })
})

describe('table formatting inheritance obeys explicit image references', () => {
  it.each(['keep', 'omit', 'move'])(
    'preserves mixed text formatting and links while images %s',
    async (action) => {
      const originalCell = cell([
        { type: 'text', text: 'Alpha   ', marks: [{ type: 'bold' }] },
        {
          type: 'text',
          text: 'Beta',
          marks: [{ type: 'italic' }, { type: 'link', attrs: { href: 'https://example.com' } }],
        },
        picture(),
        text(' end'),
      ])
      const editor = createEditor([table([[originalCell]])])
      const before = editor.state.doc
        .child(0)
        .child(0)
        .child(0)
        .child(0)
        .toJSON()
        .content.filter((node: JsonNode) => node.type !== 'docInlineImage')
      const html = (source: DocsQueryResult) =>
        action === 'move'
          ? `Alpha <img src="${refs(source)[0]}">Beta end`
          : `Alpha Beta${action === 'keep' ? `<img src="${refs(source)[0]}">` : ''} end`
      const result = await replace(
        editor,
        0,
        0,
        (source) => `<table><tr><td>${html(source)}</td></tr></table>`,
        { imageIndexes: action === 'omit' ? [0] : [] },
      )
      expect(result.isError, result.output).toBeFalsy()
      const after = editor.state.doc
        .child(0)
        .child(0)
        .child(0)
        .child(0)
        .toJSON()
        .content.filter((node: JsonNode) => node.type !== 'docInlineImage')
      expect(after).toEqual(before)
      expect(pictureCount(editor)).toBe(action === 'omit' ? 0 : 1)
    },
  )

  it('keeps non-referenceable rule decorations in an unchanged cell', async () => {
    const editor = createEditor([
      table([
        [
          cell([
            text('Label'),
            {
              type: 'docInlineImage',
              attrs: {
                dataUrl: '',
                xml: '<w:pict/>',
                rule: { colorHex: 'A0A0A0', thicknessPx: 2 },
              },
            },
          ]),
        ],
      ]),
    ])
    const before = editor.state.doc.child(0).child(0).child(0).toJSON()
    const result = await replace(editor, 0, 0, (source) => source.items[0].html)
    expect(result.isError, result.output).not.toBe(true)
    expect(editor.state.doc.child(0).child(0).child(0).toJSON()).toEqual(before)
  })

  it('preserves an atomic ruby display beside an explicitly retained image', async () => {
    const editor = createEditor([
      table([[cell([{ type: 'docRuby', attrs: { base: '字', guide: 'zi' } }, picture()])]]),
    ])
    const result = await replace(editor, 0, 0, (source) => source.items[0].html)
    expect(result.isError, result.output).toBeFalsy()
    const paragraph = editor.state.doc.child(0).child(0).child(0).child(0)
    expect(paragraph.child(0).type.name).toBe('docRuby')
    expect(paragraph.child(1).type.name).toBe('docInlineImage')
  })

  it('refuses actual ruby loss even when its neighboring image is retained', async () => {
    const editor = createEditor([
      table([
        [cell([{ type: 'docRuby', attrs: { base: 'ruby base', guide: 'guide' } }, picture()])],
      ]),
    ])
    const before = editor.getJSON()
    const result = await replace(
      editor,
      0,
      0,
      (source) => `<table><tr><td>Changed <img src="${refs(source)[0]}"></td></tr></table>`,
    )
    expect(result).toMatchObject({ isError: true, committed: false })
    expect(editor.getJSON()).toEqual(before)
  })

  it('preserves horizontal spans when an explicit image dissolves a vertical merge', async () => {
    const editor = createEditor([
      table([
        [
          cell([picture()], { rowspan: 2, colspan: 2, colwidth: [100, 100] }),
          cell([text('Right one')]),
        ],
        [cell([text('Right two')])],
      ]),
    ])
    const result = await replace(
      editor,
      0,
      0,
      (source) =>
        `<table><tr><td></td><td>Right one</td></tr><tr><td><img src="${refs(source)[0]}"></td><td>Right two</td></tr></table>`,
      { kinds: ['table'] },
    )
    expect(result.isError, result.output).toBeFalsy()
    const next = editor.state.doc.child(0)
    expect(TableMap.get(next).width).toBe(3)
    next.forEach((row) => {
      expect(row.childCount).toBe(2)
      expect(row.child(0).attrs).toMatchObject({ colspan: 2, rowspan: 1, colwidth: [100, 100] })
    })
    expect(next.child(1).child(1).textContent).toBe('Right two')
    expect(pictureCount(editor)).toBe(1)
  })

  it('round trips images and text across multiple paragraphs in the same cell', async () => {
    const originalCell = cell([])
    originalCell.content = [para([text('A '), picture()]), para([text('B '), picture()])]
    const editor = createEditor([table([[originalCell]])])
    const html = serializeRangeToHtml(editor, 0, 0)
    expect(html).toBe(
      '<table><tr><td data-row="0" data-col="0">A <img src="doc:0.0"><br>B <img src="doc:0.1"></td></tr></table>',
    )
    const result = await replace(editor, 0, 0, (source) => source.items[0].html)
    expect(result.output).not.toContain('picture(s)')
    expect(pictureCount(editor)).toBe(2)
    expect(editor.state.doc.child(0).child(0).child(0).childCount).toBe(2)
    expect(serializeRangeToHtml(editor, 0, 0)).toBe(html)
  })

  it('reads stable image references across table cells and excludes deleted or withheld images', () => {
    const editor = createEditor([
      table([
        [cell([text('Old label')]), cell([picture(), picture([{ type: 'del' }])])],
        [cell([picture([{ type: 'redaction', attrs: { label: 'private' } }])]), cell([picture()])],
      ]),
    ])
    const html = serializeRangeToHtml(editor, 0, 0)
    expect(html).toContain('<td data-row="0" data-col="1"><img src="doc:0.0"></td>')
    expect(html).toContain('<td data-row="1" data-col="1"><img src="doc:0.1"></td>')
    expect(html.match(/<img /g)).toHaveLength(2)
  })

  it('numbers only images in live table rows when resolving an explicit reference', async () => {
    const deleted = picture()
    deleted.attrs = { ...deleted.attrs, widthPx: 11 }
    const live = picture()
    live.attrs = { ...live.attrs, widthPx: 23 }
    const original = table([[cell([deleted])], [cell([live])]])
    original.content![0].attrs = { rowRevision: { kind: 'del', author: 'Reviewer' } }
    const editor = createEditor([original])
    expect(countInlineImages(editor, 0, 0)).toBe(1)
    expect(serializeRangeToHtml(editor, 0, 0)).toBe(
      '<table><tr><td data-row="1" data-col="0"><img src="doc:0.0"></td></tr></table>',
    )
    const result = await replace(
      editor,
      0,
      0,
      (source) => `<table><tr><td><img src="${refs(source)[0]}"></td></tr></table>`,
      { kinds: ['table'] },
    )
    expect(result.output).not.toContain('picture(s)')
    expect(pictureCount(editor)).toBe(1)
    expect(editor.state.doc.child(0).child(0).child(0).child(0).child(0).attrs.widthPx).toBe(23)
  })

  it.each([1, 3])(
    'omitting images removes them from %i replacement row(s) and agrees with the receipt',
    async (rows) => {
      const editor = createEditor([table([[cell([text('Old label')]), cell([picture()])]])])
      buildDocContext(editor)
      const result = await replace(editor, 0, 0, parameterTable(rows), {
        kinds: rows === 1 ? [] : ['table'],
        imageIndexes: [0],
      })
      expect(result.isError, result.output).toBeFalsy()
      expect(pictureCount(editor)).toBe(0)
      expect(result.output).toContain('1 picture(s) that were in the replaced range')
      expect(editor.state.doc.child(0).childCount).toBe(rows)
    },
  )

  it('moves an explicit image into a different empty cell instead of copying old cell content', async () => {
    const editor = createEditor([table([[cell([]), cell([picture()])]])])
    const result = await replace(
      editor,
      0,
      0,
      (source) => `<table><tr><td><img src="${refs(source)[0]}"></td><td></td></tr></table>`,
    )
    expect(result.output).not.toContain('picture(s)')
    expect(pictureCount(editor)).toBe(1)
    expect(editor.state.doc.child(0).child(0).child(0).child(0).child(0).type.name).toBe(
      'docInlineImage',
    )
    expect(editor.state.doc.child(0).child(0).child(1).child(0).childCount).toBe(0)
  })

  it('keeps one explicit image without cloning it into additional rows', async () => {
    const editor = createEditor([table([[cell([text('Old label')]), cell([picture()])]])])
    const result = await replace(
      editor,
      0,
      0,
      (source) => parameterTable(3, `<img src="${refs(source)[0]}">`),
      { kinds: ['table'] },
    )
    expect(result.isError, result.output).not.toBe(true)
    expect(editor.state.doc.child(0).childCount).toBe(3)
    expect(result.output).not.toContain('picture(s)')
    expect(pictureCount(editor)).toBe(1)
    expect(editor.state.doc.child(0).child(0).child(1).child(0).child(0).attrs.dataUrl).toBe(PNG)
  })

  it('refuses an unknown reference in a blank cell without changing the original table', async () => {
    const editor = createEditor([table([[cell([text('Old label')]), cell([picture()])]])])
    const result = await replace(editor, 0, 0, parameterTable(1, '<img src="doc:99.0">'))
    expect(result).toMatchObject({ isError: true, mutated: false })
    expect(pictureCount(editor)).toBe(1)
    expect(editor.state.doc.child(0).textContent).toContain('Old label')
  })

  it('does not restore an omitted image when a different image is explicitly duplicated', async () => {
    const editor = createEditor([table([[cell([picture()]), cell([picture()])]])])
    const result = await replace(
      editor,
      0,
      0,
      (source) =>
        `<table><tr><td><img src="${refs(source)[0]}"><img src="${refs(source)[0]}"></td><td></td></tr></table>`,
      { imageIndexes: [1] },
    )
    expect(pictureCount(editor)).toBe(2)
    expect(result.output).toContain('1 picture(s) that were in the replaced range')
  })

  it('omits the spanning image when a vertically merged template expands into plain rows', async () => {
    const editor = createEditor([
      table([[cell([text('First')]), cell([picture()], { rowspan: 2 })], [cell([text('Second')])]]),
    ])
    const result = await replace(editor, 0, 0, parameterTable(3), {
      kinds: ['table'],
      imageIndexes: [0],
    })
    expect(result.isError, result.output).toBeFalsy()
    expect(pictureCount(editor)).toBe(0)
    editor.state.doc.child(0).forEach((row) => {
      expect(row.childCount).toBe(2)
      row.forEach((nextCell) => expect(nextCell.attrs.rowspan).toBe(1))
    })
  })

  it('honors an explicit image moved into a formerly merged-away slot', async () => {
    const editor = createEditor([
      table([[cell([text('First')]), cell([picture()], { rowspan: 2 })], [cell([text('Second')])]]),
    ])
    const result = await replace(
      editor,
      0,
      0,
      (source) =>
        `<table><tr><td>First</td><td></td></tr><tr><td>Second</td><td><img src="${refs(source)[0]}"></td></tr></table>`,
      { kinds: ['table'] },
    )
    expect(result.isError, result.output).toBeFalsy()
    expect(pictureCount(editor)).toBe(1)
    expect(editor.state.doc.child(0).child(1).childCount).toBe(2)
    expect(editor.state.doc.child(0).child(1).child(1).child(0).child(0).type.name).toBe(
      'docInlineImage',
    )
    expect(result.output).not.toContain('picture(s)')
  })

  it.each(['accept', 'reject'])(
    'keeps omitted images only in the deleted revision until %s',
    async (decision) => {
      const editor = createEditor([table([[cell([text('Old label')]), cell([picture()])]])])
      const result = await replace(
        editor,
        0,
        0,
        parameterTable(3),
        { kinds: ['table'], imageIndexes: [0] },
        { author: 'Reviewer' },
      )
      expect(result.isError, result.output).toBeFalsy()
      expect(result.output).toContain('(as a tracked change the user can reject)')
      expect(editor.state.doc.child(0).attrs.blockRevision.kind).toBe('del')
      expect(editor.state.doc.child(1).attrs.blockRevision.kind).toBe('ins')
      expect(countInlineImages(editor, 0, editor.state.doc.childCount - 1)).toBe(0)
      expect(pictureCount(editor)).toBe(1)
      if (decision === 'accept') acceptAllRevisions(editor)
      else rejectAllRevisions(editor)
      expect(pictureCount(editor)).toBe(decision === 'accept' ? 0 : 1)
    },
  )
})
