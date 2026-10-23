import { afterEach, describe, expect, it } from 'vitest'
import type { Editor } from '@tiptap/core'
import {
  buildBlankDocx,
  PAGE_MARK,
  parseDocx,
  saveDocx,
  TOTAL_PAGES_MARK,
  type HeaderFooter,
  type SaveBlock,
} from '@genoffice/docx-engine'
import {
  atHfCellEdge,
  hfDocJson,
  hfValueFromDoc,
  insertHfField,
  insertHfInstrField,
  mountHfEditor,
  type HfEditorHandle,
} from '../src/renderer/editor/hf-editor'
import type { PmNode } from '../src/renderer/editor/convert'

const handles: HfEditorHandle[] = []
afterEach(() => {
  for (const h of handles.splice(0)) {
    if (!h.editor.isDestroyed) h.exit()
  }
})

function mount(value: HeaderFooter, onCommit: (next: HeaderFooter) => void = () => {}) {
  const host = document.createElement('div')
  document.body.append(host)
  const handle = mountHfEditor(host, {
    value,
    pageNo: '3',
    pageTotal: '9',
    spellcheck: false,
    onCommit,
    onExit: () => host.remove(),
  })
  handles.push(handle)
  return { host, handle, editor: handle.editor as Editor }
}

const RICH: HeaderFooter = {
  text: '',
  pageNumber: true,
  paras: [
    { runs: [], cells: [{ paras: [[{ text: 'logo row' }]] }] },
    {
      align: 'right',
      tabStops: [{ pos: 4680, val: 'right' }],
      runs: [
        { text: 'Report', bold: true },
        { text: '\tPage ' },
        { text: PAGE_MARK, italic: true },
        { text: ` of ${TOTAL_PAGES_MARK}` },
      ],
    },
    { align: 'center', runs: [{ text: 'second line', color: 'FF0000' }] },
  ],
}

describe('hfDocJson / hfValueFromDoc', () => {
  it('maps text paragraphs to editor nodes with field atoms and rows to cell nodes', () => {
    const doc = hfDocJson(RICH)
    expect(doc.content).toHaveLength(3)
    const row = doc.content![0]
    expect(row.type).toBe('hfRow')
    expect(row.attrs).toMatchObject({ hfIndex: 0 })
    expect(row.content![0].type).toBe('hfCell')
    expect(row.content![0].content![0]).toMatchObject({ type: 'docParagraph' })
    expect(row.content![0].content![0].content![0]).toMatchObject({ text: 'logo row' })
    const first = doc.content![1]
    expect(first.attrs).toMatchObject({ align: 'right', hfIndex: 1 })
    expect(JSON.parse(String(first.attrs!.tabStops))).toEqual([{ pos: 4680, val: 'right' }])
    const types = (first.content ?? []).map((n) => n.type)
    expect(types).toEqual(['text', 'text', 'hfField', 'text', 'hfField'])
    const fields = (first.content ?? []).filter((n) => n.type === 'hfField')
    expect(fields.map((n) => n.attrs?.kind)).toEqual(['PAGE', 'NUMPAGES'])
    // the sentinel's own formatting stays on the atom
    expect(fields[0].marks).toEqual([{ type: 'italic' }])
  })

  it('round-trips unchanged content and splices the row back at its slot', () => {
    const back = hfValueFromDoc(hfDocJson(RICH), RICH)
    expect(back.paras).toHaveLength(3)
    expect(back.paras![0].cells).toBeTruthy()
    expect(back.paras![1]).toMatchObject({
      align: 'right',
      tabStops: [{ pos: 4680, val: 'right' }],
    })
    expect(back.paras![1].runs.map((r) => r.text).join('')).toBe(
      `Report\tPage ${PAGE_MARK} of ${TOTAL_PAGES_MARK}`,
    )
    expect(back.paras![1].runs[0]).toMatchObject({ text: 'Report', bold: true })
    expect(back.paras![1].runs.find((r) => r.text === PAGE_MARK)).toMatchObject({ italic: true })
    expect(back.paras![2].runs[0]).toMatchObject({ text: 'second line', color: 'FF0000' })
    expect(back.pageNumber).toBe(true)
  })

  it('a new paragraph split from an existing one inherits that paragraph as its template', () => {
    const doc = hfDocJson(RICH)
    const [row, p1, p2] = doc.content!
    const extra: PmNode = {
      type: 'docParagraph',
      attrs: { ...p1.attrs },
      content: [{ type: 'text', text: 'new' }],
    }
    const back = hfValueFromDoc({ type: 'doc', content: [row, p1, extra, p2] }, RICH)
    expect(back.paras!.map((p) => p.align)).toEqual([undefined, 'right', 'right', 'center'])
    expect(back.paras![2].runs[0].text).toBe('new')
    expect(back.text).toBe(`Report\tPage ${PAGE_MARK} of ${TOTAL_PAGES_MARK}newsecond line`)
  })

  it('an untouched row round-trips without a rewrite flag; an edited cell gets one', () => {
    const same = hfValueFromDoc(hfDocJson(RICH), RICH)
    expect(same.paras![0].cells![0].rewrite).toBeUndefined()
    expect(same.paras![0].cells![0].paras).toEqual([[{ text: 'logo row' }]])

    const doc = hfDocJson(RICH)
    const cell = doc.content![0].content![0]
    cell.content![0].content = [{ type: 'text', text: 'Company ', marks: [{ type: 'bold' }] }]
    cell.content!.push({
      type: 'docParagraph',
      attrs: {},
      content: [{ type: 'text', text: 'Ltd' }],
    })
    const back = hfValueFromDoc(doc, RICH)
    const edited = back.paras![0].cells![0]
    expect(edited.rewrite).toBe(true)
    expect(edited.paras).toEqual([[{ text: 'Company ', bold: true }], [{ text: 'Ltd' }]])
    // the row keeps its slot ahead of the text paragraphs
    expect(back.paras![1].runs[0]).toMatchObject({ text: 'Report', bold: true })
    expect(back.text).toBe(`Report\tPage ${PAGE_MARK} of ${TOTAL_PAGES_MARK}second line`)
  })

  it('Backspace at a cell start and Delete at a cell end are swallowed, inner joins still work', () => {
    const { editor } = mount(RICH)
    const posOf = (text: string, atEnd = false) => {
      let pos = -1
      editor.state.doc.descendants((n, p) => {
        if (pos < 0 && n.isTextblock && n.textContent === text)
          pos = atEnd ? p + 1 + n.content.size : p + 1
        return pos < 0
      })
      return pos
    }
    const key = (k: 'Backspace' | 'Delete') =>
      editor.view.someProp('handleKeyDown', (f) =>
        f(editor.view, new KeyboardEvent('keydown', { key: k, cancelable: true })),
      )
    editor.commands.setTextSelection(posOf('logo row'))
    const before = JSON.stringify(editor.getJSON())
    expect(atHfCellEdge(editor.state, -1)).toBe(true)
    expect(key('Backspace')).toBe(true)
    editor.commands.setTextSelection(posOf('logo row', true))
    expect(atHfCellEdge(editor.state, 1)).toBe(true)
    expect(key('Delete')).toBe(true)
    expect(JSON.stringify(editor.getJSON())).toBe(before)
    // two paragraphs inside one cell still join on Backspace
    editor.commands.splitBlock()
    editor.commands.insertContent('more')
    editor.commands.setTextSelection(posOf('more'))
    expect(atHfCellEdge(editor.state, -1)).toBe(false)
    expect(editor.commands.keyboardShortcut('Backspace')).toBe(true)
    const cell = (editor.getJSON() as PmNode).content![0].content![0]
    expect(cell.content).toHaveLength(1)
    expect(cell.content![0].content![0].text).toBe('logo rowmore')
  })

  it('an untouched picture cell is not flagged when another cell changes; a format change is', () => {
    const picture = {
      dataUrl: 'data:image/png;base64,AA==',
      widthPx: 40,
      heightPx: 20,
      xml: '<w:drawing>logo</w:drawing>',
    }
    const value: HeaderFooter = {
      text: '',
      paras: [
        {
          runs: [],
          cells: [
            { paras: [[{ text: '', image: picture }]], widthTwips: 2835 },
            {
              paras: [[{ text: 'Company' }], [{ text: 'Sub' }]],
              paraProps: [undefined, { spaceAfter: 40 }],
              widthTwips: 7370,
            },
          ],
        },
      ],
    }
    let committed: HeaderFooter | null = null
    const { editor, handle } = mount(value, (next) => (committed = next))
    // the loaded doc already carries the schema's non-null defaults (leadRule: false ...)
    expect(handle.changed()).toBeNull()
    let pos = -1
    editor.state.doc.descendants((n, p) => {
      if (pos < 0 && n.isTextblock && n.textContent === 'Company') pos = p + 1 + n.content.size
      return pos < 0
    })
    editor.commands.setTextSelection(pos)
    editor.commands.insertContent(' Ltd')
    const next = handle.changed()!
    expect(next.paras![0].cells![0].rewrite).toBeUndefined()
    expect(next.paras![0].cells![1].rewrite).toBe(true)
    expect(next.paras![0].cells![1].rewriteFormat).toBeUndefined()
    expect(next.paras![0].cells![1].paras[0][0].text).toBe('Company Ltd')
    // ribbon alignment on the second cell paragraph: that paragraph's pPr is rebuilt on save
    let subPos = -1
    editor.state.doc.descendants((n, p) => {
      if (subPos < 0 && n.isTextblock && n.textContent === 'Sub') subPos = p
      return subPos < 0
    })
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(subPos, undefined, {
        ...editor.state.doc.nodeAt(subPos)!.attrs,
        align: 'right',
      }),
    )
    const formatted = handle.changed()!
    const cell = formatted.paras![0].cells![1]
    expect(cell.rewriteFormat).toEqual([1])
    expect(cell.paraProps![1]).toEqual({ spaceAfter: 40, align: 'right' })
    expect(cell.priorParaProps).toEqual([undefined, { spaceAfter: 40 }])
    handle.exit()
    expect(committed).not.toBeNull()
  })

  it('a locked row (flattened nested table) stays out of the editor and in its slot', () => {
    const locked: HeaderFooter = {
      text: '',
      paras: [
        { runs: [], cells: [{ paras: [[{ text: 'nested' }]], locked: true }] },
        { runs: [{ text: 'line' }] },
      ],
    }
    const doc = hfDocJson(locked)
    expect(doc.content!.map((n) => n.type)).toEqual(['docParagraph'])
    const back = hfValueFromDoc(doc, locked)
    expect(back.paras![0].cells![0].locked).toBe(true)
    expect(back.paras![1].runs[0].text).toBe('line')
  })

  it('an empty legacy value edits as one paragraph; a typed PAGE field sets pageNumber', () => {
    const doc = hfDocJson({ text: '' })
    expect(doc.content).toHaveLength(1)
    const back = hfValueFromDoc(
      {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { align: 'center' },
            content: [{ type: 'hfField', attrs: { kind: 'PAGE' } }],
          },
        ],
      },
      { text: '' },
    )
    expect(back.paras![0].runs[0].text).toBe(PAGE_MARK)
    expect(back.pageNumber).toBe(true)
  })
})

describe('mountHfEditor', () => {
  it('renders field atoms with the live numbers and commits typed formatting on exit', () => {
    const commits: HeaderFooter[] = []
    const { host, handle, editor } = mount(RICH, (next) => commits.push(next))
    const fields = host.querySelectorAll('.hf-field')
    expect(Array.from(fields, (f) => f.getAttribute('data-hf-num'))).toEqual(['3', '9'])
    editor.commands.focus('end')
    editor.chain().insertContent(' tail').setMark('bold').run()
    editor.commands.insertContent('X')
    expect(
      handle
        .changed()
        ?.paras?.[2].runs.map((r) => r.text)
        .join(''),
    ).toBe('second line tailX')
    handle.exit()
    expect(commits).toHaveLength(1)
    const last = commits[0].paras![2]
    expect(last.runs.at(-1)).toMatchObject({ text: 'X', bold: true })
    expect(editor.isDestroyed).toBe(true)
  })

  it('does not commit when nothing changed', () => {
    let commits = 0
    const { handle } = mount(RICH, () => commits++)
    handle.exit()
    expect(commits).toBe(0)
  })

  it('stays open when its strip is detached under the caret, exits on a real blur', async () => {
    const { host, editor } = mount(RICH)
    const blur = () => editor.view.dom.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    // Blink blurs the focused element before removing it: the host is still connected
    blur()
    host.remove()
    await Promise.resolve()
    expect(editor.isDestroyed).toBe(false)
    document.body.append(host)
    blur()
    expect(editor.isDestroyed).toBe(false)
    await Promise.resolve()
    expect(editor.isDestroyed).toBe(true)
  })

  it('inserts PAGE / NUMPAGES atoms and DATE fields through the ribbon helpers', () => {
    const commits: HeaderFooter[] = []
    const { handle, editor } = mount({ text: '' }, (next) => commits.push(next))
    insertHfField(editor, 'PAGE')
    editor.commands.insertContent(' / ')
    insertHfField(editor, 'NUMPAGES')
    insertHfInstrField(editor, 'DATE', '2026-09-23')
    handle.exit()
    const runs = commits[0].paras![0].runs
    expect(runs.map((r) => r.text).join('')).toBe(`${PAGE_MARK} / ${TOTAL_PAGES_MARK}2026-09-23`)
    expect(runs.at(-1)).toMatchObject({ text: '2026-09-23', instrField: 'DATE' })
    expect(commits[0].pageNumber).toBe(true)
  })
})

describe('edited header saves with its formatting', () => {
  it('bold run, tab, PAGE and NUMPAGES fields survive save and reparse', async () => {
    const parsed = await parseDocx(await buildBlankDocx())
    const saveBlocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    let committed: HeaderFooter | null = null
    const { handle, editor } = mount({ text: '' }, (next) => (committed = next))
    editor.commands.focus('end')
    editor.commands.insertContent([
      { type: 'text', text: 'Title', marks: [{ type: 'bold' }] },
      { type: 'text', text: '\tp. ' },
    ])
    insertHfField(editor, 'PAGE')
    handle.exit()
    const saved = await saveDocx(parsed, saveBlocks, { header: committed! })
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const hdr = await zip.file('word/header1.xml')!.async('string')
    expect(hdr).toMatch(/<w:r><w:rPr><w:b\/>.*<\/w:rPr><w:t[^>]*>Title<\/w:t>/)
    expect(hdr).toContain('<w:tab/>')
    expect(hdr).toContain('<w:instrText xml:space="preserve"> PAGE </w:instrText>')
    const reparsed = await parseDocx(saved)
    expect(reparsed.headerParas![0].runs[0]).toMatchObject({ text: 'Title', bold: true })
    expect(reparsed.headerHasPageNumber).toBe(true)
  })
})
