import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { generateCaptionXml, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { allowDirtyFields } from '../../../packages/docx-engine/tests/helpers/dirty-fields'
import { executeOps } from '../src/renderer/ai/ops'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { collectFields } from '../src/renderer/editor/field-codes'
import {
  footnoteOfBookmark,
  hasWaitingPageField,
  refreshPageFields,
  replaceFieldRuns,
  settleFieldResults,
  type PageFacts,
  type PageSource,
} from '../src/renderer/editor/field-results'
import { withHfFieldResults } from '../src/renderer/hf-sections'
import { pageAtCanvasY, type PageCut } from '../src/renderer/pagination-cuts'

type Parsed = Awaited<ReturnType<typeof parseDocx>>

const p = (inner: string) =>
  `<w:p>${inner.startsWith('<') ? inner : `<w:r><w:t xml:space="preserve">${inner}</w:t></w:r>`}</w:p>`
const t = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`

const field = (instr: string, result: string, dirty = false) =>
  `<w:r><w:fldChar w:fldCharType="begin"${dirty ? ' w:dirty="true"' : ''}/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  return { editor, parsed }
}

async function save(editor: Editor, parsed: Parsed, pages?: PageSource | null) {
  settleFieldResults(editor, parsed.blocks, pages)
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

/** each page field of the saved body: `KEYWORD=result`, with ` dirty` where it is marked */
const results = (xml: string): string[] =>
  [
    ...xml.matchAll(
      /<w:fldChar w:fldCharType="begin"( w:dirty="true")?\/><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:instrText[^>]*> (PAGEREF|PAGE|NUMPAGES|SECTIONPAGES)\b[^<]*<\/w:instrText>.*?<w:t[^>]*>([^<]*)<\/w:t>/g,
    ),
  ].map((m) => `${m[2]}=${m[3]}${m[1] ? ' dirty' : ''}`)

const page = (index: number, more: Partial<PageFacts> = {}): PageFacts => ({
  index,
  number: index + 1,
  display: String(index + 1),
  chaptered: false,
  sectionPages: 1,
  ...more,
})

/** a layout that puts each top-level block on the page `pageOfBlock` names (null: it cannot tell) */
function layout(
  editor: Editor,
  pageOfBlock: (block: number, edge?: 'top' | 'bottom') => PageFacts | null,
  more: Partial<PageSource> = {},
): PageSource & { asked: Array<[number, string]> } {
  const asked: Array<[number, string]> = []
  return {
    asked,
    pageCount: 9,
    edited: true,
    words: { above: 'above', below: 'below', onPage: (n) => `on page ${n}` },
    pageAt: (pos, edge) => {
      asked.push([pos, edge ?? 'char'])
      return pageOfBlock(editor.state.doc.resolve(pos).index(0), edge)
    },
    ...more,
  }
}

const BODY = p('Target paragraph.') + p('Back to ') + p('Page ') + p('Of ') + p('In section ')

async function withPageFields() {
  const opened = await open(BODY)
  const outcome = executeOps(opened.editor, [
    { op: 'insertBookmark', target: { blockIndexes: [0] }, name: 'Target' },
    { op: 'insertField', target: { blockIndexes: [1] }, type: 'PAGEREF', args: 'Target \\h' },
    { op: 'insertField', target: { blockIndexes: [2] }, type: 'PAGE' },
    { op: 'insertField', target: { blockIndexes: [3] }, type: 'NUMPAGES', args: '\\* ROMAN' },
    { op: 'insertField', target: { blockIndexes: [4] }, type: 'SECTIONPAGES' },
  ])
  expect(outcome.ok, outcome.error).toBe(true)
  return opened
}

describe('page fields are saved with the page the layout shows, without w:dirty', () => {
  it('the page of the bookmark, of the field, the page count and the section count', async () => {
    const { editor, parsed } = await withPageFields()
    const pages = layout(editor, (block) => page(block < 2 ? 2 : 6, { sectionPages: 4 }))
    const xml = await save(editor, parsed, pages)
    expect(results(xml)).toEqual(['PAGEREF=3', 'PAGE=7', 'NUMPAGES=IX', 'SECTIONPAGES=4'])
    expect(xml).not.toContain('w:dirty')
  })

  it('without a layout (a pass is due, Draft, Read Mode) the fields wait for Word and nothing is written', async () => {
    const { editor, parsed } = await withPageFields()
    const before = editor.state.doc
    allowDirtyFields()
    const xml = await save(editor, parsed, null)
    expect(editor.state.doc.eq(before)).toBe(true)
    expect(results(xml)).toEqual([
      'PAGEREF=1 dirty',
      'PAGE=1 dirty',
      'NUMPAGES=1 dirty',
      'SECTIONPAGES=1 dirty',
    ])
  })

  it('a save asks for a pass only while a page field waits for its page', async () => {
    const { editor, parsed } = await withPageFields()
    expect(hasWaitingPageField(editor.state.doc)).toBe(true)
    await save(
      editor,
      parsed,
      layout(editor, () => page(0)),
    )
    expect(hasWaitingPageField(editor.state.doc)).toBe(false)
    // a field of another kind that waits for Word asks for none
    const other = await open(p(t('Made ') + field('CREATEDATE', 'x', true)))
    expect(hasWaitingPageField(other.editor.state.doc)).toBe(false)
  })

  it('one place the layout cannot tell leaves that field marked, and only that one', async () => {
    const { editor, parsed } = await withPageFields()
    // the PAGE field's paragraph sits in the gap between two sheets
    const pages = layout(editor, (block) => (block === 2 ? null : page(block)))
    allowDirtyFields()
    const xml = await save(editor, parsed, pages)
    expect(results(xml)).toEqual(['PAGEREF=1', 'PAGE=1 dirty', 'NUMPAGES=IX', 'SECTIONPAGES=1'])
  })
})

describe('the number is the one the page shows', () => {
  const one = (instr: string, facts: PageFacts) => async () => {
    const { editor } = await open(p(t('Page ') + field(instr, '?', true)))
    const tr = editor.state.tr
    refreshPageFields(
      tr,
      layout(editor, () => facts),
    )
    const [range] = collectFields(tr.doc)
    return `${tr.doc.textBetween(range.from, range.to)}${range.mark.attrs.dirty ? ' dirty' : ''}`
  }
  const roman = page(2, { display: 'iii' })
  const chapter = page(4, { display: '2-5', chaptered: true })

  it('the format of the section, and a restart: the displayed number, not the count of sheets', async () => {
    expect(await one('PAGE', roman)()).toBe('iii')
    expect(await one('PAGE', page(7, { number: 1, display: '1' }))()).toBe('1')
    expect(await one('PAGE \\* MERGEFORMAT', roman)()).toBe('iii')
    expect(await one('PAGE \\* Upper', roman)()).toBe('III')
  })

  it('a format switch of the field itself wins over the section', async () => {
    expect(await one('PAGE \\* Arabic', roman)()).toBe('3')
    expect(await one('PAGE \\* ROMAN', page(3))()).toBe('IV')
    expect(await one('PAGE \\# "00"', page(3))()).toBe('04')
  })

  it('a chapter number is part of the page; with a format switch on top it stays for Word', async () => {
    expect(await one('PAGE', chapter)()).toBe('2-5')
    expect(await one('PAGE \\* roman', chapter)()).toBe('? dirty')
  })

  it('the counting formats of the standard and the worded ones are written', async () => {
    expect(await one('PAGE \\* CHINESENUM3', page(3))()).toBe('\u56db')
    expect(await one('PAGE \\* DBNUM1', page(11))()).toBe('\u4e00\u4e8c')
    expect(await one('PAGE \\* ZODIAC3', page(0))()).toBe('\u7532\u5b50')
    expect(await one('NUMPAGES \\* CardText', page(3))()).toBe('nine')
    expect(await one('PAGE \\* Ordinal', page(1))()).toBe('2nd')
    expect(await one('PAGE \\* ArabicDash', page(1))()).toBe('- 2 -')
    expect(await one('PAGE \\* Roman', page(3))()).toBe('IV')
    expect(await one('PAGE \\* Hex', page(254))()).toBe('FF')
  })

  it('a format this editor has no rule for, or a value outside its rule, stays for Word', async () => {
    expect(await one('PAGE \\* KANJINUM3', page(3))()).toBe('? dirty')
    expect(await one('PAGE \\* IROHA', page(3))()).toBe('? dirty')
    // past the pairs of the cycle, and past what the standard shows of the counting system
    expect(await one('PAGE \\* ZODIAC3', page(60))()).toBe('? dirty')
    expect(await one('PAGE \\* CHINESENUM1', page(99))()).toBe('? dirty')
    expect(await one('NUMPAGES \\* DBCHAR \\* roman', page(3))()).toBe('? dirty')
    expect(await one('NUMPAGES \\* DBCHAR \\# 00', page(3))()).toBe('? dirty')
  })

  it('DBCHAR writes full-width digits, SBCHAR plain ones, and neither a chapter number', async () => {
    expect(await one('NUMPAGES \\* DBCHAR', page(3))()).toBe('\uff19')
    expect(await one('PAGE \\* DBCHAR \\* MERGEFORMAT', page(11))()).toBe('\uff11\uff12')
    expect(await one('PAGE \\* SBCHAR', page(11, { display: 'xii' }))()).toBe('12')
    const chapter = page(0, { display: '3-1', chaptered: true })
    expect(await one('PAGE \\* DBCHAR', chapter)()).toBe('? dirty')
  })
})

describe('PAGEREF', () => {
  const cell = (xml: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${xml}</w:tc>`
  const TABLE =
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr>${cell(p(t('Row one ') + field('PAGE', '?', true)))}</w:tr>` +
    `<w:tr>${cell(`<w:p><w:bookmarkStart w:id="5" w:name="InCell"/>${t('Row two ')}<w:bookmarkEnd w:id="5"/>${field('PAGE', '?', true)}</w:p>`)}</w:tr>` +
    '</w:tbl>'

  it('a field and a bookmark in a table that runs over pages are asked where they are, not where the table starts', async () => {
    const { editor, parsed } = await open(
      TABLE + p(t('See ') + field('PAGEREF InCell \\h', '?', true)),
    )
    const table = editor.state.doc.firstChild!
    const rowTwo = 1 + table.firstChild!.nodeSize
    // the table opens on page 1; its second row is on page 2
    const pages: PageSource = {
      pageCount: 2,
      edited: true,
      pageAt: (pos) => page(pos < table.nodeSize && pos >= rowTwo ? 1 : 0),
    }
    const xml = await save(editor, parsed, pages)
    expect(results(xml)).toEqual(['PAGE=1', 'PAGE=2', 'PAGEREF=2'])
    expect(xml).not.toContain('w:dirty')
  })

  it('a bookmark in a text box is on the page of the block that holds the box', async () => {
    const box =
      '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      '<mc:Choice Requires="wps"><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/><wp:extent cx="3000000" cy="1700000"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1700000"/></a:xfrm></wps:spPr><wps:bodyPr/>' +
      `<wps:txbx><w:txbxContent><w:p><w:bookmarkStart w:id="7" w:name="Boxed"/>${t('In a box')}<w:bookmarkEnd w:id="7"/></w:p></w:txbxContent></wps:txbx>` +
      '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'
    const { editor, parsed } = await open(
      p('First page') + box + p(t('See ') + field('PAGEREF Boxed \\h', '?', true)),
    )
    // the block of the box sits on the second page, whole
    const pages = layout(editor, (block) => page(block === 1 ? 1 : 2))
    const xml = await save(editor, parsed, pages)
    expect(results(xml)).toEqual(['PAGEREF=2'])
    expect(xml).not.toContain('w:dirty')
  })

  /** a text box whose paragraph holds `inner` */
  const boxOf = (inner: string) =>
    '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
    '<mc:Choice Requires="wps"><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/><wp:extent cx="3000000" cy="1700000"/>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1700000"/></a:xfrm></wps:spPr><wps:bodyPr/>' +
    `<wps:txbx><w:txbxContent><w:p>${inner}</w:p></w:txbxContent></wps:txbx>` +
    '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'
  const BOXED =
    `<w:p><w:bookmarkStart w:id="3" w:name="Target"/>${t('Target')}<w:bookmarkEnd w:id="3"/></w:p>` +
    boxOf(
      t('Page ') +
        field('PAGE', '?', true) +
        t(' of ') +
        field('NUMPAGES', '?', true) +
        t(', see ') +
        field('PAGEREF Target \\h', '?', true) +
        t(' ') +
        field('PAGEREF Target \\p', '?', true) +
        t(' kept ') +
        field('PAGE', '8'),
    )

  it('page fields waiting in a text box take the page of the block that holds the box', async () => {
    const { editor, parsed } = await open(BOXED)
    expect(hasWaitingPageField(editor.state.doc)).toBe(true)
    const pages = layout(editor, (block) => page(block === 1 ? 2 : 0))
    allowDirtyFields()
    const xml = await save(editor, parsed, pages)
    // where the bookmark is from a box is not said, and a field that does not wait keeps its number
    expect(results(xml)).toEqual(['PAGE=3', 'NUMPAGES=9', 'PAGEREF=1', 'PAGEREF=? dirty', 'PAGE=8'])
    expect(xml).toContain('Target')
  })

  it('a box whose block runs over two pages keeps its own page waiting', async () => {
    const { editor, parsed } = await open(BOXED)
    const before = editor.state.doc
    const pages = layout(editor, (block, edge) => page(block === 1 && edge === 'bottom' ? 3 : 2))
    const tr = editor.state.tr
    // the page count and the page of a bookmark do not need the place of the box
    expect(refreshPageFields(tr, pages, parsed.blocks)).toBe(2)
    allowDirtyFields()
    const xml = await save(editor, parsed, pages)
    expect(results(xml)).toEqual([
      'PAGE=? dirty',
      'NUMPAGES=9',
      'PAGEREF=3',
      'PAGEREF=? dirty',
      'PAGE=8',
    ])
    expect(before.child(0).eq(editor.state.doc.child(0))).toBe(true)
  })

  it('a document that was not changed keeps the fields of its text boxes as the file has them', async () => {
    const { editor, parsed } = await open(BOXED)
    const before = editor.state.doc
    const pages = layout(editor, () => page(2), { edited: false })
    settleFieldResults(editor, parsed.blocks, pages)
    expect(editor.state.doc.eq(before)).toBe(true)
  })

  it('a bookmark in a footnote of the file is on the page that draws the footnote', async () => {
    const { editor, parsed } = await open(
      p(t('See ') + field('PAGEREF Noted \\h', '?', true)) +
        p(t('Where ') + field('PAGEREF Noted \\p', '?', true)),
    )
    const pages = layout(editor, () => page(0), {
      noteBookmarkPage: (name) => (name === 'Noted' ? page(4) : null),
    })
    allowDirtyFields()
    expect(results(await save(editor, parsed, pages))).toEqual(['PAGEREF=5', 'PAGEREF=? dirty'])
  })

  it('the footnote holds its bookmark while its text is the file own', () => {
    const fromFile = [
      { id: '2', text: 'First note' },
      { id: '3', text: 'Second note', bookmarks: ['Noted', '_GoBack'] },
    ]
    expect(footnoteOfBookmark('Noted', fromFile, fromFile)).toBe('3')
    expect(footnoteOfBookmark('Other', fromFile, fromFile)).toBeNull()
    // typed into: the entry is written anew as plain text
    const typed = [fromFile[0], { ...fromFile[1], text: 'Second note!' }]
    expect(footnoteOfBookmark('Noted', typed, fromFile)).toBeNull()
    expect(footnoteOfBookmark('Noted', [fromFile[0]], fromFile)).toBeNull()
    // an endnote converted into footnote 3 is not the note the file has under that id
    const converted = [{ ...fromFile[1], origin: { kind: 'endnote' as const, id: '3' } }]
    expect(footnoteOfBookmark('Noted', converted, fromFile)).toBeNull()
  })

  it('a bookmark made here opens its paragraph: the page of its first line, though it ends on the next', async () => {
    const { editor, parsed } = await withPageFields()
    const pages = layout(editor, (block, edge) => page(block === 0 && edge === 'bottom' ? 1 : 0))
    expect(results(await save(editor, parsed, pages))[0]).toBe('PAGEREF=1')
    expect(pages.asked.filter(([, edge]) => edge === 'bottom')).toEqual([])
  })

  const MID =
    `<w:p>${t('Long opening text ')}<w:bookmarkStart w:id="7" w:name="Mid"/>${t('marked')}<w:bookmarkEnd w:id="7"/></w:p>` +
    p(t('See ') + field('PAGEREF Mid \\h', '?', true))

  /** the first paragraph runs over two pages: its text from `turn` on is on the second */
  const split = (editor: Editor, turn: number): PageSource =>
    layout(editor, (block, edge) => page(block === 0 && edge === 'bottom' ? 1 : 0), {
      pageAt: (pos, edge) =>
        page(
          pos < editor.state.doc.firstChild!.nodeSize &&
            (edge === 'bottom' || (!edge && pos >= turn))
            ? 1
            : 0,
        ),
    })

  it('a bookmark that opens inside a paragraph split over two pages takes the page it opens on', async () => {
    const { editor, parsed } = await open(MID)
    // "Long opening text " is 18 characters: the bookmark opens at 19
    expect(results(await save(editor, parsed, split(editor, 19)))).toEqual(['PAGEREF=2'])
    const again = await open(MID)
    expect(results(await save(again.editor, again.parsed, split(again.editor, 20)))).toEqual([
      'PAGEREF=1',
    ])
  })

  it('a bookmark that opens behind the last character is on the page the paragraph ends on', async () => {
    const { editor, parsed } = await open(
      `<w:p>${t('Long opening text')}<w:bookmarkStart w:id="7" w:name="Mid"/><w:bookmarkEnd w:id="7"/></w:p>` +
        p(t('See ') + field('PAGEREF Mid \\h', '?', true)),
    )
    expect(results(await save(editor, parsed, split(editor, 10)))).toEqual(['PAGEREF=2'])
  })

  it('once the paragraph no longer reads as the file has it, where the bookmark opens is not known', async () => {
    const { editor, parsed } = await open(MID)
    editor.commands.insertContentAt(1, 'New ')
    allowDirtyFields()
    expect(results(await save(editor, parsed, split(editor, 19)))).toEqual(['PAGEREF=? dirty'])
  })

  it('the same bookmark in a paragraph on one page is on that page', async () => {
    const { editor, parsed } = await open(MID)
    const pages = layout(editor, () => page(4))
    expect(results(await save(editor, parsed, pages))).toEqual(['PAGEREF=5'])
  })

  it('finds the bookmark of a caption the file came with, in the caption own XML', async () => {
    const caption = generateCaptionXml('Figure', 1, 'One')
      .replace('</w:pPr>', '</w:pPr><w:bookmarkStart w:id="60" w:name="_Ref222"/>')
      .replace(/<\/w:p>$/, '<w:bookmarkEnd w:id="60"/></w:p>')
    const { editor, parsed } = await open(
      p('Intro') + caption + p(t('On page ') + field('PAGEREF _Ref222 \\h', '9')),
    )
    expect(editor.state.doc.child(1).type.name).toBe('docProtected')
    const pages = layout(editor, (block) => page(block === 1 ? 3 : 0))
    expect(results(await save(editor, parsed, pages))).toEqual(['PAGEREF=4'])
  })

  it('a bookmark the document does not have (a footnote, a text box) leaves the field as it is', async () => {
    const { editor, parsed } = await open(p(t('See ') + field('PAGEREF Nowhere \\h', '6')))
    const pages = layout(editor, () => page(0))
    const tr = editor.state.tr
    expect(refreshPageFields(tr, pages, parsed.blocks, true)).toBe(0)
    expect(tr.docChanged).toBe(false)
  })

  it('\\p: above or below on the same sheet, "on page N" on another, sheets told apart by place not number', async () => {
    const body =
      p(t('Before ') + field('PAGEREF Target \\p \\h', '?', true)) +
      `<w:p><w:bookmarkStart w:id="3" w:name="Target"/>${t('Target')}<w:bookmarkEnd w:id="3"/></w:p>` +
      p(t('After ') + field('PAGEREF Target \\p', '?', true))
    const { editor, parsed } = await open(body)
    const same = layout(editor, () => page(0))
    expect(results(await save(editor, parsed, same))).toEqual(['PAGEREF=below', 'PAGEREF=above'])

    const again = await open(body)
    // a new section restarts at 1: the target's sheet shows "1" as the field's does
    const restart = layout(again.editor, (block) =>
      block === 1 ? page(5, { number: 1, display: '1' }) : page(0),
    )
    expect(results(await save(again.editor, again.parsed, restart))).toEqual([
      'PAGEREF=on page 1',
      'PAGEREF=on page 1',
    ])
  })
})

describe('PAGEREF \\p and the rest of the field', () => {
  const LATE = `<w:p>${field('PAGEREF Late \\p', '?', true)}${t(' some text ')}<w:bookmarkStart w:id="4" w:name="Late"/>${t('marked')}<w:bookmarkEnd w:id="4"/></w:p>`

  it('a bookmark opening later in the paragraph of the field itself is below it', async () => {
    const { editor, parsed } = await open(LATE)
    const pages = layout(editor, () => page(0))
    expect(results(await save(editor, parsed, pages))).toEqual(['PAGEREF=below'])
  })

  it('in a paragraph edited since, where the bookmark opens is not known: the field stays for Word', async () => {
    const { editor, parsed } = await open(LATE)
    editor.commands.insertContentAt(editor.state.doc.content.size - 1, ' more')
    allowDirtyFields()
    const pages = layout(editor, () => page(0))
    expect(results(await save(editor, parsed, pages))).toEqual(['PAGEREF=? dirty'])
  })

  it('a capitals switch shapes the words; a number format on top of them stays for Word', async () => {
    const body =
      `<w:p><w:bookmarkStart w:id="3" w:name="Target"/>${t('Target')}<w:bookmarkEnd w:id="3"/></w:p>` +
      p(t('After ') + field('PAGEREF Target \\p \\* Upper', '?', true)) +
      p(t('Far ') + field('PAGEREF Target \\p \\* FirstCap', '?', true)) +
      p(t('Shaped ') + field('PAGEREF Target \\p \\* ROMAN', '?', true))
    const { editor, parsed } = await open(body)
    allowDirtyFields()
    const pages = layout(editor, (block) => page(block === 2 ? 1 : 0))
    expect(results(await save(editor, parsed, pages))).toEqual([
      'PAGEREF=ABOVE',
      'PAGEREF=On page 1',
      'PAGEREF=? dirty',
    ])
  })
})

describe('a field the file brought with its result', () => {
  const FILE =
    p('Intro') + p(t('Page ') + field('PAGE', '12') + t(' of ') + field('NUMPAGES', '40'))

  it('is written back byte for byte while the document is as it was opened', async () => {
    const { editor, parsed } = await open(FILE)
    const pages = layout(editor, () => page(0), { edited: false })
    const xml = await save(editor, parsed, pages)
    expect(xml).toContain(parsed.blocks[1].originalXml)
    expect(results(xml)).toEqual(['PAGE=12', 'NUMPAGES=40'])
  })

  it('is written back byte for byte when the layout agrees with it', async () => {
    const { editor, parsed } = await open(FILE)
    const pages = layout(editor, () => page(11), { pageCount: 40 })
    const before = editor.state.doc
    const xml = await save(editor, parsed, pages)
    expect(editor.state.doc.eq(before)).toBe(true)
    expect(xml).toContain(parsed.blocks[1].originalXml)
  })

  it('takes the page of the layout once the document was changed and the result is stale', async () => {
    const { editor, parsed } = await open(FILE)
    const pages = layout(editor, () => page(13), { pageCount: 41 })
    const xml = await save(editor, parsed, pages)
    expect(results(xml)).toEqual(['PAGE=14', 'NUMPAGES=41'])
  })

  it('that came marked keeps the mark while the document is as it was opened, and takes its result once it changed', async () => {
    const { editor, parsed } = await open(p(t('Page ') + field('PAGE', '1', true)))
    allowDirtyFields()
    const untouched = await save(
      editor,
      parsed,
      layout(editor, () => page(2), { edited: false }),
    )
    expect(untouched).toContain(parsed.blocks[0].originalXml)
    expect(
      results(
        await save(
          editor,
          parsed,
          layout(editor, () => page(2)),
        ),
      ),
    ).toEqual(['PAGE=3'])
  })

  it('Update Field gives every one the page it stands on, and marks none', async () => {
    const { editor, parsed } = await open(FILE)
    const pages = layout(editor, () => page(1), { edited: false, pageCount: 2 })
    const tr = editor.state.tr
    expect(refreshPageFields(tr, pages, parsed.blocks, true)).toBe(2)
    expect(tr.doc.lastChild!.textContent).toBe('Page 2 of 2')
    expect(collectFields(tr.doc).map((f) => f.mark.attrs.dirty)).toEqual([false, false])
  })
})

describe('Update Field rewrites the other fields after the page fields', () => {
  it('a field right beside a rewritten page field keeps its own run', async () => {
    const { editor, parsed } = await open(
      p(field('PAGE', '1') + field('FILENAME', 'old.docx') + field('PAGE', '1')),
    )
    const before = collectFields(editor.state.doc)
    expect(before.map((f) => f.instr)).toEqual(['PAGE', 'FILENAME', 'PAGE'])
    const tr = editor.state.tr
    refreshPageFields(
      tr,
      layout(editor, () => page(122)),
      parsed.blocks,
      true,
    )
    const name = editor.state.doc.nodeAt(before[1].from)!
    replaceFieldRuns(tr, [
      { from: before[1].from, to: before[1].to, text: 'new.docx', marks: name.marks },
    ])
    expect(
      collectFields(tr.doc).map((f) => `${f.instr}=${tr.doc.textBetween(f.from, f.to)}`),
    ).toEqual(['PAGE=123', 'FILENAME=new.docx', 'PAGE=123'])
  })
})

describe('a header or footer written anew stores the first page it shows on', () => {
  const hf = { text: 'x' }
  it('the last section through the document slots, the others by their closing block', () => {
    const asked: string[] = []
    const out = withHfFieldResults(
      {
        footer: hf,
        headerFirst: hf,
        sectionHf: [{ lastBlockIndex: 4, kind: 'footer', variant: 'even', hf }],
      },
      9,
      (sectionEnd, variant) => {
        asked.push(`${sectionEnd}:${variant}`)
        return variant === 'first' ? null : page(sectionEnd, { display: 'v', sectionPages: 3 })
      },
    )
    expect(asked).toEqual(['9:first', '9:default', '4:even'])
    expect(out.footer?.fieldResults).toEqual({ page: 'v', sectionPages: 3 })
    // a part no page shows keeps the engine's default
    expect(out.headerFirst).toBe(hf)
    expect(out.sectionHf?.[0].hf.fieldResults).toEqual({ page: 'v', sectionPages: 3 })
  })

  it('without a layout the parts go out as they are', () => {
    const options = { footer: hf }
    expect(withHfFieldResults(options, 9, undefined)).toBe(options)
  })
})

describe('the page at a place on the canvas', () => {
  const gap = (more: Record<string, unknown> = {}) => ({
    kind: 'block',
    heightPx: 100,
    marginBottomPx: 40,
    marginTopPx: 40,
    ...more,
  })
  /** three sheets of 500 px of text under a 50 px top margin, 100 px between the text of two sheets */
  const cuts = (gaps: Array<Record<string, unknown> | null>, blank: number[] = []) =>
    gaps.map((g, i) => ({
      start: i * 500,
      end: blank.includes(i) ? i * 500 : (i + 1) * 500,
      blank: blank.includes(i),
      gapOffsetPx: i * 100,
      ...(g ? { gap: g } : {}),
    })) as unknown as PageCut[]

  it('tells the sheet of a line, and tells nothing between two sheets', () => {
    const three = cuts([null, gap(), gap()])
    expect(pageAtCanvasY(three, 50, 60)).toBe(0)
    expect(pageAtCanvasY(three, 50, 549)).toBe(0)
    // the bottom margin of sheet 1 and the top margin of sheet 2 are theirs; the band between is neither
    expect(pageAtCanvasY(three, 50, 580)).toBe(0)
    expect(pageAtCanvasY(three, 50, 600)).toBeNull()
    expect(pageAtCanvasY(three, 50, 620)).toBe(1)
    expect(pageAtCanvasY(three, 50, 700)).toBe(1)
    expect(pageAtCanvasY(three, 50, 1300)).toBe(2)
  })

  it('a page turn without a band of its own leaves the sheets on both sides undecided', () => {
    for (const unbanded of [null, gap({ kind: 'cut', heightPx: 0 }), gap({ floatGap: true })]) {
      const three = cuts([null, unbanded, gap()])
      expect(pageAtCanvasY(three, 50, 60)).toBeNull()
      expect(pageAtCanvasY(three, 50, 700)).toBeNull()
      expect(pageAtCanvasY(three, 50, 1300)).toBe(2)
    }
  })

  it('a blank parity sheet draws nothing: the text beside it is on the sheet that does', () => {
    const three = cuts([null, null, gap()], [1])
    expect(pageAtCanvasY(three, 50, 60)).toBe(0)
    expect(pageAtCanvasY(cuts([]), 50, 60)).toBeNull()
  })
})
