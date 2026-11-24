/** Linked content: link fields as two markers, links kept in a block's XML, linked pictures. */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { docxRootNamespaces, parseDocx, saveDocx } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocOptions, pmDocToSavePlan } from '../src/renderer/editor/convert'

class FakeDataTransfer {
  private data = new Map<string, string>()
  setData(type: string, value: string) {
    this.data.set(type, value)
  }
  getData(type: string) {
    return this.data.get(type) ?? ''
  }
  get types() {
    return [...this.data.keys()]
  }
}
class FakeClipboardEvent extends Event {
  clipboardData: FakeDataTransfer | null
  constructor(type: string, init?: { clipboardData?: FakeDataTransfer }) {
    super(type)
    this.clipboardData = init?.clipboardData ?? null
  }
}
const g = globalThis as Record<string, unknown>
g.DataTransfer = FakeDataTransfer
g.ClipboardEvent = FakeClipboardEvent

const {
  breakLinks,
  collectLinks,
  hasLinks,
  insertLinkedFileBody,
  setLinkAutomatic,
  setLinkLocked,
  setLinksAutomatic,
  setLinksLocked,
  setLinkSource,
} = await import('../src/renderer/editor/link-fields')
const { insertFileBody } = await import('../src/renderer/editor/insert-file')

const text = (s: string) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`
const para = (inner: string) => `<w:p>${inner}</w:p>`
const begin = (instr: string, attrs = '') =>
  `<w:r><w:fldChar w:fldCharType="begin"${attrs}/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
const END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

async function opened(bodyXml: string, withImage = false) {
  const bytes = await buildDocx({ bodyXml, withImage })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks, undefined, pmDocOptions(parsed)),
  })
  return { parsed, editor }
}

async function savedXml(
  editor: Editor,
  parsed: Awaited<ReturnType<typeof parseDocx>>,
): Promise<string> {
  const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

const chars = (xml: string) =>
  [...xml.matchAll(/w:fldCharType="(begin|separate|end)"/g)].map((m) => m[1][0]).join('')

describe('link fields in the editor', () => {
  it('are listed with their source, and an untouched document saves as it came', async () => {
    const body =
      para(text('Lead ') + begin('INCLUDETEXT "/tmp/part.docx" Intro') + text('First')) +
      para(text('Second') + END) +
      para(text('Tail'))
    const { parsed, editor } = await opened(body)
    expect(hasLinks(editor.state.doc, parsed.blocks)).toBe(true)
    const links = collectLinks(editor.state.doc, parsed.blocks)
    expect(links).toHaveLength(1)
    expect(links[0]).toMatchObject({
      type: 'field',
      locked: false,
      info: { kind: 'INCLUDETEXT', source: '/tmp/part.docx', item: 'Intro' },
    })
    const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
    expect(plan.saveBlocks.slice(0, 3).every((b) => b.kind === 'original')).toBe(true)
    editor.destroy()
  })

  it('lock, take another source and break, each written to the file', async () => {
    const body =
      para(begin('INCLUDETEXT "/tmp/part.docx"') + text('First')) + para(text('Second') + END)
    const { parsed, editor } = await opened(body)
    const link = () => collectLinks(editor.state.doc, parsed.blocks)[0]
    expect(setLinkLocked(editor, link(), parsed.blocks, true)).toBe(true)
    expect(setLinkSource(editor, link(), parsed.blocks, { path: '/tmp/other.docx', url: '' })).toBe(
      true,
    )
    let xml = await savedXml(editor, parsed)
    expect(xml).toContain('w:fldCharType="begin" w:fldLock="true"/>')
    expect(xml).toContain(' INCLUDETEXT "/tmp/other.docx" ')
    expect(chars(xml)).toBe('bse')
    expect(breakLinks(editor, [link()], parsed.blocks)).toBe(1)
    expect(hasLinks(editor.state.doc, parsed.blocks)).toBe(false)
    xml = await savedXml(editor, parsed)
    expect(chars(xml)).toBe('')
    expect(xml).toContain('First')
    expect(xml).toContain('Second')
    editor.destroy()
  })

  it('lose the other marker when an edit takes one away', async () => {
    const body =
      para(begin('INCLUDETEXT "/tmp/part.docx"') + text('First')) +
      para(text('Second') + END) +
      para(text('Tail'))
    const { parsed, editor } = await opened(body)
    const { begin: at } = collectLinks(editor.state.doc, parsed.blocks)[0] as { begin: number }
    // the first paragraph goes, and the opening marker with it
    const $at = editor.state.doc.resolve(at)
    editor.view.dispatch(editor.state.tr.delete($at.before(), $at.after()))
    expect(collectLinks(editor.state.doc, parsed.blocks)).toHaveLength(0)
    let marks = 0
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docXeMark') marks++
    })
    expect(marks).toBe(0)
    expect(chars(await savedXml(editor, parsed))).toBe('')
    editor.destroy()
  })

  it('keep the outer field whole when a nested one loses a marker', async () => {
    const body =
      para(begin('INCLUDETEXT "/tmp/outer.docx"') + text('Outer ')) +
      para(text('Inner')) +
      para(text('tail') + END)
    const { parsed, editor } = await opened(body)
    const mark = editor.schema.nodes.docXeMark
    let inner = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && node.text === 'Inner') inner = pos
    })
    editor.view.dispatch(
      editor.state.tr
        .insert(inner + 5, mark.create({ linkEnd: true }))
        .insert(inner, mark.create({ linkBegin: ' INCLUDETEXT "/tmp/inner.docx" ' })),
    )
    expect(collectLinks(editor.state.doc, parsed.blocks)).toHaveLength(2)
    editor.view.dispatch(editor.state.tr.delete(inner, inner + 1))
    const left = collectLinks(editor.state.doc, parsed.blocks)
    expect(left).toHaveLength(1)
    const outer = left[0] as { begin: number; end: number; info: { source: string } }
    expect(outer.info.source).toBe('/tmp/outer.docx')
    expect(editor.state.doc.textBetween(outer.begin, outer.end, '|')).toBe('Outer |Inner|tail')
    editor.destroy()
  })

  it('wrap a file inserted with Link to File, in one undo step', async () => {
    const { parsed, editor } = await opened(para(text('Before after')))
    const otherBytes = await buildDocx({
      bodyXml: para(text('Linked one')) + para(text('Linked two')),
    })
    const other = {
      parsed: await parseDocx(otherBytes),
      namespaces: await docxRootNamespaces(otherBytes),
    }
    editor.commands.setTextSelection(8)
    expect(insertLinkedFileBody(editor, other, undefined, 'INCLUDETEXT "/tmp/o.docx"')).toEqual({
      leftOut: 0,
      linked: true,
    })
    const links = collectLinks(editor.state.doc, parsed.blocks)
    expect(links).toHaveLength(1)
    const { begin: from, end: to } = links[0] as { begin: number; end: number }
    // the inserted paragraphs stand between the host's two halves
    expect(editor.state.doc.textBetween(from, to, '|')).toBe('|Linked one|Linked two|')
    const paragraphs = () => {
      const out: string[] = []
      editor.state.doc.forEach((node) => out.push(node.textContent))
      return out
    }
    expect(paragraphs()).toEqual(['Before ', 'Linked one', 'Linked two', 'after'])
    const xml = await savedXml(editor, parsed)
    expect(chars(xml)).toBe('bse')
    const order = [
      'Before ',
      'INCLUDETEXT "/tmp/o.docx"',
      'Linked one',
      'Linked two',
      '"end"',
      'after',
    ]
    const at = order.map((s) => xml.indexOf(s))
    expect([...at].sort((a, b) => a - b)).toEqual(at)
    // an update puts the source in the place of the result, between the same markers
    const again = await buildDocx({ bodyXml: para(text('Newer text')) })
    editor.commands.setTextSelection({ from: from + 1, to })
    insertFileBody(editor, {
      parsed: await parseDocx(again),
      namespaces: await docxRootNamespaces(again),
    })
    expect(paragraphs()).toEqual(['Before ', 'Newer text', 'after'])
    const now = collectLinks(editor.state.doc, parsed.blocks)[0] as { begin: number; end: number }
    expect(editor.state.doc.textBetween(now.begin, now.end, '|')).toBe('|Newer text|')
    editor.commands.undo()
    editor.commands.undo()
    expect(editor.state.doc.textContent).toBe('Before after')
    editor.destroy()
  })

  it('wrap what goes in after the paragraph the caret ends', async () => {
    const { parsed, editor } = await opened(para(text('Host first')) + para(text('Host last')))
    const otherBytes = await buildDocx({
      bodyXml: para(text('Linked one')) + para(text('Linked two')),
    })
    const other = {
      parsed: await parseDocx(otherBytes),
      namespaces: await docxRootNamespaces(otherBytes),
    }
    editor.commands.setTextSelection(11)
    expect(insertLinkedFileBody(editor, other, undefined, 'INCLUDETEXT "/tmp/o.docx"')).toEqual({
      leftOut: 0,
      linked: true,
    })
    const xml = await savedXml(editor, parsed)
    expect(chars(xml)).toBe('bse')
    const order = ['Host first', '"separate"', 'Linked one', 'Linked two', '"end"', 'Host last']
    const at = order.map((s) => xml.indexOf(s))
    expect([...at].sort((a, b) => a - b)).toEqual(at)
    editor.destroy()
  })

  it('several links are locked in one step, and one undo unlocks them all', async () => {
    const body =
      para(text('A: ') + begin('LINK Excel.Sheet.12 "C:\\\\a.xlsx" "S!R1C1"') + text('1') + END) +
      para(text('B: ') + begin('LINK Excel.Sheet.12 "C:\\\\b.xlsx" "S!R1C1"') + text('2') + END)
    const { parsed, editor } = await opened(body, true)
    const links = () => collectLinks(editor.state.doc, parsed.blocks)
    const locks = () => links().map((l) => l.type !== 'picture' && l.locked)
    expect(setLinksLocked(editor, links(), parsed.blocks, true)).toBe(2)
    expect(locks()).toEqual([true, true])
    expect(setLinksAutomatic(editor, links(), parsed.blocks, true)).toBe(2)
    editor.commands.undo()
    editor.commands.undo()
    expect(locks()).toEqual([false, false])
    editor.destroy()
  })

  it('kept in a block as XML are listed, locked, set to automatic and broken in that XML', async () => {
    const picture = IMAGE_PARAGRAPH_XML.replace(/^<w:p>/, '').replace(/<\/w:p>$/, '')
    const body =
      para(begin('INCLUDEPICTURE "https://example.com/a.png" \\d') + picture + END) +
      para(text('Sum: ') + begin('LINK Excel.Sheet.12 "C:\\\\b.xlsx" "S!R1C1"') + text('7') + END)
    const { parsed, editor } = await opened(body, true)
    const links = () => collectLinks(editor.state.doc, parsed.blocks)
    expect(links().map((l) => [l.type, l.type === 'picture' ? '' : l.info.kind])).toEqual([
      ['blockField', 'INCLUDEPICTURE'],
      ['field', 'LINK'],
    ])
    expect(setLinkLocked(editor, links()[0], parsed.blocks, true)).toBe(true)
    expect(setLinkAutomatic(editor, links()[1], parsed.blocks, true)).toBe(true)
    expect(setLinkAutomatic(editor, links()[0], parsed.blocks, true)).toBe(false)
    let xml = await savedXml(editor, parsed)
    expect(xml).toMatch(/w:fldLock="true"\/><\/w:r><w:r><w:instrText[^>]*> INCLUDEPICTURE/)
    expect(xml).toContain('"S!R1C1" \\a ')
    expect(chars(xml)).toBe('bsebse')
    expect(breakLinks(editor, links(), parsed.blocks)).toBe(2)
    xml = await savedXml(editor, parsed)
    expect(chars(xml)).toBe('')
    expect(xml).toContain('<w:drawing>')
    expect(xml).toContain('Sum: 7')
    editor.destroy()
  })
})
