import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { DocRedaction } from '../src/renderer/editor/redaction'
import { REDACT_MARK } from '../src/renderer/ai/redact'

/**
 * The whole feature rests on the mark surviving a .docx round trip: a
 * character border is real OOXML, so Word renders it, and the label rides in
 * the run's rPr, which the engine writes back verbatim. Without either half the
 * mark is lost on the first save and the document quietly becomes readable by
 * a model again.
 */

const SECRET = '13800138000'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const extensions = [...editorExtensions, DocRedaction]
const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

function makeDocx(inner: string): Promise<Uint8Array> {
  return buildDocx({ bodyXml: `<w:p>${inner}</w:p>` })
}

function open(bytes: Uint8Array): Promise<Editor> {
  return parseDocx(bytes).then((parsed) => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions,
      content: { type: 'doc', content: blocksToPmDoc(parsed.blocks).content as JsonNode[] },
    })
    editors.add(editor)
    return editor
  })
}

async function save(editor: Editor, from: Uint8Array): Promise<Uint8Array> {
  const parsed = await parseDocx(from)
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  return saveDocx(parsed, plan.saveBlocks)
}

function selectText(editor: Editor, needle: string) {
  let start = -1
  editor.state.doc.descendants((node, pos) => {
    if (start >= 0) return
    const at = node.textContent.indexOf(needle)
    if (at >= 0) start = pos + 1 + at
  })
  expect(start).toBeGreaterThan(0)
  editor.view.dispatch(
    editor.state.tr.setSelection(
      TextSelection.create(editor.state.doc, start, start + needle.length),
    ),
  )
}

function redactionMarks(editor: Editor): Array<Record<string, unknown>> {
  const out: Array<Record<string, unknown>> = []
  editor.state.doc.descendants((node) => {
    const mark = node.marks.find((m) => m.type.name === REDACT_MARK)
    if (mark) out.push(mark.attrs as Record<string, unknown>)
  })
  return out
}

describe('a withheld span survives the .docx', () => {
  it('recovers from a file that already carries the mark', async () => {
    const docx = await makeDocx(
      '<w:r><w:t xml:space="preserve">Call </w:t></w:r>' +
        '<w:r><w:rPr><w:bdr w:val="single" w:sz="6" w:color="6B4FC0"/>' +
        '<go:redact w:label="客户电话"/></w:rPr><w:t>13800138000</w:t></w:r>' +
        '<w:r><w:t xml:space="preserve"> now</w:t></w:r>',
    )
    const editor = await open(docx)
    // the border arrives as a docTextStyle mark; the custom element in the
    // run's rPr is what makes it ours, and it carries the label back
    expect(redactionMarks(editor).map((a) => a.label)).toEqual(['客户电话'])
    expect(editor.state.doc.textContent).toContain(SECRET)
  })

  it('does not mistake an ordinary border for a withheld span', async () => {
    // a reader can underline or box text themselves; only our element counts
    const docx = await makeDocx(
      '<w:r><w:rPr><w:bdr w:val="single" w:sz="6" w:color="000000"/></w:rPr>' +
        '<w:t>important but not secret</w:t></w:r>',
    )
    const editor = await open(docx)
    expect(redactionMarks(editor)).toEqual([])
    expect(editor.state.doc.textContent).toContain('important but not secret')
  })

  it('round-trips the label, not just the border', async () => {
    const docx = await makeDocx('<w:r><w:t xml:space="preserve">Call 13800138000 now</w:t></w:r>')
    const editor = await open(docx)
    selectText(editor, SECRET)
    editor.commands.setRedaction('收货地址')

    const out = await save(editor, docx)
    const reopened = await open(out)
    expect(redactionMarks(reopened).map((a) => a.label)).toEqual(['收货地址'])
    expect(reopened.state.doc.textContent).toContain(SECRET)
  })

  it('writes the words and the border, never the placeholder', async () => {
    const docx = await makeDocx('<w:r><w:t xml:space="preserve">Call 13800138000 now</w:t></w:r>')
    const editor = await open(docx)
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')

    const out = await save(editor, docx)
    const reparsed = await parseDocx(out)
    const runs = reparsed.blocks[0]?.runs ?? []
    const flat = JSON.stringify(runs)
    expect(flat).toContain(SECRET)
    expect(flat).not.toContain('{{')
  })

  it('keeps the words, so the reader still has their data', async () => {
    const docx = await makeDocx('<w:r><w:t xml:space="preserve">Call 13800138000 now</w:t></w:r>')
    const editor = await open(docx)
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    expect(editor.state.doc.textContent).toContain(SECRET)

    const out = await save(editor, docx)
    const reopened = await open(out)
    expect(reopened.state.doc.textContent).toContain(SECRET)
  })

  it('writes the border from the mark’s own attr, not from a leftover one', async () => {
    // the bdr on the run is what Word draws; if runFromMarks stopped reading it
    // the file would still carry the label but lose the visible mark
    const docx = await makeDocx('<w:r><w:t xml:space="preserve">Call 13800138000 now</w:t></w:r>')
    const editor = await open(docx)
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    const out = await save(editor, docx)
    const reparsed = await parseDocx(out)
    const run = (reparsed.blocks[0]?.runs ?? []).find((r) => r.text === SECRET)
    expect(run?.bdr).toMatchObject({ val: 'single', sz: 6, color: '6B4FC0' })
    // and the mark survives a second round trip with the border still there
    const again = await open(out)
    expect(redactionMarks(again).map((a) => a.label)).toEqual(['客户电话'])
    expect(again.state.doc.textContent).toContain(SECRET)
  })

  it('writes a real character border, which is what Word will render', async () => {
    const docx = await makeDocx('<w:r><w:t xml:space="preserve">Call 13800138000 now</w:t></w:r>')
    const editor = await open(docx)
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    const out = await save(editor, docx)
    const reparsed = await parseDocx(out)
    const run = (reparsed.blocks[0]?.runs ?? []).find((r) => r.text === SECRET)
    expect(run?.bdr).toMatchObject({ val: 'single', sz: 6 })
  })

  it('refuses an empty label', async () => {
    const docx = await makeDocx('<w:r><w:t xml:space="preserve">Call 13800138000 now</w:t></w:r>')
    const editor = await open(docx)
    selectText(editor, SECRET)
    expect(editor.commands.setRedaction('   ')).toBe(false)
  })

  it('stops withholding and leaves the text alone', async () => {
    const docx = await makeDocx('<w:r><w:t xml:space="preserve">Call 13800138000 now</w:t></w:r>')
    const editor = await open(docx)
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    editor.commands.unsetRedaction()
    expect(redactionMarks(editor)).toEqual([])
    expect(editor.state.doc.textContent).toContain(SECRET)
  })
})

describe('the label never breaks the file', () => {
  it('escapes the characters that would end the attribute', async () => {
    const docx = await makeDocx('<w:r><w:t xml:space="preserve">Call 13800138000 now</w:t></w:r>')
    const editor = await open(docx)
    selectText(editor, SECRET)
    // the label is sanitized before it ever reaches the file
    editor.commands.setRedaction('a"b<c&d')
    const out = await save(editor, docx)
    const reparsed = await parseDocx(out)
    expect(JSON.stringify(reparsed.blocks[0]?.runs)).toContain(SECRET)
  })

  it('keeps a long document saveable', async () => {
    const long = 'lorem ipsum dolor sit amet '.repeat(400)
    const docx = await makeDocx(`<w:r><w:t xml:space="preserve">${long}</w:t></w:r>`)
    const editor = await open(docx)
    selectText(editor, 'dolor')
    editor.commands.setRedaction('关键词')
    const out = await save(editor, docx)
    const reparsed = await parseDocx(out)
    expect(JSON.stringify(reparsed.blocks[0]?.runs)).toContain('dolor')
  })
})
