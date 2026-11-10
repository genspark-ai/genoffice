import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { executeOps } from '../src/renderer/ai/ops'
import {
  blocksToPmDoc,
  inlineToRuns,
  pmDocToSavePlan,
  type PmNode,
} from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

// Word writes a Symbol-font glyph as its own <w:sym> run; regenerating the
// paragraph from the editor used to flatten both into w:t text (the
// private-use code point literally, the second one remapped to U+00D7).
const MINUS = '<w:r><w:sym w:font="Symbol" w:char="F02D"/></w:r>'
const TIMES = '<w:r><w:rPr><w:b/></w:rPr><w:sym w:font="Symbol" w:char="F0B4"/></w:r>'
const SYM_PARA =
  '<w:p><w:r><w:t xml:space="preserve">fee (91 </w:t></w:r>' +
  MINUS +
  '<w:r><w:t xml:space="preserve"> 50 </w:t></w:r>' +
  TIMES +
  '<w:r><w:t xml:space="preserve"> 0.2).</w:t></w:r></w:p>'
const OTHER_PARA = '<w:p><w:r><w:t>other</w:t></w:r></w:p>'

async function openDoc(bodyXml: string) {
  const source = await buildDocx({ bodyXml })
  const parsed = await parseDocx(source)
  const host = document.createElement('div')
  document.body.appendChild(host)
  const editor = new Editor({
    element: host,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { source, parsed, editor, host }
}

async function documentXml(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

describe('w:sym runs survive saving', () => {
  it('parses each symbol as its own run with the display glyph', async () => {
    const { parsed, editor, host } = await openDoc(SYM_PARA)
    const runs = parsed.blocks[0].runs!
    expect(runs.map((r) => r.sym ?? null)).toEqual([
      null,
      { font: 'Symbol', char: 'F02D' },
      null,
      { font: 'Symbol', char: 'F0B4' },
      null,
    ])
    expect(runs[3]).toMatchObject({ text: '×', bold: true })
    expect(editor.getText()).toBe('fee (91 − 50 × 0.2).')
    editor.destroy()
    host.remove()
  })

  it('keeps an untouched paragraph byte-exact', async () => {
    const { source, parsed, editor, host } = await openDoc(SYM_PARA + OTHER_PARA)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
    host.remove()
  })

  it('regenerates both w:sym elements after a text edit in the same paragraph', async () => {
    const { parsed, editor, host } = await openDoc(SYM_PARA + OTHER_PARA)
    const outcome = executeOps(editor, [
      { op: 'findReplace', find: '0.2).', replace: '0.2).[EE12]' },
    ])
    expect(outcome.ok).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const xml = await documentXml(await saveDocx(parsed, plan.saveBlocks))
    expect(xml).toContain(MINUS)
    expect(xml).toContain(TIMES)
    expect(xml.match(/<w:sym\b/g)).toHaveLength(2)
    expect(xml).toMatch(/50 <\/w:t><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:sym/)
    expect(xml).toContain('0.2).[EE12]</w:t>')
    expect(xml).not.toMatch(/[-×]/)
    editor.destroy()
    host.remove()
  })

  it('splits text appended under the symbol mark into one w:sym plus plain w:t', async () => {
    const { parsed, editor, host } = await openDoc(SYM_PARA)
    const outcome = executeOps(editor, [
      { op: 'findReplace', find: '\u00d7', replace: '\u00d7\u00d7yz' },
    ])
    expect(outcome.ok).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xml = await documentXml(await saveDocx(parsed, plan.saveBlocks))
    expect(xml.match(/<w:sym\b/g)).toHaveLength(3)
    expect(xml).toMatch(
      /<w:sym w:font="Symbol" w:char="F0B4"\/><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t[^>]*>yz</,
    )
    expect(xml).not.toMatch(/[\uf000-\uf0ff\u00d7]/)
    editor.destroy()
    host.remove()
  })

  it('writes a symbol replaced by ordinary text as plain w:t', async () => {
    const { parsed, editor, host } = await openDoc(SYM_PARA)
    const outcome = executeOps(editor, [{ op: 'findReplace', find: '×', replace: 'x' }])
    expect(outcome.ok).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xml = await documentXml(await saveDocx(parsed, plan.saveBlocks))
    expect(xml).toContain(MINUS)
    expect(xml.match(/<w:sym\b/g)).toHaveLength(1)
    expect(xml).toContain('>x</w:t>')
    editor.destroy()
    host.remove()
  })
})

describe('w:sym drawn by its installed font', () => {
  it('the private-use editor text goes back to the run as the decoded glyph', () => {
    const marks = [{ type: 'docSym', attrs: { font: 'Wingdings', char: 'F0FC' } }]
    const runs = inlineToRuns([
      { type: 'text', text: 'a' },
      { type: 'text', text: '\uF0FC', marks },
      { type: 'text', text: 'b' },
    ])
    expect(runs.map((r) => r.text)).toEqual(['a', '✓', 'b'])
    expect(runs[1].sym).toEqual({ font: 'Wingdings', char: 'F0FC' })
  })

  it('an unmapped glyph stays private-use under its font', () => {
    const marks = [{ type: 'docSym', attrs: { font: 'Marlett', char: 'F061' } }]
    const runs = inlineToRuns([{ type: 'text', text: '\uF061', marks }])
    expect(runs[0]).toMatchObject({ text: '\uF061', sym: { font: 'Marlett', char: 'F061' } })
  })
})
