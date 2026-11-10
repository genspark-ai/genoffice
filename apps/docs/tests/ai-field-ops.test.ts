import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { executeOps } from '../src/renderer/ai/ops'
import { evaluateIf, formatFieldDate, formatSeqNumber } from '../src/renderer/ai/field-ops'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const para = (t: string, attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, ...attrs },
  content: [text(t)],
})

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
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

function fieldsOf(
  editor: Editor,
): Array<{ text: string; instr?: string; dirty?: boolean; ref?: string }> {
  const out: Array<{ text: string; instr?: string; dirty?: boolean; ref?: string }> = []
  editor.state.doc.descendants((node) => {
    if (!node.isText) return true
    const instr = node.marks.find((m) => m.type.name === 'instrField')
    const ref = node.marks.find((m) => m.type.name === 'refField')
    if (instr)
      out.push({ text: node.text!, instr: String(instr.attrs.instr), dirty: instr.attrs.dirty })
    else if (ref) out.push({ text: node.text!, ref: String(ref.attrs.name) })
    return true
  })
  return out
}

const at = (index: number) => ({ blockIndexes: [index] })

describe('insertField', () => {
  it('numbers SEQ fields in document order and marks them for Word', () => {
    const editor = createEditor([para('Figure '), para('Figure '), para('Table ')])
    const outcome = executeOps(editor, [
      { op: 'insertField', target: at(0), type: 'SEQ', args: 'Figure \\* ARABIC' },
      { op: 'insertField', target: at(1), type: 'SEQ', args: 'Figure \\* ARABIC' },
      { op: 'insertField', target: at(2), type: 'seq', args: 'Table \\* ROMAN' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(fieldsOf(editor)).toEqual([
      { text: '1', instr: 'SEQ Figure \\* ARABIC', dirty: true },
      { text: '2', instr: 'SEQ Figure \\* ARABIC', dirty: true },
      { text: 'I', instr: 'SEQ Table \\* ROMAN', dirty: true },
    ])
    expect(editor.state.doc.child(0).textContent).toBe('Figure 1')
    expect(outcome.results[0]!.detail).toContain('{ SEQ Figure \\* ARABIC } = "1"')
  })

  it('places the field after the given text and inherits the run formatting', () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: null },
        content: [
          { type: 'text', text: 'See page ', marks: [{ type: 'bold' }] },
          text(' for details.'),
        ],
      },
    ])
    const outcome = executeOps(editor, [
      { op: 'insertField', target: at(0), type: 'PAGE', afterText: 'See page' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(editor.state.doc.child(0).textContent).toBe('See page1  for details.')
    const field = editor.state.doc.child(0).child(1)
    expect(field.marks.map((m) => m.type.name).sort()).toEqual(['bold', 'instrField'])
    expect(field.marks.find((m) => m.type.name === 'instrField')!.attrs).toMatchObject({
      instr: 'PAGE',
      dirty: true,
    })
  })

  it('computes DATE with a picture, MERGEFIELD and literal IF locally (not dirty)', () => {
    const editor = createEditor([para('Dated: '), para('Dear '), para('Status: ')])
    const outcome = executeOps(editor, [
      { op: 'insertField', target: at(0), type: 'DATE', args: '\\@ "yyyy"' },
      { op: 'insertField', target: at(1), type: 'MERGEFIELD', args: 'FirstName' },
      { op: 'insertField', target: at(2), type: 'IF', args: '3 > 2 "open" "closed"' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    const fields = fieldsOf(editor)
    expect(fields[0]).toEqual({
      text: String(new Date().getFullYear()),
      instr: 'DATE \\@ "yyyy"',
      dirty: false,
    })
    expect(fields[1]).toEqual({ text: '«FirstName»', instr: 'MERGEFIELD FirstName', dirty: false })
    expect(fields[2]).toEqual({ text: 'open', instr: 'IF 3 > 2 "open" "closed"', dirty: false })
    // an empty computed result is stored as Word's single space, not an invalid empty text node
    const empty = executeOps(editor, [
      { op: 'insertField', target: at(2), type: 'IF', args: '1 = 1 "" "no"' },
      { op: 'insertField', target: at(2), type: 'SEQ', args: 'Blank \\r 0 \\* ROMAN' },
    ])
    expect(empty.ok, empty.error).toBe(true)
    expect(
      fieldsOf(editor)
        .slice(3)
        .map((f) => f.text),
    ).toEqual([' ', ' '])
  })

  it('REF takes the bookmarked paragraph text; a missing bookmark lists the known ones', () => {
    const editor = createEditor([para('Conclusion paragraph.'), para('See '), para('Also ')])
    const outcome = executeOps(editor, [
      { op: 'insertBookmark', target: at(0), name: 'Conclusion' },
      { op: 'insertField', target: at(1), type: 'REF', args: 'Conclusion \\h' },
      { op: 'insertField', target: at(2), type: 'PAGEREF', args: 'Conclusion \\h' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(editor.state.doc.child(0).attrs.bookmarks).toEqual(['Conclusion'])
    expect(fieldsOf(editor)).toEqual([
      { text: 'Conclusion paragraph.', ref: 'Conclusion' },
      { text: '1', instr: 'PAGEREF Conclusion \\h', dirty: true },
    ])
    const missing = executeOps(editor, [
      { op: 'insertField', target: at(1), type: 'REF', args: 'Nowhere' },
    ])
    expect(missing.ok).toBe(false)
    expect(missing.error).toContain('no bookmark named "Nowhere"')
    expect(missing.error).toContain('Conclusion')
    const dup = executeOps(editor, [{ op: 'insertBookmark', target: at(1), name: 'Conclusion' }])
    expect(dup.error).toContain('already exists')
  })

  it('refuses fields it cannot compute without a result, and blocked keywords', () => {
    const editor = createEditor([para('By ')])
    const noResult = executeOps(editor, [{ op: 'insertField', target: at(0), type: 'AUTHOR' }])
    expect(noResult.ok).toBe(false)
    expect(noResult.error).toContain('pass result')
    const withResult = executeOps(editor, [
      { op: 'insertField', target: at(0), type: 'AUTHOR', result: 'Ada' },
    ])
    expect(withResult.ok, withResult.error).toBe(true)
    expect(fieldsOf(editor)).toEqual([{ text: 'Ada', instr: 'AUTHOR', dirty: true }])
    const toc = executeOps(editor, [{ op: 'insertField', target: at(0), type: 'TOC' }])
    expect(toc.error).toContain('use insertToc')
    const both = executeOps(editor, [
      { op: 'insertField', target: at(0), type: 'PAGE', position: 'start', afterText: 'By' },
    ])
    expect(both.error).toContain('not both')
  })

  it('afterText never lands inside an existing field result', () => {
    const seq = { type: 'instrField', attrs: { instr: 'SEQ Figure \\* ARABIC', dirty: true } }
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: null },
        content: [
          text('Figure '),
          { type: 'text', text: '1', marks: [seq] },
          { type: 'text', text: '2', marks: [{ type: 'bold' }, seq] },
          text(' shows the flow; see Fig. 1 above.'),
        ],
      },
    ])
    const later = executeOps(editor, [
      { op: 'insertField', target: at(0), type: 'PAGE', afterText: '1' },
    ])
    expect(later.ok, later.error).toBe(true)
    expect(editor.state.doc.child(0).textContent).toBe(
      'Figure 12 shows the flow; see Fig. 11 above.',
    )
    expect(fieldsOf(editor)).toEqual([
      { text: '1', instr: 'SEQ Figure \\* ARABIC', dirty: true },
      { text: '2', instr: 'SEQ Figure \\* ARABIC', dirty: true },
      { text: '1', instr: 'PAGE', dirty: true },
    ])
    const onlyInside = executeOps(editor, [
      { op: 'insertField', target: at(0), type: 'PAGE', afterText: 'Figure 1' },
    ])
    expect(onlyInside.ok).toBe(false)
    expect(onlyInside.error).toContain('ends inside the field { SEQ Figure \\* ARABIC }')
    const whole = executeOps(editor, [
      { op: 'insertField', target: at(0), type: 'PAGE', afterText: 'Figure 12' },
    ])
    expect(whole.ok, whole.error).toBe(true)
    expect(editor.state.doc.child(0).textContent).toBe(
      'Figure 121 shows the flow; see Fig. 11 above.',
    )
    expect(fieldsOf(editor).slice(0, 3)).toEqual([
      { text: '1', instr: 'SEQ Figure \\* ARABIC', dirty: true },
      { text: '2', instr: 'SEQ Figure \\* ARABIC', dirty: true },
      { text: '1', instr: 'PAGE', dirty: true },
    ])
  })

  it('needs exactly one paragraph-like target', () => {
    const editor = createEditor([para('a'), para('b')])
    const many = executeOps(editor, [
      { op: 'insertField', target: { nodeType: 'docParagraph' }, type: 'PAGE' },
    ])
    expect(many.error).toContain('exactly one paragraph (matched 2)')
  })
})

describe('updateFields', () => {
  it('keeps partial formatting on a split result when only the dirty flag changes', () => {
    const editor = createEditor([para('Status: ')])
    let outcome = executeOps(editor, [
      { op: 'insertField', target: at(0), type: 'IF', args: '3 > 2 "open" "closed"' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    outcome = executeOps(editor, [{ op: 'setMatchedFont', text: 'op', target: at(0), bold: true }])
    expect(outcome.ok, outcome.error).toBe(true)
    outcome = executeOps(editor, [{ op: 'updateFields' }])
    expect(outcome.ok, outcome.error).toBe(true)
    const nodes: Array<{ text: string; bold: boolean; dirty: unknown }> = []
    editor.state.doc.descendants((node) => {
      const instr = node.marks.find((m) => m.type.name === 'instrField')
      if (node.isText && instr)
        nodes.push({
          text: node.text!,
          bold: node.marks.some((m) => m.type.name === 'bold'),
          dirty: instr.attrs.dirty,
        })
      return true
    })
    expect(nodes).toEqual([
      { text: 'op', bold: true, dirty: true },
      { text: 'en', bold: false, dirty: true },
    ])
  })

  it('a REF inside its own bookmark reads the paragraph without itself and stays stable', () => {
    const editor = createEditor([para('Chapter one.', { bookmarks: ['Self'] })])
    let outcome = executeOps(editor, [
      { op: 'insertField', target: at(0), type: 'REF', args: 'Self' },
      { op: 'insertField', target: at(0), type: 'REF', args: 'Self', position: 'start' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    const refs = () => fieldsOf(editor).map((f) => f.text)
    expect(refs()).toEqual(['Chapter one.', 'Chapter one.'])
    for (let i = 0; i < 2; i++) {
      outcome = executeOps(editor, [{ op: 'updateFields' }])
      expect(outcome.ok, outcome.error).toBe(true)
      expect(refs()).toEqual(['Chapter one.', 'Chapter one.'])
    }
    expect(editor.state.doc.textContent).toBe('Chapter one.Chapter one.Chapter one.')
  })

  it('renumbers SEQ fields, refreshes REF and DATE, and marks the rest dirty', () => {
    const editor = createEditor([
      para('Intro.', { bookmarks: ['Intro'] }),
      para('Figure '),
      para('Figure '),
      para('Figure '),
      para('See '),
      para('Pages: '),
    ])
    let outcome = executeOps(editor, [
      { op: 'insertField', target: at(1), type: 'SEQ', args: 'Figure' },
      { op: 'insertField', target: at(2), type: 'SEQ', args: 'Figure' },
      { op: 'insertField', target: at(3), type: 'SEQ', args: 'Figure' },
      { op: 'insertField', target: at(4), type: 'REF', args: 'Intro' },
      { op: 'insertField', target: at(5), type: 'NUMPAGES', result: '9' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    // drop the middle figure and rename the bookmarked paragraph, then refresh
    outcome = executeOps(editor, [
      { op: 'deleteBlocks', target: at(2) },
      { op: 'findReplace', target: at(0), find: 'Intro.', replace: 'Introduction.' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    // NUMPAGES was inserted dirty; clear it to prove updateFields re-marks it
    const numPages = fieldsOf(editor).find((f) => f.instr === 'NUMPAGES')
    expect(numPages?.dirty).toBe(true)
    outcome = executeOps(editor, [{ op: 'updateFields' }])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(fieldsOf(editor)).toEqual([
      { text: '1', instr: 'SEQ Figure', dirty: true },
      { text: '2', instr: 'SEQ Figure', dirty: true },
      { text: 'Introduction.', ref: 'Intro' },
      { text: '9', instr: 'NUMPAGES', dirty: true },
    ])
    expect(outcome.results[0]!.detail).toContain('3 recomputed')
    const seqOnly = executeOps(editor, [{ op: 'updateFields', scope: 'SEQ' }])
    expect(seqOnly.ok).toBe(true)
    const bad = executeOps(editor, [{ op: 'updateFields', scope: 'PAGE' }])
    expect(bad.error).toContain('scope must be one of')
  })
  it('treats a field result split by formatting as one field', () => {
    const seq = { type: 'instrField', attrs: { instr: 'SEQ Figure \\r 10', dirty: true } }
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: null },
        content: [
          text('Figure '),
          { type: 'text', text: '1', marks: [{ type: 'bold' }, seq] },
          { type: 'text', text: '0', marks: [seq] },
        ],
      },
      para('Figure '),
    ])
    const next = executeOps(editor, [
      { op: 'insertField', target: at(1), type: 'SEQ', args: 'Figure' },
    ])
    expect(next.ok, next.error).toBe(true)
    expect(editor.state.doc.child(1).textContent).toBe('Figure 2')
    const update = executeOps(editor, [{ op: 'updateFields' }])
    expect(update.ok, update.error).toBe(true)
    expect(editor.state.doc.child(0).textContent).toBe('Figure 10')
    expect(editor.state.doc.child(1).textContent).toBe('Figure 11')
    // the result did not change, so the bold half keeps its own node
    expect(fieldsOf(editor)).toEqual([
      { text: '1', instr: 'SEQ Figure \\r 10', dirty: true },
      { text: '0', instr: 'SEQ Figure \\r 10', dirty: true },
      { text: '11', instr: 'SEQ Figure', dirty: true },
    ])
    expect(update.results[0]!.detail).toContain('2 recomputed')
  })
})

describe('field helpers', () => {
  it('formats SEQ numbers and date pictures like Word', () => {
    expect(formatSeqNumber(4, undefined)).toBe('4')
    expect(formatSeqNumber(4, 'ROMAN')).toBe('IV')
    expect(formatSeqNumber(27, 'alphabetic')).toBe('aa')
    const d = new Date(2026, 8, 15, 14, 5, 9)
    expect(formatFieldDate(d, 'dddd, MMMM d, yyyy')).toBe('Tuesday, September 15, 2026')
    expect(formatFieldDate(d, "yyyy-MM-dd 'at' h:mm am/pm")).toBe('2026-09-15 at 2:05 pm')
    expect(evaluateIf('"a" = "a" "yes" "no"')).toBe('yes')
    // measured in Word: unquoted numbers compare as numbers, quoted ones as text
    expect(evaluateIf('10 < 9 "yes" "no"')).toBe('no')
    expect(evaluateIf('010 = 10 "yes" "no"')).toBe('yes')
    expect(evaluateIf('"10" < "9" "yes" "no"')).toBe('yes')
    expect(evaluateIf('"010" = "10" "yes" "no"')).toBe('no')
    // text is case-sensitive but ordered alphabetically; = / <> take ? and * in the second operand
    expect(evaluateIf('"ABC" = "abc" "yes" "no"')).toBe('no')
    expect(evaluateIf('"a" < "B" "yes" "no"')).toBe('yes')
    expect(evaluateIf('"Zed" > "apple" "yes" "no"')).toBe('yes')
    expect(evaluateIf('"abc" = "a*" "yes" "no"')).toBe('yes')
    expect(evaluateIf('"abc" <> "A*" "yes" "no"')).toBe('yes')
    expect(evaluateIf('"abc" = "a?c" "yes" "no"')).toBe('yes')
    expect(evaluateIf('"a.c" = "abc" "yes" "no"')).toBe('no')
    expect(evaluateIf('{ MERGEFIELD x } = 1 "y" "n"')).toBeUndefined()
  })
})

describe('field ops through the docx save path', () => {
  it('a REF with a Word-computed switch keeps its instruction and dirty flag through reopen', async () => {
    const source = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>Summary paragraph.</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r></w:p>',
    })
    const save = async (editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) => {
      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      const out = await saveDocx(parsed, plan.saveBlocks)
      return {
        out,
        xml: await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('string'),
      }
    }
    const dirtyRef =
      /w:dirty="true"\/><\/w:r><w:r><w:instrText xml:space="preserve"> REF Summary \\p </
    const parsed = await parseDocx(source)
    const editor = createEditor(blocksToPmDoc(parsed.blocks).content as JsonNode[])
    let outcome = executeOps(editor, [
      { op: 'insertBookmark', target: at(0), name: 'Summary' },
      { op: 'insertField', target: at(1), type: 'REF', args: 'Summary \\p' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    const first = await save(editor, parsed)
    expect(first.xml).toMatch(dirtyRef)
    const reparsed = await parseDocx(first.out)
    expect(reparsed.blocks[1]!.runs!.find((r) => r.refField === 'Summary')).toMatchObject({
      refInstr: ' REF Summary \\p ',
      fldDirty: true,
    })
    // the bookmark text changes; a \p result is Word's, so updateFields leaves it and keeps the flag
    const reopened = createEditor(blocksToPmDoc(reparsed.blocks).content as JsonNode[])
    outcome = executeOps(reopened, [
      { op: 'findReplace', target: at(0), find: 'Summary paragraph.', replace: 'Changed.' },
      { op: 'updateFields' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(fieldsOf(reopened)).toEqual([{ text: 'Summary paragraph.', ref: 'Summary' }])
    expect((await save(reopened, reparsed)).xml).toMatch(dirtyRef)
  })

  it('writes dirty complex fields and a bookmark Word can resolve', async () => {
    const source = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>Summary paragraph.</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t xml:space="preserve">Figure </w:t></w:r></w:p>' +
        '<w:p><w:r><w:t xml:space="preserve">Back to </w:t></w:r></w:p>' +
        '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r></w:p>',
    })
    const parsed = await parseDocx(source)
    const editor = createEditor(blocksToPmDoc(parsed.blocks).content as JsonNode[])
    const outcome = executeOps(editor, [
      { op: 'insertBookmark', target: at(0), name: 'Summary' },
      { op: 'insertField', target: at(1), type: 'SEQ', args: 'Figure \\* ARABIC' },
      { op: 'insertField', target: at(2), type: 'REF', args: 'Summary \\h' },
      { op: 'insertField', target: at(3), type: 'PAGEREF', args: 'Summary \\h' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const out = await saveDocx(parsed, plan.saveBlocks)
    const xml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('string')
    expect(xml).toContain('w:name="Summary"')
    expect(xml).toMatch(
      /<w:fldChar w:fldCharType="begin" w:dirty="true"\/><\/w:r><w:r><w:instrText xml:space="preserve"> SEQ Figure \\\* ARABIC <\/w:instrText>/,
    )
    expect(xml).toMatch(/ REF Summary \\h <\/w:instrText>[\s\S]*Summary paragraph\./)
    expect(xml).toMatch(/w:dirty="true"[\s\S]{0,80} PAGEREF Summary \\h /)
    const reparsed = await parseDocx(out)
    expect(reparsed.blocks[0]!.bookmarks).toEqual(['Summary'])
    expect(reparsed.blocks[2]!.runs?.some((r) => r.refField === 'Summary')).toBe(true)
    // PAGEREF folds back into an editable inline field carrying the dirty flag
    expect(reparsed.blocks[3]!.type).toBe('paragraph')
    expect(reparsed.blocks[3]!.runs?.find((r) => r.instrField)).toMatchObject({
      text: '1',
      instrField: 'PAGEREF Summary \\h',
      fldDirty: true,
    })
  })
})
