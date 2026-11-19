import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { executeOps } from '../src/renderer/ai/ops'
import { serializeRangeToHtml } from '../src/renderer/ai/protocol'
import { setModuleLang } from '../src/renderer/i18n/locale'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const run = (
  t: string,
  style?: Record<string, unknown>,
  extra: JsonNode['marks'] = [],
): JsonNode => ({
  type: 'text',
  text: t,
  ...(style || extra.length
    ? { marks: [...(style ? [{ type: 'docTextStyle', attrs: style }] : []), ...extra] }
    : {}),
})
const para = (content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
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

function runs(editor: Editor): Array<{ text: string; color: unknown; size: unknown }> {
  const out: Array<{ text: string; color: unknown; size: unknown }> = []
  editor.state.doc.descendants((n) => {
    if (!n.isText) return true
    const ts = n.marks.find((m) => m.type.name === 'docTextStyle')
    out.push({
      text: n.text ?? '',
      color: ts?.attrs.color ?? null,
      size: ts?.attrs.sizeHalfPoints ?? null,
    })
    return true
  })
  return out
}

setModuleLang('en')

describe('read_blocks shows color and size only where a run departs from its block', () => {
  it('a uniformly styled block reads as plain text', () => {
    const editor = createEditor([
      para([
        run('All ', { color: '333333', sizeHalfPoints: 24 }),
        run('same', { color: '333333', sizeHalfPoints: 24 }),
      ]),
    ])
    expect(serializeRangeToHtml(editor, 0, 0)).toBe('<p>All same</p>')
  })

  it('a red correction and an oversized word get a read-only span', () => {
    const editor = createEditor([
      para([
        run('Body text ', { sizeHalfPoints: 21 }),
        run('fixed', { color: 'FF0000', sizeHalfPoints: 21 }, [{ type: 'bold' }]),
        run(' and ', { sizeHalfPoints: 21 }),
        run('BIG', { sizeHalfPoints: 36 }),
      ]),
    ])
    expect(serializeRangeToHtml(editor, 0, 0)).toBe(
      '<p>Body text <span style="color:#FF0000"><strong>fixed</strong></span> and <span style="font-size:18pt">BIG</span></p>',
    )
  })

  it("Word's explicit 'auto' color reads as no color, and deleted runs do not set the baseline", () => {
    const editor = createEditor([
      para([
        run('old red text that is gone', { color: 'FF0000' }, [
          { type: 'del', attrs: { author: 'a' } },
        ]),
        run('plain', { color: 'auto' }),
        run(' text', {}),
      ]),
    ])
    expect(serializeRangeToHtml(editor, 0, 0)).toBe('<p>plain text</p>')
  })

  it('a block whose dominant color is red shows the normal run as auto', () => {
    const editor = createEditor([
      para([run('Mostly red text here', { color: 'FF0000' }), run(' ok', {})]),
    ])
    expect(serializeRangeToHtml(editor, 0, 0)).toBe(
      '<p>Mostly red text here<span style="color:auto"> ok</span></p>',
    )
  })
})

describe('setMatchedFont by current color / size', () => {
  const doc = () => [
    para([
      run('Keep ', { sizeHalfPoints: 21 }),
      run('red', { color: 'FF0000', sizeHalfPoints: 21 }),
      run(' and ', { sizeHalfPoints: 21 }),
      run('link', { color: '0563C1', sizeHalfPoints: 21 }, [
        { type: 'link', attrs: { href: 'https://x', rId: null } },
      ]),
    ]),
    para([
      run('small note', { sizeHalfPoints: 18 }),
      run(' red too', { color: 'FF0000', sizeHalfPoints: 18 }),
    ]),
  ]

  it('matchColor restyles only runs of that color, leaving links and other colors alone', () => {
    const editor = createEditor(doc())
    const outcome = executeOps(editor, [
      { op: 'setMatchedFont', matchColor: '#ff0000', color: null },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(outcome.results[0]).toMatchObject({ changed: 2, detail: '2' })
    // cleared runs merge with their plain neighbours; the link keeps its own color
    expect(runs(editor).map((r) => [r.text, r.color])).toEqual([
      ['Keep red and ', null],
      ['link', '0563C1'],
      ['small note red too', null],
    ])
  })

  it('matchFontSize changes the size of matching runs only', () => {
    const editor = createEditor(doc())
    const outcome = executeOps(editor, [{ op: 'setMatchedFont', matchFontSize: 9, fontSize: 10.5 }])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(runs(editor).map((r) => r.size)).toEqual([21, 21, 21, 21, 21, 21])
  })

  it('text and matchColor combine: only red occurrences of the word', () => {
    const editor = createEditor([
      para([run('red word', { color: 'FF0000' }), run(' plain word', {})]),
    ])
    executeOps(editor, [{ op: 'setMatchedFont', text: 'word', matchColor: '#FF0000', bold: true }])
    const boldRuns: string[] = []
    editor.state.doc.descendants((n) => {
      if (n.isText && n.marks.some((m) => m.type.name === 'bold')) boldRuns.push(n.text ?? '')
      return true
    })
    expect(boldRuns).toEqual(['word'])
  })

  it('matchColor "auto" targets runs without an explicit color', () => {
    const editor = createEditor(doc())
    const outcome = executeOps(editor, [
      { op: 'setMatchedFont', matchColor: 'auto', color: '#000000' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(runs(editor).map((r) => r.color)).toEqual([
      '000000',
      'FF0000',
      '000000',
      '0563C1',
      '000000',
      'FF0000',
    ])
  })

  it("matchColor 'auto' also catches Word's explicit auto; matchFontSize 'default' catches unsized runs", () => {
    const editor = createEditor([
      para([run('explicit auto', { color: 'auto' }), run(' red', { color: 'FF0000' })]),
      para([run('unsized', {}), run(' sized', { sizeHalfPoints: 24 })]),
    ])
    expect(executeOps(editor, [{ op: 'setMatchedFont', matchColor: 'auto', bold: true }]).ok).toBe(
      true,
    )
    expect(
      executeOps(editor, [{ op: 'setMatchedFont', matchFontSize: 'default', italic: true }]).ok,
    ).toBe(true)
    const marks = (t: string) => {
      let found: string[] = []
      editor.state.doc.descendants((n) => {
        if (n.isText && n.text === t) found = n.marks.map((m) => m.type.name).sort()
        return true
      })
      return found
    }
    expect(marks('explicit auto')).toContain('bold')
    expect(marks(' red')).not.toContain('bold')
    expect(marks('unsized')).toContain('italic')
    expect(marks(' sized')).not.toContain('italic')
  })

  it('a selection that cuts into a matching run restyles only the selected part', () => {
    const editor = createEditor([para([run('aaaa', { color: 'FF0000' }), run('bbbb', {})])])
    // select "aabb": the last two red characters plus two plain ones
    editor.commands.setTextSelection({ from: 3, to: 7 })
    const outcome = executeOps(editor, [
      { op: 'setMatchedFont', matchColor: '#FF0000', target: { scope: 'selection' }, color: null },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(runs(editor).map((r) => [r.text, r.color])).toEqual([
      ['aa', 'FF0000'],
      ['aa', null],
      ['bbbb', null],
    ])
  })

  it('rejects an op with nothing to match or a bad matchColor', () => {
    const editor = createEditor(doc())
    expect(executeOps(editor, [{ op: 'setMatchedFont', bold: true }]).error).toMatch(
      /what to match/,
    )
    expect(
      executeOps(editor, [{ op: 'setMatchedFont', matchColor: 'red', bold: true }]).error,
    ).toMatch(/matchColor/)
    expect(
      executeOps(editor, [{ op: 'setMatchedFont', matchFontSize: null, bold: true }]).error,
    ).toMatch(/matchFontSize/)
  })
})
