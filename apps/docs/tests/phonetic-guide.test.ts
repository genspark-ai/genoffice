/** Phonetic Guide: a reading goes on selected text, is changed on a guide, and comes off again. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyRuby, rubyTargetOf } from '../src/renderer/editor/phonetic-guide'

describe('phonetic guide', () => {
  let editor: Editor
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'see ' },
              { type: 'text', text: '\u6708\u4eae', marks: [{ type: 'bold' }] },
              { type: 'text', text: ' rise' },
            ],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'second' }] },
        ],
      },
    })
  })
  afterEach(() => editor.destroy())

  const rubies = () => {
    const out: Array<{ base: string; rt: string; xml: string; bold: boolean }> = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docRuby')
        out.push({
          base: node.attrs.base,
          rt: node.attrs.rt,
          xml: node.attrs.xml,
          bold: node.marks.some((m) => m.type.name === 'bold'),
        })
    })
    return out
  }

  it('takes selected text inside one paragraph, and nothing else', () => {
    editor.commands.setTextSelection({ from: 5, to: 7 })
    expect(rubyTargetOf(editor.state)).toEqual({
      from: 5,
      to: 7,
      base: '\u6708\u4eae',
      rt: '',
      align: null,
      reading: { font: '', sizePt: 5.5, offsetPt: 0 },
    })
    // a caret, and a selection running into the next paragraph, are not text to annotate
    editor.commands.setTextSelection(5)
    expect(rubyTargetOf(editor.state)).toBeNull()
    editor.commands.setTextSelection({ from: 5, to: 16 })
    expect(rubyTargetOf(editor.state)).toBeNull()
  })

  it('puts the reading on the text and keeps its formatting', () => {
    editor.commands.setTextSelection({ from: 5, to: 7 })
    expect(applyRuby(editor, rubyTargetOf(editor.state)!, { rt: ' yu\u00e8liang ' }, 'zh-CN')).toBe(
      true,
    )
    const [ruby] = rubies()
    expect(ruby).toMatchObject({ base: '\u6708\u4eae', rt: 'yu\u00e8liang', bold: true })
    expect(ruby.xml).toContain('<w:rubyBase>')
    expect(ruby.xml).toContain('w:lid w:val="zh-CN"')
    // the paragraph still reads the same
    expect(editor.getText()).toContain('see \u6708\u4eae rise')
  })

  it('changes the reading of a selected guide, and an empty one takes the guide off', () => {
    editor.commands.setTextSelection({ from: 5, to: 7 })
    applyRuby(editor, rubyTargetOf(editor.state)!, { rt: 'yue' }, 'zh-CN')
    // the new guide is left selected
    const guide = rubyTargetOf(editor.state)!
    expect(guide).toMatchObject({ base: '\u6708\u4eae', rt: 'yue' })
    applyRuby(editor, guide, { rt: 'yu\u00e8' }, 'zh-CN')
    expect(rubies().map((r) => r.rt)).toEqual(['yu\u00e8'])

    editor.commands.setNodeSelection(5)
    applyRuby(editor, rubyTargetOf(editor.state)!, { rt: '' }, 'zh-CN')
    expect(rubies()).toEqual([])
    expect(editor.state.doc.firstChild!.textContent).toBe('see \u6708\u4eae rise')
    // the base text has its formatting back
    expect(editor.state.doc.nodeAt(5)!.marks.some((m) => m.type.name === 'bold')).toBe(true)
  })

  it('changing the reading of a guide keeps the rest of what the guide says', () => {
    // a guide as a file brings it: its own alignment and a formatted reading
    const fromFile =
      '<w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="14"/></w:rubyPr>' +
      '<w:rt><w:r><w:rPr><w:i/></w:rPr><w:t>old</w:t></w:r></w:rt>' +
      '<w:rubyBase><w:r><w:rPr><w:b/></w:rPr><w:t>base</w:t></w:r></w:rubyBase></w:ruby>'
    editor.commands.insertContentAt(1, {
      type: 'docRuby',
      attrs: { base: 'base', rt: 'old', xml: fromFile },
    } as never)
    editor.commands.setNodeSelection(1)
    applyRuby(editor, rubyTargetOf(editor.state)!, { rt: 'new' }, 'zh-CN')
    const [guide] = rubies()
    expect(guide.rt).toBe('new')
    expect(guide.xml).toBe(
      fromFile.replace('<w:t>old</w:t>', '<w:t xml:space="preserve">new</w:t>'),
    )
  })

  it('a new guide takes the alignment picked, and a guide changes its own', () => {
    editor.commands.setTextSelection({ from: 5, to: 7 })
    applyRuby(editor, rubyTargetOf(editor.state)!, { rt: 'yue', align: 'left' }, 'zh-CN')
    expect(rubies()[0].xml).toContain('<w:rubyAlign w:val="left"/>')
    const guide = rubyTargetOf(editor.state)!
    expect(guide.align).toBe('left')
    applyRuby(editor, guide, { rt: 'yue', align: 'distributeLetter' }, 'zh-CN')
    expect(rubies()[0].xml).toContain('<w:rubyAlign w:val="distributeLetter"/>')
    expect(rubies()[0].xml).not.toContain('w:val="left"')
    expect(rubyTargetOf(editor.state)!.align).toBe('distributeLetter')
  })

  it('a flush-right guide draws its base as its own box, the other alignments as bare text', () => {
    editor.commands.setTextSelection({ from: 5, to: 7 })
    editor.commands.setMark('underline')
    editor.commands.setMark('strike')
    applyRuby(editor, rubyTargetOf(editor.state)!, { rt: 'yue', align: 'right' }, 'zh-CN')
    expect(rubies()[0].xml).toContain('<w:rubyAlign w:val="right"/>')
    const drawn = () => editor.view.dom.querySelector('ruby.doc-ruby')!
    expect(drawn().getAttribute('data-ruby-align')).toBe('right')
    const base = drawn().querySelector<HTMLElement>(':scope > .doc-ruby-base')!
    expect(base.textContent).toBe('\u6708\u4eae')
    // the marks' lines do not reach inside the box, so the base carries them itself
    expect(base.style.textDecorationLine).toBe('underline line-through')
    expect(drawn().querySelector('rt')!.textContent).toBe('yue')
    applyRuby(editor, rubyTargetOf(editor.state)!, { rt: 'yue', align: 'center' }, 'zh-CN')
    expect(drawn().querySelector('.doc-ruby-base')).toBeNull()
    expect(drawn().firstChild!.nodeType).toBe(Node.TEXT_NODE)
  })

  it('changes the base text of a guide and keeps the formatting of its base run', () => {
    const fromFile =
      '<w:ruby><w:rubyPr><w:rubyAlign w:val="center"/></w:rubyPr>' +
      '<w:rt><w:r><w:t>old</w:t></w:r></w:rt>' +
      '<w:rubyBase><w:r><w:rPr><w:b/></w:rPr><w:t>base</w:t></w:r></w:rubyBase></w:ruby>'
    editor.commands.insertContentAt(1, {
      type: 'docRuby',
      attrs: { base: 'base', rt: 'old', xml: fromFile },
    } as never)
    editor.commands.setNodeSelection(1)
    applyRuby(editor, rubyTargetOf(editor.state)!, { rt: 'old', base: 'words' }, 'zh-CN')
    const [guide] = rubies()
    expect(guide.base).toBe('words')
    expect(guide.xml).toContain(
      '<w:rubyBase><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">words</w:t></w:r></w:rubyBase>',
    )
    expect(editor.getText()).toContain('words')
  })

  it('a base typed over selected text replaces it under the new guide', () => {
    editor.commands.setTextSelection({ from: 5, to: 7 })
    applyRuby(editor, rubyTargetOf(editor.state)!, { rt: 'ri', base: '\u65e5' }, 'zh-CN')
    expect(rubies().map((r) => [r.base, r.rt, r.bold])).toEqual([['\u65e5', 'ri', true]])
    expect(editor.getText()).toContain('see \u65e5 rise')
  })

  it('leaves plain text alone when no reading is given', () => {
    editor.commands.setTextSelection({ from: 5, to: 7 })
    expect(applyRuby(editor, rubyTargetOf(editor.state)!, { rt: '  ' }, 'zh-CN')).toBe(false)
    expect(rubies()).toEqual([])
  })

  it('a guide takes the font, size and distance of its reading, and reads them back', () => {
    editor.commands.setTextSelection({ from: 5, to: 7 })
    applyRuby(
      editor,
      rubyTargetOf(editor.state)!,
      { rt: 'yue', reading: { font: 'Arial', sizePt: 8, offsetPt: 3 } },
      'zh-CN',
    )
    const made = rubies()[0].xml
    expect(made).toContain('<w:hps w:val="16"/><w:hpsRaise w:val="25"/><w:hpsBaseText w:val="21"/>')
    expect(made).toContain('w:ascii="Arial"')
    expect(rubyTargetOf(editor.state)!.reading).toEqual({ font: 'Arial', sizePt: 8, offsetPt: 3 })
    const rt = editor.view.dom.querySelector('ruby rt') as HTMLElement
    expect(rt.style.fontSize).toBe('8pt')
    expect(rt.style.bottom).toBe('3pt')
    expect(rt.style.fontFamily).toContain('Arial')
    // only the distance changes: the rest of the guide stays as it was
    applyRuby(editor, rubyTargetOf(editor.state)!, { rt: 'yue', reading: { offsetPt: 0 } }, 'zh-CN')
    expect(rubies()[0].xml).toBe(made.replace('w:hpsRaise w:val="25"', 'w:hpsRaise w:val="19"'))
    expect((editor.view.dom.querySelector('ruby rt') as HTMLElement).style.bottom).toBe('')
  })
})
