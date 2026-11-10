/**
 * Web-HTML paste font handling (r181): whole-paragraph web copies carry the
 * site's styles on BLOCK elements (the mark parser reads only spans), and
 * generic-only family chains leave runs font-less — both pasted in the theme
 * font instead of the surrounding text's font.
 */
import { Editor } from '@tiptap/core'
import { DOMParser as PmDOMParser, Fragment } from '@tiptap/pm/model'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  caretFontSlots,
  fillMissingRunFonts,
  isForeignPasteHtml,
  pushDownWebInlineStyles,
} from '../src/renderer/editor/paste-web-html'

const makeEditor = () =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [{ type: 'docParagraph' }] },
  })

const parseHtml = (editor: Editor, html: string) => {
  const dom = new window.DOMParser().parseFromString(html, 'text/html')
  return PmDOMParser.fromSchema(editor.schema).parse(dom.body)
}

const runStyles = (doc: ReturnType<typeof parseHtml>) => {
  const runs: Array<{ text: string; attrs: Record<string, unknown> | null }> = []
  doc.descendants((node) => {
    if (node.isText) {
      const style = node.marks.find((m) => m.type.name === 'docTextStyle')
      runs.push({ text: node.text ?? '', attrs: style ? style.attrs : null })
    }
    return true
  })
  return runs
}

describe('isForeignPasteHtml', () => {
  it('treats browser fragments as foreign and our clipboard HTML as internal', () => {
    expect(isForeignPasteHtml('<p style="font-family:Roboto">x</p>')).toBe(true)
    expect(isForeignPasteHtml('<div data-pm-slice="1 1 []"><p>x</p></div>')).toBe(false)
    expect(isForeignPasteHtml('<span data-doc-style="{}">x</span>')).toBe(false)
  })
})

describe('pushDownWebInlineStyles', () => {
  it('moves block-level font/size/color onto a span the mark parser reads', () => {
    const editor = makeEditor()
    const html = pushDownWebInlineStyles(
      '<p style="margin:0 0 16px;color:rgb(32,33,34);font-family:Roboto, sans-serif;font-size:14px;line-height:22.4px">premier paragraphe</p>',
    )
    const runs = runStyles(parseHtml(editor, html))
    expect(runs).toHaveLength(1)
    expect(runs[0]!.attrs).toMatchObject({
      font: 'Roboto',
      fontAscii: 'Roboto',
      sizeHalfPoints: 21,
      color: '202122',
    })
    editor.destroy()
  })

  it('keeps an existing span’s own declarations and only fills the missing ones', () => {
    const editor = makeEditor()
    const html = pushDownWebInlineStyles(
      '<p style="font-family:Roboto;font-size:14px"><span style="font-family:Georgia">cited</span></p>',
    )
    const runs = runStyles(parseHtml(editor, html))
    expect(runs[0]!.attrs).toMatchObject({ font: 'Georgia', sizeHalfPoints: 21 })
    editor.destroy()
  })

  it('reaches text in nested and shared spans', () => {
    const editor = makeEditor()
    const html = pushDownWebInlineStyles(
      '<p style="font-family:Roboto"><span>plain<b>bold</b></span><a href="#"><span style="color:rgb(0,0,255)">link</span></a></p>',
    )
    const runs = runStyles(parseHtml(editor, html))
    expect(runs.map((r) => r.text)).toEqual(['plain', 'bold', 'link'])
    for (const run of runs) expect(run.attrs).toMatchObject({ font: 'Roboto' })
    expect(runs[2]!.attrs).toMatchObject({ color: '0000FF' })
    editor.destroy()
  })

  it('leaves table subtrees alone', () => {
    const html =
      '<table style="font-family:Arial"><tbody><tr><td>a</td><td>b</td></tr></tbody></table>'
    expect(pushDownWebInlineStyles(html)).toBe(html)
  })

  it('returns the input unchanged when nothing declares pushdown styles', () => {
    const html = '<p>bare text</p>'
    expect(pushDownWebInlineStyles(html)).toBe(html)
  })
})

describe('caretFontSlots + fillMissingRunFonts', () => {
  it('fills font-less pasted runs with the caret font and keeps concrete fonts', () => {
    const editor = makeEditor()
    const schema = editor.schema
    const markType = schema.marks.docTextStyle!
    const caret = [markType.create({ font: 'Calibri', fontAscii: 'Calibri', color: 'FF0000' })]
    const slots = caretFontSlots(caret, markType)
    // only the font slots travel — the caret's color must not leak
    expect(slots).toEqual({ font: 'Calibri', fontAscii: 'Calibri' })

    const paragraph = schema.nodes.docParagraph!.create(null, [
      schema.text('generic ', [markType.create({ color: '202122', sizeHalfPoints: 21 })]),
      schema.text('roboto', [markType.create({ font: 'Roboto', fontAscii: 'Roboto' })]),
      schema.text(' bare'),
    ])
    const filled = fillMissingRunFonts(Fragment.from(paragraph), slots!, schema)
    const runs: Array<Record<string, unknown> | null> = []
    filled.forEach((node) =>
      node.descendants((child) => {
        if (child.isText) {
          const style = child.marks.find((m) => m.type === markType)
          runs.push(style ? style.attrs : null)
        }
        return true
      }),
    )
    // font-less run with other styling: gains the caret font, keeps its own attrs
    expect(runs[0]).toMatchObject({
      font: 'Calibri',
      fontAscii: 'Calibri',
      color: '202122',
      sizeHalfPoints: 21,
    })
    // concrete source font is source formatting — untouched
    expect(runs[1]).toMatchObject({ font: 'Roboto' })
    expect((runs[1] as Record<string, unknown>).color ?? null).toBeNull()
    // completely unmarked text gains a font-only mark
    expect(runs[2]).toMatchObject({ font: 'Calibri', fontAscii: 'Calibri' })
    editor.destroy()
  })

  it('leaves table cell runs font-less like the pushdown does', () => {
    const editor = makeEditor()
    const schema = editor.schema
    const markType = schema.marks.docTextStyle!
    const table = parseHtml(
      editor,
      '<table><tbody><tr><td>cell</td></tr></tbody></table><p>after</p>',
    ).content
    const filled = fillMissingRunFonts(table, { font: 'Calibri', fontAscii: 'Calibri' }, schema)
    const fonts: Array<string | null> = []
    filled.forEach((node) =>
      node.descendants((child) => {
        if (child.isText) {
          const style = child.marks.find((m) => m.type === markType)
          fonts.push((style?.attrs.font as string | undefined) ?? null)
        }
        return true
      }),
    )
    expect(fonts).toEqual([null, 'Calibri'])
    editor.destroy()
  })

  it('returns null slots when the caret text has no explicit font', () => {
    const editor = makeEditor()
    const markType = editor.schema.marks.docTextStyle!
    expect(caretFontSlots([], markType)).toBeNull()
    expect(caretFontSlots([markType.create({ color: '202122' })], markType)).toBeNull()
    editor.destroy()
  })
})
