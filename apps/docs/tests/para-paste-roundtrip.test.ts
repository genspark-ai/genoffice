/**
 * Paragraph formatting must survive the clipboard HTML round-trip (alpha
 * ledger r117): copy/paste re-parses renderHTML output, and the paragraph CSS
 * (margin-inline-start etc.) is never parsed back — so before the data-para
 * payload existed, a paste dropped every paragraph attr. In a document whose
 * body paragraphs carry a NEGATIVE left indent, the pasted paragraph then sat
 * visibly to the RIGHT of its siblings.
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'

function roundtripPara(
  attrs: Record<string, unknown>,
  text = 'bassiste de David Bowie',
): Record<string, unknown> {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', attrs, content: [{ type: 'text', text }] }],
    },
  })
  const html = editor.getHTML()
  editor.commands.setContent(html)
  const out = { ...editor.state.doc.child(0).attrs }
  editor.destroy()
  return out
}

function parseHtml(html: string): { para: Record<string, unknown>; editor: Editor } {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: html,
  })
  const para = { ...editor.state.doc.child(0).attrs }
  editor.destroy()
  return { para, editor }
}

describe('paragraph formatting HTML round-trip (r117)', () => {
  it('keeps a negative left indent', () => {
    expect(roundtripPara({ indentLeft: -284 })).toMatchObject({ indentLeft: -284 })
  })

  it('keeps align, indents, spacing and style', () => {
    expect(
      roundtripPara({
        styleId: 'Corps',
        align: 'justify',
        indentLeft: -284,
        indentRight: -142,
        indentFirstLine: 708,
        spaceBefore: 120,
        spaceAfter: 0,
        lineSpacing: 1.5,
      }),
    ).toMatchObject({
      styleId: 'Corps',
      align: 'justify',
      indentLeft: -284,
      indentRight: -142,
      indentFirstLine: 708,
      spaceBefore: 120,
      spaceAfter: 0,
      lineSpacing: 1.5,
    })
  })

  it('keeps explicit opt-out booleans (snapToGrid/autoSpace false)', () => {
    expect(roundtripPara({ snapToGrid: false, autoSpace: false })).toMatchObject({
      snapToGrid: false,
      autoSpace: false,
    })
  })

  it('does NOT resurrect identity/anchor attrs — a pasted paragraph is new content', () => {
    const out = roundtripPara({
      docxIndex: 7,
      bookmarks: ['target1'],
      indentLeft: -284,
      pPrChange: '{"author":"x"}',
    })
    expect(out.indentLeft).toBe(-284)
    expect(out.docxIndex).toBeNull()
    expect(out.bookmarks).toBeNull()
    expect(out.pPrChange).toBeNull()
  })

  it('keeps formatting on headings and list items', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docHeading',
            attrs: { level: 2, indentLeft: -284, align: 'center' },
            content: [{ type: 'text', text: 'Titre' }],
          },
          {
            type: 'docListItem',
            attrs: { kind: 'ordered', ilvl: 0, indentRight: 240, spaceAfter: 60 },
            content: [{ type: 'text', text: 'item' }],
          },
        ],
      },
    })
    const html = editor.getHTML()
    editor.commands.setContent(html)
    expect({ ...editor.state.doc.child(0).attrs }).toMatchObject({
      level: 2,
      indentLeft: -284,
      align: 'center',
    })
    // before r117, div.doc-li had NO parse rule: a pasted list item fell apart
    expect(editor.state.doc.child(1).type.name).toBe('docListItem')
    expect({ ...editor.state.doc.child(1).attrs }).toMatchObject({
      kind: 'ordered',
      ilvl: 0,
      indentRight: 240,
      spaceAfter: 60,
    })
    editor.destroy()
  })

  it('ignores malformed or wrong-typed data-para payloads', () => {
    expect(parseHtml(`<p data-para="not json">x</p>`).para.indentLeft).toBeNull()
    expect(
      parseHtml(`<p data-para='{"indentLeft":"evil","docxIndex":9}'>x</p>`).para,
    ).toMatchObject({ indentLeft: null, docxIndex: null })
    expect(parseHtml(`<p data-para='{"indentLeft":-284}'>x</p>`).para.indentLeft).toBe(-284)
  })

  it('foreign HTML without data-para still parses as a plain paragraph', () => {
    expect(parseHtml('<p style="margin-left:20pt">x</p>').para).toMatchObject({
      indentLeft: null,
      docxIndex: null,
    })
  })
})

describe('docTextStyle exact clipboard round-trip (r117)', () => {
  function roundtripMark(attrs: Record<string, unknown>): Record<string, unknown> | undefined {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [{ type: 'text', text: 'texte', marks: [{ type: 'docTextStyle', attrs }] }],
          },
        ],
      },
    })
    const html = editor.getHTML()
    editor.commands.setContent(html)
    const mark = editor.state.doc
      .child(0)
      .child(0)
      .marks.find((m) => m.type.name === 'docTextStyle')
    const out = mark ? { ...mark.attrs } : undefined
    editor.destroy()
    return out
  }

  it('keeps attrs the CSS heuristics used to lose', () => {
    expect(
      roundtripMark({
        highlight: 'yellow',
        shading: 'D9D9D9',
        caps: 'small',
        em: 'dot',
        charScaleEm: -0.02,
      }),
    ).toMatchObject({
      highlight: 'yellow',
      shading: 'D9D9D9',
      caps: 'small',
      em: 'dot',
      charScaleEm: -0.02,
    })
  })

  it('keeps color/size/fonts exactly', () => {
    expect(
      roundtripMark({
        color: 'C00000',
        sizeHalfPoints: 21,
        font: 'SimSun',
        fontAscii: 'Garamond',
        csFont: 'Amiri',
      }),
    ).toMatchObject({
      color: 'C00000',
      sizeHalfPoints: 21,
      font: 'SimSun',
      fontAscii: 'Garamond',
      csFont: 'Amiri',
    })
  })
})
