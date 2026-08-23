/**
 * HTML clipboard round-trip for dual-slot fonts: renderHTML encodes the slots as a
 * font-family chain; parsing that chain back must restore the same font/fontAscii.
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { fontAttrsFromFamilyChain } from '../src/renderer/editor/marks'
import { cssDualFontFamily, cssFontFamily } from '../src/renderer/line-metrics'

function roundtrip(attrs: Record<string, unknown>): Record<string, unknown> | undefined {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: null },
          content: [
            { type: 'text', text: '合同 Contract', marks: [{ type: 'docTextStyle', attrs }] },
          ],
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

describe('dual-slot font HTML round-trip', () => {
  it('mixed run keeps both slots', () => {
    expect(roundtrip({ font: 'SimSun', fontAscii: 'Times New Roman' })).toMatchObject({
      font: 'SimSun',
      fontAscii: 'Times New Roman',
    })
  })

  it('CJK-only run keeps the primary slot and no Latin slot', () => {
    expect(roundtrip({ font: 'KaiTi' })).toMatchObject({ font: 'KaiTi', fontAscii: null })
  })

  it('Latin-only run keeps both slots equal (parse.ts convention)', () => {
    expect(roundtrip({ font: 'Arial', fontAscii: 'Arial' })).toMatchObject({
      font: 'Arial',
      fontAscii: 'Arial',
    })
  })
})

describe('fontAttrsFromFamilyChain', () => {
  it('inverts the encoded chains', () => {
    expect(fontAttrsFromFamilyChain(cssDualFontFamily('Times New Roman', 'SimSun'))).toEqual({
      font: 'SimSun',
      fontAscii: 'Times New Roman',
    })
    expect(fontAttrsFromFamilyChain(cssFontFamily('KaiTi'))).toEqual({ font: 'KaiTi' })
    expect(fontAttrsFromFamilyChain(cssFontFamily('Times New Roman'))).toEqual({
      font: 'Times New Roman',
      fontAscii: 'Times New Roman',
    })
  })

  it('never mistakes the bundled CJK tofu-fallback for a user font', () => {
    expect(fontAttrsFromFamilyChain("'Georgia','Noto Serif CJK SC',serif")).toEqual({
      font: 'Georgia',
      fontAscii: 'Georgia',
    })
    // genuinely picked as first family it still counts
    expect(fontAttrsFromFamilyChain("'Noto Sans CJK SC',sans-serif")).toEqual({
      font: 'Noto Sans CJK SC',
    })
  })

  it('skips internal GenOffice aliases even at the chain head', () => {
    expect(
      fontAttrsFromFamilyChain("'GenOffice Songti SC','STSong','SimSun','Noto Serif CJK SC',serif"),
    ).toEqual({ font: 'STSong' })
  })

  it('handles foreign single-family styles', () => {
    expect(fontAttrsFromFamilyChain('SimSun')).toEqual({ font: 'SimSun' })
    expect(fontAttrsFromFamilyChain('Calibri')).toEqual({ font: 'Calibri', fontAscii: 'Calibri' })
    expect(fontAttrsFromFamilyChain(undefined)).toEqual({})
  })
})

describe('complex-script (w:cs) render chain', () => {
  function editorHtml(text: string, attrs: Record<string, unknown>): string {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: null },
            content: [{ type: 'text', text, marks: [{ type: 'docTextStyle', attrs }] }],
          },
        ],
      },
    })
    const html = editor.getHTML()
    editor.destroy()
    return html
  }

  it('csFont leads the family chain, then the Latin chain', () => {
    const html = editorHtml('مرحبا Contract', {
      font: 'Calibri',
      fontAscii: 'Calibri',
      csFont: 'Arabic Typesetting',
    })
    expect(html).toContain(
      'font-family: &quot;Arabic Typesetting&quot;, &quot;Naskh Digits GO&quot;, &quot;Times Punct GO&quot;, &quot;Noto Naskh Arabic TNR&quot;, &quot;Calibri&quot;, &quot;Carlito GO&quot;, &quot;Noto Sans CJK SC&quot;, &quot;Times New Roman&quot;, &quot;Liberation Serif&quot;, &quot;Geeza Pro&quot;, &quot;Al Bayan&quot;, sans-serif',
    )
  })

  it('run marks get csFont only when the text is complex-script', () => {
    const runs = [
      { text: 'مرحبا', csFont: 'Amiri', font: 'Calibri', fontAscii: 'Calibri' },
      { text: 'latin only', csFont: 'Amiri', font: 'Calibri', fontAscii: 'Calibri' },
    ]
    const doc = blocksToPmDoc([{ type: 'paragraph', docxIndex: 0, runs } as never])
    const marks = doc.content![0].content!.map(
      (n) => n.marks!.find((m) => m.type === 'docTextStyle')!.attrs!,
    )
    expect(marks[0].csFont).toBe('Amiri')
    expect(marks[1].csFont).toBeNull()
  })
})
