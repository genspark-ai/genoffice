import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { COVER_PRESETS, buildCoverNodes, insertCoverPage } from '../src/renderer/editor/cover-pages'

function createEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: 'First body paragraph' }],
        },
      ],
    },
  })
}

describe('COVER_PRESETS', () => {
  it('provides 12 presets with unique ids and names', () => {
    expect(COVER_PRESETS).toHaveLength(12)
    expect(new Set(COVER_PRESETS.map((p) => p.id)).size).toBe(12)
    expect(new Set(COVER_PRESETS.map((p) => p.name)).size).toBe(12)
  })

  it('every preset contains the title placeholder', () => {
    for (const preset of COVER_PRESETS) {
      expect(preset.paras.some((p) => p.text === '文档标题')).toBe(true)
    }
  })
})

describe('buildCoverNodes', () => {
  it('ends with a page-break paragraph pushing content to page 2', () => {
    const nodes = buildCoverNodes(COVER_PRESETS[0])
    const last = nodes[nodes.length - 1]
    expect(last.type).toBe('docParagraph')
    expect(last.attrs?.pageBreakBefore).toBe(true)
  })

  it('maps styling onto paragraph attrs and text marks', () => {
    const preset = COVER_PRESETS.find((p) => p.id === 'banded')!
    const nodes = buildCoverNodes(preset)
    const band = nodes.find((n) => n.content?.some((c) => c.text === '文档标题'))!
    expect(band.attrs?.shadingFill).toBeTruthy()
    expect(band.attrs?.align).toBe('center')
    const marks = band.content![0].marks!
    expect(marks.some((m) => m.type === 'bold')).toBe(true)
    const style = marks.find((m) => m.type === 'docTextStyle')!
    expect(style.attrs?.color).toBe('FFFFFF')
    expect(style.attrs?.sizeHalfPoints).toBeGreaterThan(24)
  })
})

describe('insertCoverPage', () => {
  it('inserts at document start and keeps original content after it', () => {
    const editor = createEditor()
    insertCoverPage(editor, COVER_PRESETS[0])
    const json = editor.getJSON()
    const texts = (json.content ?? []).map(
      (n) => n.content?.map((c) => ('text' in c ? c.text : '')).join('') ?? '',
    )
    expect(texts[texts.length - 1]).toBe('First body paragraph')
    expect(texts.some((t) => t.includes('文档标题'))).toBe(true)
    expect(texts.indexOf('First body paragraph')).toBeGreaterThan(0)
    editor.destroy()
  })

  it('every preset round-trips through the editor schema without loss', () => {
    for (const preset of COVER_PRESETS) {
      const editor = createEditor()
      insertCoverPage(editor, preset)
      const all = JSON.stringify(editor.getJSON())
      expect(all).toContain('文档标题')
      editor.destroy()
    }
  })
})
