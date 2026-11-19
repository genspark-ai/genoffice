import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  COVER_END_MARK,
  COVER_PRESETS,
  COVER_START_MARK,
  buildCoverNodes,
  coverPageRange,
  insertCoverPage,
  removeCoverPage,
} from '../src/renderer/editor/cover-pages'

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
    const band = nodes.find((n) => n.content?.some((c) => c.text === '\u6587\u6863\u6807\u9898'))!
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

describe('removeCoverPage', () => {
  const texts = (editor: Editor) =>
    (editor.getJSON().content ?? []).map(
      (n) => n.content?.map((c) => ('text' in c ? c.text : '')).join('') ?? '',
    )

  it('is unavailable until a gallery cover exists', () => {
    const editor = createEditor()
    expect(coverPageRange(editor.state.doc)).toBeNull()
    expect(removeCoverPage(editor)).toBe(false)
    expect(texts(editor)).toEqual(['First body paragraph'])
    editor.destroy()
  })

  it('marks the first cover paragraph and the trailing break with hidden bookmarks', () => {
    const nodes = buildCoverNodes(COVER_PRESETS[0])
    expect(nodes[0].attrs?.hiddenBookmarks).toEqual([COVER_START_MARK])
    expect(nodes[nodes.length - 1].attrs?.hiddenBookmarks).toEqual([COVER_END_MARK])
  })

  it('removes exactly the inserted cover and its page break, keeping the body', () => {
    const editor = createEditor()
    insertCoverPage(editor, COVER_PRESETS[0])
    expect(coverPageRange(editor.state.doc)).toEqual({
      from: 0,
      to: editor.state.doc.nodeSize - 2 - editor.state.doc.lastChild!.nodeSize,
    })
    expect(removeCoverPage(editor)).toBe(true)
    expect(texts(editor)).toEqual(['First body paragraph'])
    expect(coverPageRange(editor.state.doc)).toBeNull()
    expect(editor.state.doc.firstChild!.attrs.pageBreakBefore).toBe(false)
    editor.destroy()
  })

  it('survives edits inside the cover: the range follows the markers, not the preset shape', () => {
    const editor = createEditor()
    insertCoverPage(editor, COVER_PRESETS[0])
    const titlePos = editor.state.doc.resolve(editor.state.doc.child(0).nodeSize + 1).end()
    editor.chain().setTextSelection(titlePos).insertContent(' edited').splitBlock().run()
    expect(removeCoverPage(editor)).toBe(true)
    expect(texts(editor)).toEqual(['First body paragraph'])
    editor.destroy()
  })

  it('inserting another cover replaces the current one instead of stacking', () => {
    const editor = createEditor()
    insertCoverPage(editor, COVER_PRESETS[0])
    insertCoverPage(editor, COVER_PRESETS[1])
    const all = texts(editor)
    expect(all.filter((t) => t.includes('\u6587\u6863\u6807\u9898'))).toHaveLength(1)
    expect(all[all.length - 1]).toBe('First body paragraph')
    expect(removeCoverPage(editor)).toBe(true)
    expect(texts(editor)).toEqual(['First body paragraph'])
    editor.destroy()
  })
})
