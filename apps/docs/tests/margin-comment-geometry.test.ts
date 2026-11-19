import { describe, expect, it } from 'vitest'
import type { CommentInfo } from '@genoffice/docx-engine'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { syncMarginAnnotations } from '../src/renderer/editor/margin-annotations'

const box = (left: number, top: number, width: number, height: number): DOMRect =>
  ({ left, top, right: left + width, bottom: top + height, width, height }) as DOMRect

function fixture() {
  const app = document.createElement('div')
  app.className = 'app'
  app.setAttribute('data-comment-cards', '')
  const wrap = document.createElement('div')
  wrap.className = 'page-wrap'
  const pm = document.createElement('div')
  pm.className = 'ProseMirror doc-page'
  app.append(wrap)
  wrap.append(pm)
  document.body.append(app)
  wrap.getBoundingClientRect = () => box(0, 0, 1460, 2400)
  pm.getBoundingClientRect = () => box(0, 0, 1200, 2400)
  const span = (id: string, rects: DOMRect[]) => {
    const el = document.createElement('span')
    el.className = 'doc-comment'
    el.dataset.commentIds = id
    el.getClientRects = () => rects as unknown as DOMRectList
    pm.append(el)
    return el
  }
  const comments: CommentInfo[] = [{ id: '1', author: 'A', text: 'Check this text' }]
  return { app, wrap, pm, span, comments }
}

describe('margin comment geometry', () => {
  it('connects to the last visible fragment of the full marked range', () => {
    const f = fixture()
    f.span('1', [box(200, 100, 80, 16)])
    f.span('1', [box(280, 100, 40, 16), box(100, 116, 70, 16)])
    syncMarginAnnotations(f.wrap, f.pm, f.comments, 1)
    expect(f.wrap.querySelector('.comment-leader')!.getAttribute('d')).toMatch(/^M 170 131 L /)
    expect(f.wrap.querySelector<HTMLElement>('.comment-bubble')!.style.top).toBe('116px')
    f.app.remove()
  })

  it('places cards beside their owning page rather than the widest canvas', () => {
    const f = fixture()
    f.span('1', [box(200, 100, 80, 16)])
    const gap = document.createElement('div')
    gap.className = 'page-gap'
    gap.style.setProperty('--gap-page-x', '0px')
    gap.style.setProperty('--gap-page-w', '1200px')
    gap.getBoundingClientRect = () => box(0, 1000, 1200, 40)
    f.pm.append(gap)
    f.span('2', [box(400, 1200, 80, 16)])
    f.comments.push({ id: '2', author: 'A', text: 'Wide page' })
    syncMarginAnnotations(f.wrap, f.pm, f.comments, 1, [], undefined, undefined, undefined, {
      left: 200,
      width: 800,
    })
    const cards = [...f.wrap.querySelectorAll<HTMLElement>('.comment-bubble')]
    expect(cards.map((c) => c.closest<HTMLElement>('.comment-page')!.style.left)).toEqual([
      '1014px',
      '1214px',
    ])
    const icons = [...f.wrap.querySelectorAll<HTMLElement>('.comment-balloon')]
    expect(icons.map((c) => c.style.left)).toEqual(['952px', '1152px'])
    f.app.remove()
  })

  it('keeps text coordinates correct when reserving the comment column moves the wrap', () => {
    const f = fixture()
    f.wrap.getBoundingClientRect = () =>
      box(f.wrap.classList.contains('has-markup-area') ? 0 : 100, 0, 1460, 2400)
    f.pm.getBoundingClientRect = () => box(100, 0, 1200, 2400)
    f.span('1', [box(200, 100, 80, 16)])
    syncMarginAnnotations(f.wrap, f.pm, f.comments, 1)
    expect(f.wrap.querySelector('.comment-leader')!.getAttribute('d')).toMatch(/^M 280 115 L /)
    f.app.remove()
  })

  it('anchors cross-paragraph ranges at their end marker rather than start block', () => {
    const f = fixture()
    const first = document.createElement('p')
    first.dataset.idx = '0'
    first.getClientRects = () => [box(200, 100, 300, 16)] as unknown as DOMRectList
    const last = document.createElement('p')
    last.dataset.idx = '1'
    last.getClientRects = () => [box(100, 1200, 400, 16)] as unknown as DOMRectList
    f.pm.append(first, last)
    syncMarginAnnotations(f.wrap, f.pm, f.comments, 1, [
      { docxIndex: 0, originalXml: '<w:p><w:commentRangeStart w:id="1"/><w:r/></w:p>' },
      { docxIndex: 1, originalXml: '<w:p><w:r/><w:commentRangeEnd w:id="1"/></w:p>' },
    ])
    expect(f.wrap.querySelector<HTMLElement>('.comment-bubble')!.style.top).toBe('1200px')
    f.app.remove()
  })

  it('keeps a cross-paragraph endpoint attached while text before and after it changes', () => {
    const f = fixture()
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0, commentStarts: ['1'] },
            content: [{ type: 'text', text: 'start' }],
          },
          {
            type: 'docParagraph',
            attrs: { docxIndex: 1, commentEnds: ['1'] },
            content: [{ type: 'text', text: 'target remaining' }],
          },
        ],
      },
    })
    f.pm.replaceWith(editor.view.dom)
    editor.view.dom.classList.add('doc-page')
    editor.view.dom.getBoundingClientRect = () => box(0, 0, 1200, 2400)
    editor.view.coordsAtPos = (pos) => ({
      left: pos * 10,
      right: pos * 10,
      top: 1200,
      bottom: 1216,
    })
    const blocks = [
      {
        docxIndex: 0,
        originalXml: '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>start</w:t></w:r></w:p>',
      },
      {
        docxIndex: 1,
        originalXml:
          '<w:p><w:r><w:t>target</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:t> remaining</w:t></w:r></w:p>',
      },
    ]
    const sync = () =>
      syncMarginAnnotations(f.wrap, editor.view.dom, f.comments, 1, blocks, editor.view)
    sync()
    expect(f.wrap.querySelector('.comment-leader')!.getAttribute('d')).toMatch(/^M 140 1215 L /)
    editor.view.dispatch(editor.state.tr.insertText('more ', 8))
    sync()
    expect(f.wrap.querySelector('.comment-leader')!.getAttribute('d')).toMatch(/^M 190 1215 L /)
    editor.view.dispatch(editor.state.tr.insertText('after ', 20))
    sync()
    expect(f.wrap.querySelector('.comment-leader')!.getAttribute('d')).toMatch(/^M 190 1215 L /)
    editor.destroy()
    f.app.remove()
  })
})
