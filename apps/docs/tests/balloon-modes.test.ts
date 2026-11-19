/**
 * Show Markup → Balloons: which annotations the margin layer builds and which
 * inline widgets the editor shows in each of Word's three modes.
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { CommentInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { syncMarginAnnotations } from '../src/renderer/editor/margin-annotations'
import { revisionDisplayState } from '../src/renderer/editor/marks'
import { inlineCommentState, setInlineComments } from '../src/renderer/editor/balloon-markers'
import {
  balloonRevKinds,
  currentBalloonMode,
  loadBalloonMode,
  type BalloonMode,
} from '../src/renderer/editor/revision-view'

const del = { type: 'del', attrs: { author: 'Bob', date: '2026-07-02T10:00:00Z' } }
const rpr = {
  type: 'rprChange',
  attrs: { author: 'Bob', date: '2026-07-02T10:00:00Z', old: { bold: false } },
}
const comment = { type: 'comment', attrs: { ids: '7' } }
const comments: CommentInfo[] = [
  { id: '7', author: 'Mary Shaw', text: 'Check this', date: '2026-07-02T10:00:00Z' },
]

const content = {
  type: 'doc',
  content: [
    {
      type: 'docParagraph',
      attrs: { docxIndex: null },
      content: [
        { type: 'text', text: 'keep ' },
        { type: 'text', text: 'gone', marks: [del] },
        { type: 'text', text: ' bolded', marks: [{ type: 'bold' }, rpr] },
        { type: 'text', text: ' noted', marks: [comment] },
      ],
    },
  ],
}

const rect = { top: 100, bottom: 116, left: 50, right: 200, width: 150, height: 16 }
const wrapRect = { top: 0, bottom: 1000, left: 0, right: 800, width: 800, height: 1000 }

function mount(balloons: BalloonMode) {
  revisionDisplayState.mode = 'all'
  revisionDisplayState.balloons = balloons
  revisionDisplayState.balloonsActive = true
  setInlineComments(comments)
  const app = document.createElement('div')
  app.className = `app${balloons === 'revisions' ? ' rev-balloon' : ''}`
  app.dataset.balloons = balloons
  app.setAttribute('data-comment-cards', '')
  const wrap = document.createElement('div')
  wrap.className = 'page-wrap'
  const host = document.createElement('div')
  wrap.appendChild(host)
  app.appendChild(wrap)
  document.body.appendChild(app)
  const editor = new Editor({ element: host, extensions: editorExtensions, content })
  const pm = editor.view.dom as HTMLElement
  syncMarginAnnotations(wrap, pm, comments, 1, [], editor.view, {})
  return { app, wrap, pm, editor }
}

describe('Show Markup → Balloons', () => {
  const proto = Element.prototype
  const origRects = proto.getClientRects
  const origBox = proto.getBoundingClientRect
  beforeEach(() => {
    proto.getClientRects = function () {
      return [rect] as unknown as DOMRectList
    }
    proto.getBoundingClientRect = function () {
      const el = this as Element
      return (el.classList.contains('page-wrap') ? wrapRect : rect) as DOMRect
    }
  })
  afterEach(() => {
    proto.getClientRects = origRects
    proto.getBoundingClientRect = origBox
    document.body.innerHTML = ''
    revisionDisplayState.balloons = 'comments'
    revisionDisplayState.balloonsActive = false
    inlineCommentState.tags = new Map()
    localStorage.clear()
  })

  it('Show Revisions in Balloons: Deleted + Formatted + comment balloons, deletion marker inline', () => {
    const { wrap, pm, editor } = mount('revisions')
    expect(wrap.querySelectorAll('.rev-bubble-del')).toHaveLength(1)
    expect(wrap.querySelector('.rev-bubble-del .rev-bubble-text')?.textContent).toBe('gone')
    expect(wrap.querySelectorAll('.rev-bubble-fmt')).toHaveLength(1)
    expect(wrap.querySelectorAll('.comment-bubble:not(.rev-bubble)')).toHaveLength(1)
    expect(wrap.classList.contains('has-markup-area')).toBe(true)
    // Word: revision balloons in a markup strip on the page, the comment card in its own column right of it
    expect(wrap.querySelectorAll('.markup-strip').length).toBeGreaterThan(0)
    const strip = wrap.querySelector<HTMLElement>('.markup-strip')!
    const rev = wrap.querySelector<HTMLElement>('.rev-bubble-del')!
    const card = wrap.querySelector<HTMLElement>('.comment-bubble:not(.rev-bubble)')!
    expect(parseFloat(rev.style.left)).toBeGreaterThanOrEqual(parseFloat(strip.style.left))
    expect(parseFloat(card.style.left)).toBeGreaterThanOrEqual(
      parseFloat(strip.style.left) + parseFloat(strip.style.width),
    )
    expect(wrap.style.getPropertyValue('--markup-w')).toBe('460px')
    const marker = pm.querySelector<HTMLElement>('.doc-del-marker')
    expect(marker?.nextElementSibling?.classList.contains('doc-del')).toBe(true)
    expect(editor.view.posAtDOM(marker!, 0)).toBe(1 + 'keep '.length)
    expect(pm.querySelector('.doc-comment-marker')).toBeNull()
    editor.destroy()
  })

  it('Show All Revisions Inline: no balloons, no markup column, [initials N] comment tag', () => {
    const { wrap, pm, editor } = mount('inline')
    expect(wrap.querySelectorAll('.comment-bubble')).toHaveLength(0)
    expect(wrap.classList.contains('has-markup-area')).toBe(false)
    expect(wrap.style.getPropertyValue('--markup-w')).toBe('')
    const tag = pm.querySelector<HTMLElement>('.doc-comment-marker')
    expect(tag?.textContent).toBe('[MS1]')
    expect(tag?.title).toBe('Mary Shaw: Check this')
    expect(tag?.dataset.commentId).toBe('7')
    expect(pm.querySelector('.doc-del-marker')).toBeNull()
    // the deletion stays in the flow: its line gets a change bar like any inline revision
    expect(wrap.querySelectorAll('.change-bar').length).toBeGreaterThan(0)
    editor.destroy()
  })

  it('Show Only Comments and Formatting in Balloons (default): no Deleted balloon, deletion inline', () => {
    const { wrap, pm, editor } = mount('comments')
    expect(wrap.querySelectorAll('.rev-bubble-del')).toHaveLength(0)
    expect(wrap.querySelectorAll('.rev-bubble-fmt')).toHaveLength(1)
    expect(wrap.querySelectorAll('.comment-bubble:not(.rev-bubble)')).toHaveLength(1)
    expect(pm.querySelector('.doc-del-marker')).toBeNull()
    expect(pm.querySelector('.doc-comment-marker')).toBeNull()
    expect(pm.querySelector('.doc-del')?.textContent).toBe('gone')
    editor.destroy()
  })

  it('mode helpers: kinds per mode, root attribute, persisted default', () => {
    expect([...balloonRevKinds('revisions')]).toEqual(['del', 'fmt'])
    expect([...balloonRevKinds('comments')]).toEqual(['fmt'])
    expect([...balloonRevKinds('inline')]).toEqual([])
    expect(loadBalloonMode()).toBe('comments')
    localStorage.setItem('aidocs.balloonMode', 'inline')
    expect(loadBalloonMode()).toBe('inline')
    localStorage.setItem('aidocs.balloonMode', 'bogus')
    expect(loadBalloonMode()).toBe('comments')
    const root = document.createElement('div')
    expect(currentBalloonMode(root)).toBeNull()
    root.dataset.balloons = 'revisions'
    expect(currentBalloonMode(root.appendChild(document.createElement('p')))).toBe('revisions')
  })
})
