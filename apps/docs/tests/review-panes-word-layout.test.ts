import { describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { CommentInfo } from '@genoffice/docx-engine'
import { LocaleProvider, setModuleLang } from '../src/renderer/i18n/locale'
import { ReviewingPane } from '../src/renderer/components/ReviewingPane'
import { loadLeftPane, saveLeftPane } from '../src/renderer/components/pane-width'
import { commentsView, iconClickTarget } from '../src/renderer/components/comments-mode'
import { summarizeRevisions, type ReviewEntry } from '../src/renderer/editor/revision-filters'
import { syncMarginAnnotations } from '../src/renderer/editor/margin-annotations'
import type { RevisionRange } from '../src/renderer/editor/revisions'
import { balloonKindOf } from '../src/renderer/editor/revision-view'

Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })
setModuleLang('en')
Element.prototype.scrollIntoView = () => {}

const rev = (kind: RevisionRange['kind'], from: number, extra: Partial<RevisionRange> = {}) =>
  ({
    kind,
    from,
    to: from + 2,
    author: 'Ann',
    date: '2026-09-01T10:00:00Z',
    ...extra,
  }) as RevisionRange

/** ins, del, one move pair, a run format change and an inserted cell */
const MIXED: RevisionRange[] = [
  rev('ins', 1),
  rev('del', 5),
  rev('moveFrom', 9, { moveId: 'm1' }),
  rev('moveTo', 13, { moveId: 'm1' }),
  rev('rPrChange', 17),
  rev('cellIns', 21, { author: 'Ben' }),
]

describe('Reviewing Pane summary (Word lines)', () => {
  it('counts insertions, deletions, moves (a pair once) and formatting', () => {
    expect(summarizeRevisions(MIXED)).toEqual({
      insertions: 2,
      deletions: 1,
      moves: 1,
      formatting: 1,
    })
  })

  it('renders the five Word lines with the open-thread count', () => {
    const entries: ReviewEntry[] = MIXED.map((r, index) => ({
      rev: r,
      index,
      excerpt: 'text',
      inkIndex: r.author === 'Ann' ? 0 : 1,
    }))
    const comments = [
      { id: '0', author: 'Ann', text: 'open', done: false },
      { id: '1', author: 'Ann', text: 'closed', done: true },
    ] as CommentInfo[]
    const container = document.createElement('div')
    document.body.appendChild(container)
    const root = createRoot(container)
    const pane = createElement(ReviewingPane, {
      entries,
      hiddenCount: 0,
      authors: [
        { name: 'Ann', inkIndex: 0 },
        { name: 'Ben', inkIndex: 1 },
      ],
      openComments: comments.filter((c) => !c.parentId && !c.done).length,
      activeIndex: -1,
      canEdit: true,
      onSelect: () => {},
      onAccept: () => {},
      onReject: () => {},
      onClose: () => {},
    })
    act(() => root.render(createElement(LocaleProvider, { initial: 'en', children: pane })))

    const lines = [...container.querySelectorAll('.review-pane-sum-line')].map((el) =>
      el.textContent!.replace(/\s+/g, ' ').trim(),
    )
    expect(lines).toEqual(['Insertions:2', 'Deletions:1', 'Moves:1', 'Formatting:1', 'Comments:1'])
    expect(container.querySelector('.review-pane-count')!.textContent).toBe('5')
    // flat rows: kind + author on one line, no per-kind tally sentence
    const heads = [...container.querySelectorAll('.review-row-head')].map((el) => el.textContent)
    expect(heads[0]).toContain('Inserted')
    expect(heads[0]).toContain('Ann')
    expect(container.textContent).not.toContain('changes')
    expect(container.querySelectorAll('.review-row .review-row-act')).toHaveLength(12)
    act(() => root.unmount())
    container.remove()
  })
})

describe('margin comment cards vs the Comments panel', () => {
  type Rect = { top: number; bottom: number; left: number; right: number; height: number }
  const rect = (top: number): Rect => ({ top, bottom: top + 16, left: 10, right: 50, height: 16 })
  const build = (flags: { cards?: boolean; focus?: string; popup?: string }) => {
    const app = document.createElement('div')
    app.className = 'app'
    if (flags.cards) app.setAttribute('data-comment-cards', '')
    if (flags.focus) app.setAttribute('data-comment-focus', flags.focus)
    if (flags.popup) app.setAttribute('data-comment-popup', flags.popup)
    const wrap = document.createElement('div')
    wrap.className = 'page-wrap'
    const pm = document.createElement('div')
    pm.className = 'ProseMirror'
    const span = document.createElement('span')
    span.className = 'doc-comment'
    span.dataset.commentIds = '7'
    span.getClientRects = () => [rect(100)] as unknown as DOMRectList
    pm.appendChild(span)
    wrap.appendChild(pm)
    app.appendChild(wrap)
    document.body.appendChild(app)
    const comments = [
      { id: '7', author: 'Ann Lee', text: 'hello', date: '2026-09-01T10:00:00Z' },
      { id: '8', author: 'Ben', text: 'reply', parentId: '7' },
    ] as CommentInfo[]
    const calls: string[] = []
    const handlers = {
      onCommentReply: (id: string, text: string) => calls.push(`reply:${id}:${text}`),
      onCommentResolve: (id: string, done: boolean) => calls.push(`resolve:${id}:${done}`),
      onCommentEdit: (id: string, text: string) => calls.push(`edit:${id}:${text}`),
      onCommentDelete: (id: string) => calls.push(`delete:${id}`),
    }
    syncMarginAnnotations(wrap, pm, comments, 1, [], undefined, handlers)
    return { app, wrap, pm, comments, handlers, calls }
  }

  it('draws a Word-style card (avatar, author, reply, Reply box) in Contextual mode', () => {
    const { app, wrap } = build({ cards: true })
    const bubble = wrap.querySelector('.comment-bubble')!
    expect(bubble).not.toBeNull()
    expect(bubble.querySelector('.comment-avatar')!.textContent).toBe('AL')
    expect(bubble.querySelector('.comment-author')!.textContent).toBe('Ann Lee')
    expect(bubble.querySelector('.comment-reply .comment-author')!.textContent).toBe('Ben')
    expect(bubble.querySelector('input.comment-reply-input')).not.toBeNull()
    expect(wrap.classList.contains('has-markup-area')).toBe(true)
    // Word keeps the anchor glyph in the page margin beside the card
    expect(wrap.querySelectorAll('.comment-balloon')).toHaveLength(1)
    // comments only: a comment column, no revision markup strip
    expect(wrap.querySelector('.markup-strip')).toBeNull()
    expect(wrap.style.getPropertyValue('--markup-w')).toBe('260px')
    app.remove()
  })

  it('card head carries Resolve and ⋯ → Edit / Delete, like the panel card', () => {
    const { app, wrap, calls } = build({ cards: true, focus: '7' })
    const bubble = wrap.querySelector<HTMLElement>('.comment-bubble')!
    const like = bubble.querySelector<HTMLButtonElement>('.comment-entry .comment-like')!
    like.click()
    expect(like.classList.contains('on')).toBe(true)
    expect(calls).toEqual([])
    like.click()
    expect(like.classList.contains('on')).toBe(false)
    expect(bubble.querySelector('.comment-reply .comment-like')).toBeNull()
    bubble.querySelector<HTMLButtonElement>('.comment-entry .comment-resolve')!.click()
    expect(calls).toEqual(['resolve:7:true'])
    expect(bubble.querySelector('.comment-reply .comment-resolve')).toBeNull()
    const more = bubble.querySelector<HTMLButtonElement>('.comment-entry .comment-more')!
    more.click()
    const menu = bubble.querySelector<HTMLElement>('.comment-entry .comment-menu')!
    expect(menu.hidden).toBe(false)
    expect(menu.querySelector('.comment-card-del')!.textContent).toBe(
      'Delete comment (and replies)',
    )
    menu.querySelector<HTMLButtonElement>('.comment-card-edit')!.click()
    const ta = bubble.querySelector<HTMLTextAreaElement>('.comment-entry textarea')!
    expect(ta.value).toBe('hello')
    ta.value = 'hello there'
    bubble.querySelector<HTMLButtonElement>('.comment-compose-actions .primary')!.click()
    expect(calls).toContain('edit:7:hello there')
    // Resolve commits an edit still open in the card before marking the thread done
    more.click()
    bubble.querySelector<HTMLButtonElement>('.comment-entry .comment-card-edit')!.click()
    bubble.querySelector<HTMLTextAreaElement>('.comment-entry textarea')!.value = 'final words'
    bubble.querySelector<HTMLButtonElement>('.comment-entry .comment-resolve')!.click()
    expect(calls.slice(-2)).toEqual(['edit:7:final words', 'resolve:7:true'])
    const rm = bubble.querySelector<HTMLButtonElement>('.comment-reply .comment-more')!
    rm.click()
    bubble.querySelector<HTMLButtonElement>('.comment-reply .comment-card-del')!.click()
    expect(calls).toContain('delete:8')
    app.remove()
  })

  it('Reply posts on Enter and a half-typed reply survives the rebuild that selecting triggers', () => {
    const { app, wrap, pm, comments, handlers, calls } = build({ cards: true, focus: '7' })
    const input = wrap.querySelector<HTMLInputElement>('.comment-reply-input')!
    input.focus()
    input.value = 'half'
    syncMarginAnnotations(wrap, pm, comments, 1, [], undefined, handlers)
    const again = wrap.querySelector<HTMLInputElement>('.comment-reply-input')!
    expect(again).not.toBe(input)
    expect(again.value).toBe('half')
    expect(document.activeElement).toBe(again)
    again.value = 'half done'
    again.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    expect(calls).toEqual(['reply:7:half done'])
    expect(again.value).toBe('')
    app.remove()
  })

  it('keeps the anchor icon but no card while Show Comments is off or the panel hosts the threads', () => {
    const { app, wrap } = build({})
    expect(wrap.querySelectorAll('.comment-bubble')).toHaveLength(0)
    expect(wrap.querySelectorAll('.comment-balloon')).toHaveLength(1)
    expect(wrap.classList.contains('has-markup-area')).toBe(false)
    app.remove()
  })

  it('fills the selected icon and outlines the selected card', () => {
    const { app, wrap } = build({ cards: true, focus: '7' })
    expect(wrap.querySelector('.comment-balloon')!.classList.contains('on')).toBe(true)
    expect(wrap.querySelector('.comment-bubble')!.classList.contains('active')).toBe(true)
    app.remove()
  })

  it('pops one thread up beside the page without reserving the markup column', () => {
    const { app, wrap } = build({ popup: '7' })
    const bubble = wrap.querySelector('.comment-bubble')!
    expect(bubble.classList.contains('comment-popup')).toBe(true)
    expect(wrap.querySelector('.comment-balloon')!.classList.contains('on')).toBe(true)
    expect(wrap.classList.contains('has-markup-area')).toBe(false)
    app.remove()
  })
})

describe('Navigation pane persistence', () => {
  it('persists the open pane and migrates the legacy showNav flag', () => {
    localStorage.clear()
    localStorage.setItem('aidocs.showNav', '1')
    expect(loadLeftPane()).toBe('nav')
    saveLeftPane(null)
    expect(loadLeftPane()).toBe(null)
    saveLeftPane('nav')
    expect(loadLeftPane()).toBe('nav')
  })
})

describe('Show Comments: Contextual / List', () => {
  it('off shows neither cards nor panel; List is the panel; Contextual is cards where the margin can host them', () => {
    expect(commentsView(false, 'contextual', true, true)).toEqual({ cards: false, panel: false })
    expect(commentsView(true, 'list', true, false)).toEqual({ cards: false, panel: true })
    expect(commentsView(true, 'contextual', true, false)).toEqual({ cards: true, panel: false })
  })

  it('Contextual without a card host (Simple Markup, inline) shows nothing until the panel is requested', () => {
    expect(commentsView(true, 'contextual', false, false)).toEqual({ cards: false, panel: false })
    expect(commentsView(true, 'contextual', false, true)).toEqual({ cards: false, panel: true })
    expect(commentsView(true, 'contextual', true, true)).toEqual({ cards: true, panel: true })
  })

  it('anchor icon selects the hosted card or panel entry, else pops the thread up', () => {
    expect(iconClickTarget({ cards: true, panel: false })).toBe('focus')
    expect(iconClickTarget({ cards: false, panel: true })).toBe('focus')
    expect(iconClickTarget({ cards: false, panel: false })).toBe('popup')
  })
})

describe('balloonKindOf', () => {
  it('maps every property-change kind to a fmt balloon and deletions to del', () => {
    for (const k of [
      'pPrChange',
      'rPrChange',
      'paraMarkRPrChange',
      'tblPrChange',
      'tblGridChange',
      'trPrChange',
      'tblPrExChange',
      'tcPrChange',
      'sectPrChange',
      'numberingChange',
      'cellMerge',
    ])
      expect(balloonKindOf(k)).toBe('fmt')
    for (const k of ['del', 'paraMarkDel', 'moveFrom', 'rowDel', 'cellDel', 'blockDel'])
      expect(balloonKindOf(k)).toBe('del')
    for (const k of ['ins', 'paraMarkIns', 'moveTo', 'rowIns', 'cellIns', 'blockIns'])
      expect(balloonKindOf(k)).toBeNull()
  })
})
