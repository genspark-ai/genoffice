import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { collectRevisions, type RevisionRange } from '../src/renderer/editor/revisions'
import {
  DEFAULT_REVISION_FILTER,
  authorInkIndex,
  buildReviewEntries,
  filterRevisions,
  revisionAuthorCss,
  revisionAuthors,
  revisionExcerpt,
  revisionsInSelection,
  revisionShown,
} from '../src/renderer/editor/revision-filters'
import {
  handleRevision,
  handleRevisionRange,
  type ReviewContext,
} from '../src/renderer/review-actions'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const text = (t: string, marks?: JsonNode['marks']): JsonNode => ({
  type: 'text',
  text: t,
  ...(marks ? { marks } : {}),
})
const mark = (type: string, author: string) => ({
  type,
  attrs: { author, date: '2026-07-01T10:00:00Z' },
})
const para = (...content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
})

function createEditor(content: JsonNode[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

const rev = (kind: RevisionRange['kind'], author: string, from = 0, to = 1): RevisionRange => ({
  from,
  to,
  kind,
  author,
})

describe('revisionShown / filterRevisions', () => {
  const ranges = [
    rev('ins', 'Ann', 1, 4),
    rev('del', 'Bob', 5, 8),
    rev('rPrChange', 'Ann', 9, 12),
    rev('pPrChange', 'Cy', 13, 20),
    rev('rowIns', 'Bob', 21, 30),
  ]

  it('shows everything by default', () => {
    expect(filterRevisions(ranges, DEFAULT_REVISION_FILTER)).toHaveLength(5)
  })

  it('drops format changes when Formatting is unticked', () => {
    const shown = filterRevisions(ranges, { insDel: true, format: false, hiddenAuthors: [] })
    expect(shown.map((r) => r.kind)).toEqual(['ins', 'del', 'rowIns'])
  })

  it('drops insertions, deletions and table revisions when Insertions and Deletions is unticked', () => {
    const shown = filterRevisions(ranges, { insDel: false, format: true, hiddenAuthors: [] })
    expect(shown.map((r) => r.kind)).toEqual(['rPrChange', 'pPrChange'])
  })

  it('hides an unticked reviewer regardless of kind', () => {
    const filter = { insDel: true, format: true, hiddenAuthors: ['Ann'] }
    expect(filterRevisions(ranges, filter).map((r) => r.author)).toEqual(['Bob', 'Cy', 'Bob'])
    expect(revisionShown(filter, rev('ins', 'Ann'))).toBe(false)
    expect(revisionShown(filter, rev('ins', 'Bob'))).toBe(true)
  })
})

describe('revisionsInSelection', () => {
  const ranges = [rev('ins', 'Ann', 1, 4), rev('del', 'Bob', 6, 9), rev('ins', 'Ann', 12, 15)]

  it('returns the ranges intersecting a non-empty selection', () => {
    expect(revisionsInSelection(ranges, 3, 7)).toEqual([ranges[0], ranges[1]])
    expect(revisionsInSelection(ranges, 4, 6)).toEqual([])
  })

  it('treats a caret as "inside" when it sits within or at the edge of a range', () => {
    expect(revisionsInSelection(ranges, 9, 9)).toEqual([ranges[1]])
    expect(revisionsInSelection(ranges, 10, 10)).toEqual([])
  })
})

describe('author inks', () => {
  it('numbers authors by first appearance and cycles after eight', () => {
    const ranges = ['Ann', 'Bob', 'Ann', 'Cy'].map((a) => rev('ins', a))
    const authors = revisionAuthors(ranges)
    expect(authors).toEqual(['Ann', 'Bob', 'Cy'])
    expect(authorInkIndex(authors, 'Ann')).toBe(0)
    expect(authorInkIndex(authors, 'Cy')).toBe(2)
    expect(authorInkIndex(authors, 'nobody')).toBe(0)
    const many = Array.from({ length: 10 }, (_, i) => `u${i}`)
    expect(authorInkIndex(many, 'u8')).toBe(0)
    expect(authorInkIndex(many, 'u9')).toBe(1)
  })

  it('emits no rule for the first author and escapes names in the others', () => {
    expect(revisionAuthorCss(['Ann'], [])).toBe('')
    const css = revisionAuthorCss(['Ann', 'O"Brien'], [])
    expect(css).toContain('[data-ins-author="O\\"Brien"]')
    expect(css).toContain('--docs-paper-rev-ink:var(--docs-paper-rev-ink-1)')
    expect(css).not.toContain('"Ann"')
  })

  it('hides unticked reviewers as accepted, paragraph chrome included', () => {
    const css = revisionAuthorCss(['Ann'], ['Ann'])
    expect(css).toContain('.doc-del[data-del-author="Ann"]{display:none}')
    expect(css).toContain(
      '.doc-ins[data-ins-author="Ann"]{color:inherit;text-decoration:none;font-weight:inherit;font-style:inherit}',
    )
    expect(css).toContain('.has-move-from[data-rev-author="Ann"]{display:none}')
    expect(css).toContain('.has-ppr-change[data-ppr-change-author="Ann"]::after{content:none}')
  })
})

describe('buildReviewEntries', () => {
  it('lists every revision in document order with an excerpt and ink', () => {
    const long = 'x'.repeat(80)
    const editor = createEditor([
      para(text('keep '), text('added', [mark('ins', 'Ann')]), text(' end')),
      para(text(long, [mark('del', 'Bob')])),
    ])
    const ranges = collectRevisions(editor.state.doc)
    const entries = buildReviewEntries(editor.state.doc, ranges)
    expect(entries.map((e) => [e.rev.kind, e.excerpt.length, e.inkIndex])).toEqual([
      ['ins', 5, 0],
      ['del', 60, 1],
    ])
    expect(entries[0].excerpt).toBe('added')
    expect(entries[1].excerpt.endsWith('…')).toBe(true)
    expect(revisionExcerpt(editor.state.doc, { from: 0, to: 0 })).toBe('')
    editor.destroy()
  })
})

describe('handleRevision scopes', () => {
  const ctxFor = (editor: Editor, filter = DEFAULT_REVISION_FILTER) => {
    const shown: RevisionRange[] = []
    const status: string[] = []
    const ctx = {
      editor,
      dirtyRef: { current: false },
      setStatus: (s: string) => status.push(s),
      revisionFilter: filter,
      showRevision: (r: RevisionRange) => shown.push(r),
      storyRevisions: () => [],
      currentStoryRevision: () => null,
      applyStoryRevisions: () => 0,
    } as unknown as ReviewContext
    return { ctx, shown, status }
  }
  const threeChanges = () => [
    para(text('a '), text('one', [mark('ins', 'Ann')]), text(' b')),
    para(text('c '), text('two', [mark('del', 'Bob')]), text(' d')),
    para(text('e '), text('three', [mark('ins', 'Ann')]), text(' f')),
  ]

  it('"selection" only touches changes intersecting the selection', () => {
    const editor = createEditor(threeChanges())
    const ranges = collectRevisions(editor.state.doc)
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, ranges[1].from, ranges[1].to),
      ),
    )
    const { ctx } = ctxFor(editor)
    handleRevision(ctx, 'accept', 'selection')
    const left = collectRevisions(editor.state.doc)
    expect(left.map((r) => r.kind)).toEqual(['ins', 'ins'])
    expect(editor.state.doc.textContent).not.toContain('two')
    expect(ctx.dirtyRef.current).toBe(true)
    editor.destroy()
  })

  it('"shown" respects the Markup filter and reports the count', () => {
    const editor = createEditor(threeChanges())
    const { ctx, status } = ctxFor(editor, { insDel: true, format: true, hiddenAuthors: ['Bob'] })
    handleRevision(ctx, 'reject', 'shown')
    expect(collectRevisions(editor.state.doc).map((r) => r.author)).toEqual(['Bob'])
    expect(status[0]).toMatch(/2/)
    editor.destroy()
  })

  it('an author scope leaves the other reviewers alone', () => {
    const editor = createEditor(threeChanges())
    handleRevision(ctxFor(editor).ctx, 'accept', { author: 'Ann' })
    expect(collectRevisions(editor.state.doc).map((r) => r.author)).toEqual(['Bob'])
    editor.destroy()
  })

  it('"next" handles the change at the caret and moves on to the following shown one', () => {
    const editor = createEditor(threeChanges())
    const ranges = collectRevisions(editor.state.doc)
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, ranges[0].from + 1)),
    )
    const { ctx, shown } = ctxFor(editor, { insDel: true, format: true, hiddenAuthors: ['Bob'] })
    handleRevision(ctx, 'accept', 'next')
    expect(collectRevisions(editor.state.doc)).toHaveLength(2)
    expect(shown).toHaveLength(1)
    expect(shown[0].author).toBe('Ann')
    expect(editor.state.doc.textBetween(shown[0].from, shown[0].to)).toBe('three')
    editor.destroy()
  })

  it('a single range is re-resolved against the live document before it is applied', () => {
    const editor = createEditor(threeChanges())
    const snapshot = collectRevisions(editor.state.doc)
    const { ctx, status } = ctxFor(editor)
    // the pane's snapshot is stale by the time the user clicks: text was typed before the first change
    editor.view.dispatch(editor.state.tr.insertText('xy', 1))
    handleRevisionRange(ctx, 'reject', snapshot[2])
    expect(editor.state.doc.textContent).not.toContain('three')
    expect(editor.state.doc.textContent).toContain('one')
    expect(editor.state.doc.textContent).toContain('two')
    // a change that no longer sits there is reported, not applied to whatever occupies the range now
    handleRevisionRange(ctx, 'reject', snapshot[2])
    expect(status).toHaveLength(1)
    expect(collectRevisions(editor.state.doc)).toHaveLength(2)
    editor.destroy()
  })
})
