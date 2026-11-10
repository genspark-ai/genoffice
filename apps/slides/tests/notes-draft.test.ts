import { describe, expect, it } from 'vitest'
import {
  notesBaselineAfterFlush,
  notesDraftIndex,
  retargetNotesDraft,
  startNotesDraft,
} from '../src/shared/notes-draft'

const deck = (...parts: string[]) => parts.map((partPath) => ({ partPath }))
const draft = { index: 1, text: 'typing', base: 'old', partPath: 'ppt/slides/slide2.xml' }

describe('retargetNotesDraft (genoffice#900)', () => {
  it('keeps a draft when the edited deck left that slide and its notes alone', () => {
    const slides = deck('ppt/slides/slide1.xml', 'ppt/slides/slide2.xml')
    expect(retargetNotesDraft(draft, slides, 'old')).toBe(draft)
  })

  it('follows the slide when slides were inserted or reordered before it', () => {
    const slides = deck('ppt/slides/slide9.xml', 'ppt/slides/slide1.xml', 'ppt/slides/slide2.xml')
    expect(retargetNotesDraft(draft, slides, 'old')).toEqual({ ...draft, index: 2 })
  })

  it('drops the draft when the deck edit changed that slide’s notes', () => {
    const slides = deck('ppt/slides/slide1.xml', 'ppt/slides/slide2.xml')
    expect(retargetNotesDraft(draft, slides, 'rewritten by AI')).toBeNull()
  })

  it('drops the draft when the slide was removed', () => {
    expect(retargetNotesDraft(draft, deck('ppt/slides/slide1.xml'), 'old')).toBeNull()
  })

  it('survives an AI run when typing resumed after a flush that did not reload the notes', () => {
    const slides = deck('ppt/slides/slide1.xml', 'ppt/slides/slide2.xml')
    // blur/save/onBeforeRun flush the draft; the baseline must follow the written text
    const flushed = { index: 1, text: 'first half' }
    const resumed = startNotesDraft(null, flushed, 1, 'first half, second half', slides[1].partPath)
    expect(resumed.base).toBe('first half')
    expect(retargetNotesDraft(resumed, slides, 'first half')).toBe(resumed)
    // without the baseline update the flushed text would read as an AI rewrite
    const stale = startNotesDraft(
      null,
      { index: 1, text: 'old' },
      1,
      'first half, second half',
      'x',
    )
    expect(
      retargetNotesDraft({ ...stale, partPath: slides[1].partPath }, slides, 'first half'),
    ).toBeNull()
  })

  it('keeps the base across keystrokes and ignores a baseline for another slide', () => {
    const first = startNotesDraft(null, { index: 1, text: 'old' }, 1, 'o', 'p')
    expect(startNotesDraft(first, { index: 1, text: 'later' }, 1, 'ol', 'p').base).toBe('old')
    expect(startNotesDraft(null, { index: 0, text: 'other' }, 1, 'o', 'p').base).toBe('')
  })

  it('takes the read-back text as the baseline after a flush, so normalization cannot drop a draft', () => {
    const slides = deck('ppt/slides/slide1.xml', 'ppt/slides/slide2.xml')
    const loaded = notesBaselineAfterFlush({ index: 1, text: 'old' }, 1, 'typed')
    expect(loaded).toEqual({ index: 1, text: 'typed' })
    const resumed = startNotesDraft(null, loaded, 1, 'typed more', slides[1].partPath)
    expect(retargetNotesDraft(resumed, slides, 'typed')).toBe(resumed)
  })

  it('lets a late flush of another slide leave the current baseline alone', () => {
    const current = { index: 2, text: 'current' }
    expect(notesBaselineAfterFlush(current, 1, 'late')).toBe(current)
    expect(notesBaselineAfterFlush(null, 1, 'first')).toEqual({ index: 1, text: 'first' })
  })

  it('falls back to the index when slides carry no part path', () => {
    const anon = { index: 1, text: 't', base: '' }
    expect(notesDraftIndex(anon, [{}, {}])).toBe(1)
    expect(notesDraftIndex(anon, [{}])).toBe(-1)
    expect(retargetNotesDraft(anon, [{}, {}], '')).toBe(anon)
  })
})
