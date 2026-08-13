/**
 * Speaker notes tests (Task 3):
 *  1. Read notes (pptx that has a notesSlide)
 *  2. Write new notes (blank deck, no notesSlide → part auto-created)
 *  3. Write to a real pptx fixture → save+reopen roundtrip
 *  4. Iron rule: notesSlide bytes of slides without written notes stay unchanged
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  createBlankPptx,
  getSlideNotes,
  openPptx,
  savePptx,
  setSlideNotes,
  insertBlankSlide,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

describe('speaker notes', () => {
  it('getSlideNotes returns empty string on a fresh deck without notes', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(getSlideNotes(opened.archive, opened.deck.slides[0]!.path)).toBe('')
  })

  it('getSlideNotes reads back immediately after setSlideNotes', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(setSlideNotes(opened, 0, 'first line\nsecond line')).toBe(true)
    expect(getSlideNotes(opened.archive, opened.deck.slides[0]!.path)).toBe(
      'first line\nsecond line',
    )
  })

  it('save → reopen persists notes (notesSlide part auto-created)', async () => {
    const opened = await openPptx(await createBlankPptx())
    setSlideNotes(opened, 0, 'Introduce yourself first & special chars <test>')
    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideNotes(reopened.archive, reopened.deck.slides[0]!.path)).toBe(
      'Introduce yourself first & special chars <test>',
    )
    // notesSlide part is registered in [Content_Types].xml
    expect(reopened.archive.readText('[Content_Types].xml')).toContain('notesSlide+xml')
    // notesMaster has been created
    expect(reopened.archive.has('ppt/notesMasters/notesMaster1.xml')).toBe(true)
  })

  it('real pptx fixture → write notes → save+reopen correct', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    setSlideNotes(opened, 0, 'Speaker notes for the first slide')
    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideNotes(reopened.archive, reopened.deck.slides[0]!.path)).toBe(
      'Speaker notes for the first slide',
    )
    // Other slides' notes unmodified (no notesSlide, or original note bytes unchanged)
    if (opened.deck.slides.length > 1) {
      const slide1Before = getSlideNotes(opened.archive, opened.deck.slides[1]!.path)
      const slide1After = getSlideNotes(reopened.archive, reopened.deck.slides[1]!.path)
      expect(slide1After).toBe(slide1Before)
    }
  })

  it('clearing notes then saving: reopen returns empty string', async () => {
    const opened = await openPptx(await createBlankPptx())
    setSlideNotes(opened, 0, 'temporary note')
    setSlideNotes(opened, 0, '')
    expect(getSlideNotes(opened.archive, opened.deck.slides[0]!.path)).toBe('')
    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideNotes(reopened.archive, reopened.deck.slides[0]!.path)).toBe('')
  })

  it('iron rule: slides without written notes gain no notesSlide, existing notesSlide bytes unchanged', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const deck = opened.deck
    const _slidePath = deck.slides[0]!.path

    // If the fixture already has notesSlides, record the original bytes
    const notesPathBefore: string[] = []
    for (const s of deck.slides) {
      for (const rel of opened.archive.readRels(s.path).values()) {
        if (rel.type.endsWith('/notesSlide')) {
          const p = `ppt/${rel.target.replace(/^\.\.\//, '')}`
          notesPathBefore.push(p)
        }
      }
    }

    // Write notes only on page 0; leave other slides untouched
    setSlideNotes(opened, 0, 'only slide 0 changed')
    const saved = await savePptx(opened)
    const reopened = await openPptx(saved)

    // Multi-slide pptx: from slide 1 on, slides without written notes keep notesSlide bytes unchanged (or still absent)
    if (deck.slides.length > 1) {
      for (let i = 1; i < deck.slides.length; i++) {
        const sl = deck.slides[i]!
        let origNotes: string | null = null
        for (const rel of opened.archive.readRels(sl.path).values()) {
          if (rel.type.endsWith('/notesSlide')) {
            origNotes = opened.archive.readText(`ppt/${rel.target.replace(/^\.\.\//, '')}`)
          }
        }
        let afterNotes: string | null = null
        for (const rel of reopened.archive.readRels(reopened.deck.slides[i]!.path).values()) {
          if (rel.type.endsWith('/notesSlide')) {
            afterNotes = reopened.archive.readText(`ppt/${rel.target.replace(/^\.\.\//, '')}`)
          }
        }
        expect(afterNotes).toBe(origNotes)
      }
    }

    // Slide 0 notes were written correctly
    expect(getSlideNotes(reopened.archive, reopened.deck.slides[0]!.path)).toBe(
      'only slide 0 changed',
    )
  })

  it('multiple slides: per-slide notes do not interfere', async () => {
    const opened = await openPptx(await createBlankPptx())
    // Insert a second slide
    insertBlankSlide(opened, 0)
    setSlideNotes(opened, 0, 'slide 0 notes')
    setSlideNotes(opened, 1, 'slide 1 notes')

    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideNotes(reopened.archive, reopened.deck.slides[0]!.path)).toBe('slide 0 notes')
    expect(getSlideNotes(reopened.archive, reopened.deck.slides[1]!.path)).toBe('slide 1 notes')
  })
})
