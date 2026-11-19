/**
 * Section read/write tests: round-trip, add/remove/rename/move.
 * Fixture 01_standard_business.pptx has 5 slides and no sections.
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  addSection,
  getSections,
  moveSection,
  moveSlide,
  normalizeSections,
  openPptx,
  removeSection,
  renameSection,
  savePptx,
  setSections,
  type OpenedPptx,
  type SectionInfo,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const open = () => openPptx(fx('01_standard_business.pptx'))
const presXml = (o: OpenedPptx) => o.archive.readText('ppt/presentation.xml')!

describe('sections (p14:sectionLst)', () => {
  it('getSections returns [] for legacy documents without sections', async () => {
    const opened = await open()
    expect(getSections(opened)).toEqual([])
  })

  it('setSections → savePptx → openPptx round-trip is consistent and does not touch Content_Types/rels', async () => {
    const opened = await open()
    const ctBefore = opened.archive.readText('[Content_Types].xml')
    const relsBefore = opened.archive.readText('ppt/_rels/presentation.xml.rels')
    const secs = [
      {
        id: '{11111111-2222-3333-4444-555555555555}',
        name: 'Opening & "Intro"',
        slideIndices: [0, 1],
      },
      { id: '{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}', name: 'Body', slideIndices: [2, 3, 4] },
    ]
    setSections(opened, secs)

    const xml = presXml(opened)
    expect(xml).toContain('<p14:sectionLst')
    expect(xml).toContain('{521415D9-36F7-43E2-AB2F-B90AF26B5E84}')
    // Special characters escaped as attribute values
    expect(xml).toContain('name="Opening &amp; &quot;Intro&quot;"')
    expect(opened.archive.readText('[Content_Types].xml')).toBe(ctBefore)
    expect(opened.archive.readText('ppt/_rels/presentation.xml.rels')).toBe(relsBefore)

    const reopened = await openPptx(await savePptx(opened))
    expect(getSections(reopened)).toEqual(secs)
  })

  it('p14:sldId references the numeric slide id (not rId)', async () => {
    const opened = await open()
    setSections(opened, [
      { id: '{11111111-2222-3333-4444-555555555555}', name: 'A', slideIndices: [0] },
    ])
    // The fixture's first slide has sldId 256
    expect(presXml(opened)).toContain('<p14:sldId id="256"/>')
  })

  it('addSection: sectionless document auto-gains a default section; then split an existing section', async () => {
    const opened = await open()
    const r1 = addSection(opened, 2, 'New Section')!
    expect(r1.map((s) => ({ name: s.name, slideIndices: s.slideIndices }))).toEqual([
      { name: 'Default Section', slideIndices: [0, 1] },
      { name: 'New Section', slideIndices: [2, 3, 4] },
    ])
    // Split off another section before slide 4 (index 3)
    const r2 = addSection(opened, 3, 'Split Again')!
    expect(r2.map((s) => ({ name: s.name, slideIndices: s.slideIndices }))).toEqual([
      { name: 'Default Section', slideIndices: [0, 1] },
      { name: 'New Section', slideIndices: [2] },
      { name: 'Split Again', slideIndices: [3, 4] },
    ])
    // Out-of-range index rejected
    expect(addSection(opened, 5, 'x')).toBeNull()
  })

  it('renameSection: rename takes effect and round-trips', async () => {
    const opened = await open()
    const [a] = addSection(opened, 0, 'Old Name')!
    const r = renameSection(opened, a!.id, 'New Name')!
    expect(r[0]!.name).toBe('New Name')
    expect(renameSection(opened, '{00000000-0000-0000-0000-000000000000}', 'x')).toBeNull()
    const reopened = await openPptx(await savePptx(opened))
    expect(getSections(reopened)[0]!.name).toBe('New Name')
  })

  it('removeSection(keepSlides): slides merge into the adjacent section; sectionLst removed when all sections are gone', async () => {
    const opened = await open()
    addSection(opened, 2, 'Second Half')
    const secs = getSections(opened) // [default section 0-1, second half 2-4]
    const r1 = removeSection(opened, secs[1]!.id, { keepSlides: true })!
    expect(r1).toHaveLength(1)
    expect(r1[0]!.slideIndices).toEqual([0, 1, 2, 3, 4])
    // Delete the first (only) section → no sections
    const r2 = removeSection(opened, r1[0]!.id, { keepSlides: true })!
    expect(r2).toEqual([])
    expect(presXml(opened)).not.toContain('p14:sectionLst')
    // All slides retained
    expect(opened.deck.slides).toHaveLength(5)
  })

  it('removeSection: removing the first section merges its slides into the next one', async () => {
    const opened = await open()
    addSection(opened, 2, 'Second Half')
    const secs = getSections(opened)
    const r = removeSection(opened, secs[0]!.id, { keepSlides: true })!
    expect(r).toHaveLength(1)
    expect(r[0]!.name).toBe('Second Half')
    expect(r[0]!.slideIndices).toEqual([0, 1, 2, 3, 4])
  })

  it('moveSection: whole section moves up, sldIdLst / deck.slides / section indices stay in sync, round-trip holds', async () => {
    const opened = await open()
    const pathsBefore = opened.deck.slides.map((s) => s.path)
    addSection(opened, 2, 'Second Half') // [default section 0-1, second half 2-4]
    const secs = getSections(opened)
    const r = moveSection(opened, secs[1]!.id, 'up')!
    expect(r.map((s) => ({ name: s.name, slideIndices: s.slideIndices }))).toEqual([
      { name: 'Second Half', slideIndices: [0, 1, 2] },
      { name: 'Default Section', slideIndices: [3, 4] },
    ])
    // deck.slides order: original slides 2,3,4 moved to the front
    expect(opened.deck.slides.map((s) => s.path)).toEqual([
      pathsBefore[2],
      pathsBefore[3],
      pathsBefore[4],
      pathsBefore[0],
      pathsBefore[1],
    ])
    // Edge cases: moving first section up / last section down rejected
    expect(moveSection(opened, r[0]!.id, 'up')).toBeNull()
    expect(moveSection(opened, r[1]!.id, 'down')).toBeNull()

    // Round-trip: slide order and sections consistent after save and reopen
    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.slides.map((s) => s.path)).toEqual([
      pathsBefore[2],
      pathsBefore[3],
      pathsBefore[4],
      pathsBefore[0],
      pathsBefore[1],
    ])
    expect(
      getSections(reopened).map((s) => ({ name: s.name, slideIndices: s.slideIndices })),
    ).toEqual([
      { name: 'Second Half', slideIndices: [0, 1, 2] },
      { name: 'Default Section', slideIndices: [3, 4] },
    ])
  })

  it('moveSlide: moving a slide down in a sectionless document keeps sldIdLst / deck.slides in sync and round-trips', async () => {
    const opened = await open()
    const p = opened.deck.slides.map((s) => s.path)
    expect(moveSlide(opened, 0, 3)).toBe(true)
    expect(opened.deck.slides.map((s) => s.path)).toEqual([p[1], p[2], p[3], p[0], p[4]])
    // A sectionless document should not gain sections just from moving a slide
    expect(presXml(opened)).not.toContain('p14:sectionLst')

    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.slides.map((s) => s.path)).toEqual([p[1], p[2], p[3], p[0], p[4]])

    // Edge cases: same position / out of range rejected
    expect(moveSlide(opened, 2, 2)).toBe(false)
    expect(moveSlide(opened, 9, 0)).toBe(false)
  })

  it('moveSlide: moving across sections assigns the slide to the target section, sections stay contiguous', async () => {
    const opened = await open()
    addSection(opened, 2, 'Second Half') // [default section 0-1, second half 2-4]
    const p = opened.deck.slides.map((s) => s.path)
    // Move slide 1 (default section) to index 3 (inside the second-half section)
    expect(moveSlide(opened, 0, 3)).toBe(true)
    expect(opened.deck.slides.map((s) => s.path)).toEqual([p[1], p[2], p[3], p[0], p[4]])
    expect(
      getSections(opened).map((s) => ({ name: s.name, slideIndices: s.slideIndices })),
    ).toEqual([
      { name: 'Default Section', slideIndices: [0] },
      { name: 'Second Half', slideIndices: [1, 2, 3, 4] },
    ])
    // Move back before the section start: assigned to the default section
    expect(moveSlide(opened, 3, 0)).toBe(true)
    expect(
      getSections(opened).map((s) => ({ name: s.name, slideIndices: s.slideIndices })),
    ).toEqual([
      { name: 'Default Section', slideIndices: [0, 1] },
      { name: 'Second Half', slideIndices: [2, 3, 4] },
    ])
    expect(opened.deck.slides.map((s) => s.path)).toEqual(p)

    // Round-trip: section structure consistent after save and reopen
    const reopened = await openPptx(await savePptx(opened))
    expect(
      getSections(reopened).map((s) => ({ name: s.name, slideIndices: s.slideIndices })),
    ).toEqual([
      { name: 'Default Section', slideIndices: [0, 1] },
      { name: 'Second Half', slideIndices: [2, 3, 4] },
    ])
  })

  it('moveSlide: moving all slides out of a section leaves it empty', async () => {
    const opened = await open()
    addSection(opened, 4, 'Tail Section') // [default section 0-3, tail section 4]
    expect(moveSlide(opened, 4, 0)).toBe(true)
    expect(
      getSections(opened).map((s) => ({ name: s.name, slideIndices: s.slideIndices })),
    ).toEqual([
      { name: 'Default Section', slideIndices: [0, 1, 2, 3, 4] },
      { name: 'Tail Section', slideIndices: [] },
    ])
  })

  it('setSections([]) clears sections and leaves no empty extLst in presentation.xml', async () => {
    const opened = await open()
    addSection(opened, 0, 'Only Section')
    setSections(opened, [])
    expect(getSections(opened)).toEqual([])
    expect(presXml(opened)).not.toContain('<p:extLst>')
  })
})

describe('normalizeSections (#1024)', () => {
  const sec = (name: string, slideIndices: number[]): SectionInfo => ({
    id: `{${name}}`,
    name,
    slideIndices,
  })

  it('a section starts at its minimum index, not its first', () => {
    const r = normalizeSections([sec('A', [3, 1]), sec('B', [2])], 6)
    expect(r.lead).toEqual([0])
    expect(r.starts).toEqual([1, 2])
    expect(r.sections.map((s) => s.slideIndices)).toEqual([[1], [2, 3, 4, 5]])
  })

  it('an empty section inherits the next section start; no sections covers everything', () => {
    const r = normalizeSections([sec('A', [2, 3, 4]), sec('B', [])], 5)
    expect(r.lead).toEqual([0, 1])
    expect(r.sections.map((s) => s.slideIndices)).toEqual([[2, 3, 4], []])
    expect(normalizeSections([], 4)).toEqual({
      lead: [0, 1, 2, 3],
      starts: [],
      sections: [],
    })
  })

  it('a very large slideIndices list does not overflow the call stack', () => {
    // Math.min(...indices) throws RangeError past ~1e5 arguments; the minimum is
    // now folded pairwise, so the size of the list no longer matters.
    const many = Array.from({ length: 500_000 }, (_, i) => i % 10)
    const r = normalizeSections([sec('A', many)], 10)
    expect(r.starts).toEqual([0])
    expect(r.lead).toEqual([])
    expect(r.sections[0]!.slideIndices).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 9])
  })

  it('folding pairwise matches Math.min spread for unsorted and single-element lists', () => {
    const cases: number[][] = [[5], [4, 9, 2, 7], [2, 2, 2], [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]]
    for (const indices of cases) {
      const r = normalizeSections([sec('A', indices)], 10)
      expect(r.starts).toEqual([Math.min(...indices)])
      expect(r.sections[0]!.slideIndices).toEqual(
        Array.from({ length: 10 - Math.min(...indices) }, (_, i) => Math.min(...indices) + i),
      )
    }
  })
})
