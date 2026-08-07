/**
 * Built-in standard layouts: virtual entries for placeholder-less decks, on-demand
 * injection into the package (layout + rels + content types + master registration),
 * and application via setSlideLayout / insertSlideWithLayout.
 */
import { describe, it, expect } from 'vitest'
import {
  openPptx,
  savePptx,
  createBlankPptx,
  builtinLayoutInfos,
  ensureBuiltinLayout,
  setSlideLayout,
  insertSlideWithLayout,
  listSlideLayouts,
  shouldOfferBuiltinLayouts,
  appendRawElements,
} from '../src/index'
import { placeholderSpXml } from '../src/layout'
import { relsPathFor } from '../src/zip'

const open = async () => openPptx(await createBlankPptx())

describe('builtinLayoutInfos', () => {
  it('offers the standard set minus name collisions, geometry scaled to the deck size', async () => {
    const opened = await open()
    const existing = listSlideLayouts(opened.archive)
    expect(existing).toHaveLength(1)
    expect(existing[0]!.placeholders).toHaveLength(0)

    const infos = builtinLayoutInfos(opened.deck.size, new Set(existing.map((l) => l.name)))
    // 'Blank' collides with the blank deck's own layout name
    expect(infos.map((i) => i.name)).toEqual([
      'Title Slide',
      'Title and Content',
      'Section Header',
      'Two Content',
      'Title Only',
    ])
    expect(infos.every((i) => i.path.startsWith('builtin:'))).toBe(true)
    const title = infos[1]!.placeholders.find((p) => p.type === 'title')!
    expect(title).toMatchObject({ x: 838200, y: 365125, cx: 10515600, cy: 1325563 })
  })
})

describe('ensureBuiltinLayout', () => {
  it('injects a fully registered layout part and is idempotent', async () => {
    const opened = await open()
    const { archive, deck } = opened
    const path = ensureBuiltinLayout(archive, deck.size, 'titleContent')!
    expect(path).toBe('ppt/slideLayouts/slideLayout2.xml')

    const xml = archive.readText(path)!
    expect(xml).toContain('name="Title and Content"')
    expect(xml).toContain('type="obj"')
    expect(archive.readText(relsPathFor(path))).toContain('slideMaster1.xml')
    expect(archive.readText('[Content_Types].xml')).toContain(`PartName="/${path}"`)
    const master = archive.readText('ppt/slideMasters/slideMaster1.xml')!
    const rid = /<Relationship Id="(rId\d+)"[^>]*slideLayout2\.xml"/.exec(
      archive.readText(relsPathFor('ppt/slideMasters/slideMaster1.xml'))!,
    )![1]
    expect(master).toContain(`r:id="${rid}"`)

    expect(ensureBuiltinLayout(archive, deck.size, 'titleContent')).toBe(path)
    expect(listSlideLayouts(archive)).toHaveLength(2)
  })

  it('rejects unknown keys', async () => {
    const opened = await open()
    expect(ensureBuiltinLayout(opened.archive, opened.deck.size, 'nope')).toBeNull()
  })
})

describe('applying built-in layouts', () => {
  it('setSlideLayout re-points rels and adds the missing placeholders as empty boxes', async () => {
    const opened = await open()
    const path = ensureBuiltinLayout(opened.archive, opened.deck.size, 'twoContent')!
    const fresh = setSlideLayout(opened, 0, path)!
    expect(opened.archive.readText(relsPathFor(fresh.path))).toContain('slideLayout2.xml')

    const phs = fresh.elements.filter((e) => e.placeholder)
    expect(phs).toHaveLength(3) // title + two body
    expect(phs.map((e) => e.placeholder).sort()).toEqual(['body', 'body', 'title'])

    // Applying again: slots already filled, nothing duplicated
    const again = setSlideLayout(opened, 0, path)!
    expect(again.elements.filter((e) => e.placeholder)).toHaveLength(3)
  })

  it('insertSlideWithLayout creates a slide scaffolded with the layout placeholders', async () => {
    const opened = await open()
    const path = ensureBuiltinLayout(opened.archive, opened.deck.size, 'titleSlide')!
    const slide = insertSlideWithLayout(opened, 0, path)!
    const types = slide.elements.map((e) => e.placeholder).sort()
    expect(types).toEqual(['ctrTitle', 'subTitle'])
  })

  it('slide footer/date/number placeholders do not block content slots', async () => {
    const opened = await open()
    // ftr idx=2 on the slide must not count as content slot body:2
    appendRawElements(opened, 0, [
      placeholderSpXml({ type: 'ftr', idx: '2', x: 0, y: 6300000, cx: 3000000, cy: 365125 }, 50),
    ])
    const path = ensureBuiltinLayout(opened.archive, opened.deck.size, 'twoContent')!
    const fresh = setSlideLayout(opened, 0, path)!
    const bodies = fresh.elements.filter((e) => e.placeholder === 'body')
    expect(bodies).toHaveLength(2)
  })

  it('remaining built-ins stay on offer after one is injected (save/reopen)', async () => {
    const opened = await open()
    ensureBuiltinLayout(opened.archive, opened.deck.size, 'titleSlide')
    const reopened = await openPptx(await savePptx(opened))
    const layouts = listSlideLayouts(reopened.archive)
    expect(shouldOfferBuiltinLayouts(layouts)).toBe(true)
    const offered = builtinLayoutInfos(reopened.deck.size, new Set(layouts.map((l) => l.name)))
    expect(offered.map((i) => i.name)).toEqual([
      'Title and Content',
      'Section Header',
      'Two Content',
      'Title Only',
    ])
    // A deck with a foreign placeholder-bearing layout keeps its own template untouched
    expect(shouldOfferBuiltinLayouts([{ name: 'Custom', placeholders: [{}] }])).toBe(false)
  })

  it('survives a save/reopen round-trip', async () => {
    const opened = await open()
    const path = ensureBuiltinLayout(opened.archive, opened.deck.size, 'titleContent')!
    setSlideLayout(opened, 0, path)
    const reopened = await openPptx(await savePptx(opened))
    expect(listSlideLayouts(reopened.archive).map((l) => l.name)).toContain('Title and Content')
    expect(reopened.deck.slides[0]!.elements.filter((e) => e.placeholder)).toHaveLength(2)
  })
})
