/**
 * Layout selection tests:
 *  1. listSlideLayouts enumerates the slideLayouts in the pptx
 *  2. insertSlideWithLayout inserts a new slide → rels point to the chosen layout, placeholder count correct
 *  3. save → reopen roundtrip does not break existing slides
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { openPptx, savePptx, listSlideLayouts, insertSlideWithLayout } from '../src/index'
import { resolveTarget, relsPathFor } from '../src/zip'
import { parseLayoutPlaceholders } from '../src/layout'
import { parsePlaceholderMap, resolvePlaceholderTransform } from '../src/placeholder'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

describe('listSlideLayouts', () => {
  it('returns a non-empty layout list, each entry with path / name / layoutType', async () => {
    const { archive } = await openPptx(fx('01_standard_business.pptx'))
    const layouts = listSlideLayouts(archive)
    expect(layouts.length).toBeGreaterThan(0)
    for (const lay of layouts) {
      expect(lay.path).toMatch(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/)
      expect(typeof lay.name).toBe('string')
      expect(typeof lay.layoutType).toBe('string')
      expect(Array.isArray(lay.placeholders)).toBe(true)
    }
  })

  it('parses placeholder geometry regardless of attribute order', () => {
    const xml =
      '<p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp>' +
      '<p:nvSpPr><p:cNvPr id="1" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off y="100" x="200"/><a:ext cy="900000" cx="1200000"/></a:xfrm></p:spPr>' +
      '</p:sp></p:spTree></p:cSld></p:sldLayout>'
    expect(parseLayoutPlaceholders(xml)).toEqual([
      {
        type: 'body',
        idx: '1',
        x: 200,
        y: 100,
        cx: 1200000,
        cy: 900000,
        hint: 'Click to add text',
      },
    ])
  })

  it('includes graphicFrame obj placeholders in both placeholder scans', () => {
    const xml =
      '<p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:graphicFrame>' +
      '<p:nvGraphicFramePr><p:cNvPr id="2" name="Content"/><p:cNvGraphicFramePr/><p:nvPr><p:ph type="obj" idx="0"/></p:nvPr></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></p:xfrm>' +
      '</p:graphicFrame></p:spTree></p:cSld></p:sldLayout>'
    expect(parseLayoutPlaceholders(xml)).toEqual([
      { type: 'obj', idx: '0', x: 100, y: 200, cx: 300, cy: 400, hint: 'Click to add content' },
    ])
    const map = parsePlaceholderMap(xml)
    expect(resolvePlaceholderTransform(map, undefined, 'obj', '0')?.offset).toEqual({
      x: 100,
      y: 200,
      cx: 300,
      cy: 400,
    })
  })

  it('layout paths sorted by number ascending', async () => {
    const { archive } = await openPptx(fx('01_standard_business.pptx'))
    const layouts = listSlideLayouts(archive)
    const nums = layouts.map((l) => parseInt(/(\d+)/.exec(l.path)?.[1] ?? '0', 10))
    expect(nums).toEqual([...nums].sort((a, b) => a - b))
  })
})

describe('insertSlideWithLayout', () => {
  it('new slide rels point to the chosen layout (correct after save+reopen)', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const layouts = listSlideLayouts(opened.archive)
    expect(layouts.length).toBeGreaterThan(0)

    const targetLayout = layouts[0]!
    const before = opened.deck.slides.length

    const newSlide = insertSlideWithLayout(opened, 0, targetLayout.path)
    expect(newSlide).not.toBeNull()
    expect(opened.deck.slides.length).toBe(before + 1)
    // New slide inserted at index=1 (after sourceIndex=0)
    expect(opened.deck.slides[1]).toBe(newSlide)

    // save + reopen
    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.slides.length).toBe(before + 1)

    // Verify the new slide's rels point to the chosen layout
    const newPath = reopened.deck.slides[1]!.path
    const relsPath = relsPathFor(newPath)
    const _relsXml = reopened.archive.readText(relsPath) ?? ''
    // The Target in rels has the form '../../ppt/slideLayouts/slideLayoutN.xml';
    // resolving it from the slide path should equal layoutPath
    const relEntries = [...reopened.archive.readRels(newPath).values()]
    const layoutRel = relEntries.find(
      (r) => r.type.endsWith('/slideLayout') || r.type.includes('slideLayout'),
    )
    expect(layoutRel).toBeDefined()
    const resolvedLayout = resolveTarget(newPath, layoutRel!.target)
    expect(resolvedLayout).toBe(targetLayout.path)

    // Original slide 0 is intact (still has elements)
    expect(reopened.deck.slides[0]!.elements.length).toBeGreaterThan(0)
  })

  it('when the chosen layout has placeholders, the new slide contains matching ph elements', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const layouts = listSlideLayouts(opened.archive)
    // Find a layout that has a title placeholder
    const withTitle = layouts.find((l) =>
      l.placeholders.some((p) => p.type === 'title' || p.type === 'ctrTitle'),
    )
    if (!withTitle) return // skip if the fixture has no such layout

    const newSlide = insertSlideWithLayout(opened, 0, withTitle.path)
    expect(newSlide).not.toBeNull()

    const saved = await savePptx(opened)
    const reopened = await openPptx(saved)
    const slide = reopened.deck.slides[1]!
    // The new slide's XML should contain p:ph tags (placeholder textboxes)
    const slideXml = reopened.archive.readText(slide.path) ?? ''
    expect(slideXml).toContain('<p:ph')
  })

  it('roundtrip: original slide bytes unchanged before/after the insertion', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const layouts = listSlideLayouts(opened.archive)
    const targetLayout = layouts[0]!

    // Record slide 0's XML before insertion
    const originalPath = opened.deck.slides[0]!.path
    const originalXml = opened.archive.readText(originalPath) ?? ''

    insertSlideWithLayout(opened, 0, targetLayout.path)
    const reopened = await openPptx(await savePptx(opened))

    // Slide 0 (the unchanged slide) should keep identical XML
    const afterXml = reopened.archive.readText(reopened.deck.slides[0]!.path) ?? ''
    expect(afterXml).toBe(originalXml)
  })
})
