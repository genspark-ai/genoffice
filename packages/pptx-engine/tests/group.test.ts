/**
 * Element group/ungroup (groupElements / ungroupElement) tests:
 * 1. buildGrpSpXml unit tests (generates correct p:grpSp XML).
 * 2. calcBoundingBox unit tests (bounding-box computation).
 * 3. groupElements integration: after grouping, save and reopen; child coordinates correct, original bytes passed through.
 * 4. ungroupElement integration: after ungrouping, children promoted to top level with correct coordinate conversion.
 * 5. Edge cases: fewer than 2 elements / contains passthrough → rejected.
 */
import { describe, it, expect } from 'vitest'
import { calcBoundingBox, buildGrpSpXml } from '../src/insert'
import {
  appendRawElements,
  createBlankPptx,
  groupElements,
  ungroupElement,
  openPptx,
  savePptx,
} from '../src/index'
import { parseSlide } from '../src/parse'
import type { GroupElement, SlideElement } from '../src/types'
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

// ── Unit tests: calcBoundingBox ──────────────────────────────────────────────

describe('calcBoundingBox', () => {
  const makeEl = (x: number, y: number, cx: number, cy: number): SlideElement => ({
    id: 'el',
    type: 'shape',
    anchor: { spIndex: 0, originalXml: '', range: [0, 0] },
    transform: { offset: { x, y, cx, cy }, rot: 0, flipH: false, flipV: false },
  })

  it('calculates bounding box from two elements', () => {
    const a = makeEl(100, 200, 300, 400)
    const b = makeEl(500, 100, 200, 300)
    const bbox = calcBoundingBox([a, b])
    expect(bbox.x).toBe(100)
    expect(bbox.y).toBe(100)
    expect(bbox.cx).toBe(600) // 500+200 - 100
    expect(bbox.cy).toBe(500) // max(200+400, 100+300) - 100 = 600 - 100 = 500
  })

  it('single element → bbox equals element offset', () => {
    const a = makeEl(50, 60, 100, 200)
    const bbox = calcBoundingBox([a])
    expect(bbox).toEqual({ x: 50, y: 60, cx: 100, cy: 200 })
  })
})

// ── Unit tests: buildGrpSpXml ────────────────────────────────────────────────

const MAKE_SP = (x: number, y: number, cx: number, cy: number) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="sp"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`

const MOCK_SLIDE: any = {
  elements: [
    {
      id: 'el1',
      type: 'shape',
      anchor: { originalXml: MAKE_SP(100, 100, 200, 200) },
      transform: { offset: { x: 100, y: 100, cx: 200, cy: 200 } },
    },
  ],
  originalXml: '<p:sld><p:cSld><p:spTree><p:grpSpPr/></p:spTree></p:cSld></p:sld>',
}

describe('buildGrpSpXml', () => {
  it('generates valid p:grpSp with chOff/chExt matching bbox', () => {
    const bbox = { x: 100, y: 100, cx: 500, cy: 400 }
    const childXml = MAKE_SP(100, 100, 200, 200)
    const xml = buildGrpSpXml(MOCK_SLIDE, bbox, childXml)
    expect(xml).toContain('<p:grpSp>')
    expect(xml).toContain('<p:grpSpPr>')
    expect(xml).toContain(`<a:off x="100" y="100"/>`)
    expect(xml).toContain(`<a:ext cx="500" cy="400"/>`)
    expect(xml).toContain(`<a:chOff x="100" y="100"/>`)
    expect(xml).toContain(`<a:chExt cx="500" cy="400"/>`)
    expect(xml).toContain('</p:grpSp>')
    // Child element XML kept as-is
    expect(xml).toContain('<p:sp>')
  })
})

// ── Integration tests: groupElements + ungroupElement ────────────────────────

describe('groupElements integration', () => {
  it('rejects group with < 2 elements', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const firstId = slide.elements.find((e) => e.type === 'text' || e.type === 'shape')?.id
    if (!firstId) return
    const result = groupElements(opened, 0, [firstId])
    expect(result).toBeNull()
  })

  it('rejects group with passthrough elements', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const editableIds = slide.elements
      .filter((e) => e.type === 'text' || e.type === 'shape' || e.type === 'picture')
      .slice(0, 1)
      .map((e) => e.id)
    const ptIds = slide.elements
      .filter((e) => e.type === 'passthrough')
      .slice(0, 1)
      .map((e) => e.id)
    if (editableIds.length === 0 || ptIds.length === 0) return
    const result = groupElements(opened, 0, [...editableIds, ...ptIds])
    expect(result).toBeNull()
  })

  it('groups 2 editable elements, saves and reopens correctly', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    // Find groupable elements
    let slideIdx = -1
    let ids: string[] = []
    for (let i = 0; i < opened.deck.slides.length; i++) {
      const els = opened.deck.slides[i]!.elements.filter(
        (e) => e.type === 'text' || e.type === 'shape' || e.type === 'picture',
      )
      if (els.length >= 2) {
        slideIdx = i
        ids = els.slice(0, 2).map((e) => e.id)
        break
      }
    }
    if (slideIdx < 0) return // fixture has too few groupable elements

    const slide = opened.deck.slides[slideIdx]!
    const originalCount = slide.elements.length

    // Record the selected elements' bounding box (for verification)
    const elA = slide.elements.find((e) => e.id === ids[0])!
    const elB = slide.elements.find((e) => e.id === ids[1])!
    const minX = Math.min(elA.transform.offset.x, elB.transform.offset.x)
    const minY = Math.min(elA.transform.offset.y, elB.transform.offset.y)

    const result = groupElements(opened, slideIdx, ids)
    expect(result).not.toBeNull()
    // Element count drops by one after grouping (2 become 1 group)
    expect(result!.slide.elements.length).toBe(originalCount - 1)
    // The group element exists
    const grp = result!.slide.elements.find((e) => e.id === result!.groupId)
    expect(grp).toBeDefined()
    expect(grp!.type).toBe('group')
    // Group position = top-left of the bounding box
    expect(grp!.transform.offset.x).toBe(minX)
    expect(grp!.transform.offset.y).toBe(minY)

    // Save and reopen
    const out = await savePptx(opened)
    const reopened = await openPptx(out)
    const freshSlide = reopened.deck.slides[slideIdx]!
    const freshGrp = freshSlide.elements.find((e) => e.type === 'group')
    expect(freshGrp).toBeDefined()
    // The group contains 2 child elements
    const g = freshGrp as import('../src/types').GroupElement
    expect(g.children.length).toBe(2)
  })

  it('ungroupElement lifts children to top level', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    // Find groupable elements
    let slideIdx = -1
    let ids: string[] = []
    for (let i = 0; i < opened.deck.slides.length; i++) {
      const els = opened.deck.slides[i]!.elements.filter(
        (e) => e.type === 'text' || e.type === 'shape' || e.type === 'picture',
      )
      if (els.length >= 2) {
        slideIdx = i
        ids = els.slice(0, 2).map((e) => e.id)
        break
      }
    }
    if (slideIdx < 0) return

    const original = opened.deck.slides[slideIdx]!
    const originalCount = original.elements.length

    // Group first
    const grouped = groupElements(opened, slideIdx, ids)
    if (!grouped) return

    // Then ungroup
    const fresh = ungroupElement(opened, slideIdx, grouped.groupId)
    expect(fresh).not.toBeNull()

    // After ungrouping, element count returns to originalCount (2 children promoted, group gone = originalCount)
    // Note: after materialize, ids of the original originalXml children may change, so only check the count
    expect(fresh!.elements.length).toBe(originalCount)
  })

  it('nested group ungroup: only direct children promoted, sp inside the nested group not extracted by mistake', async () => {
    const sp = (name: string, x: number, y: number, cx: number, cy: number) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`
    const inner =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="21" name="inner"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000" cy="2000"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="2000" cy="2000"/></a:xfrm></p:grpSpPr>` +
      sp('X', 0, 0, 1000, 1000) +
      `</p:grpSp>`
    const outer =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="outer"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="1000" y="1000"/><a:ext cx="4000" cy="4000"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="4000" cy="4000"/></a:xfrm></p:grpSpPr>` +
      inner +
      sp('Y', 2000, 2000, 1000, 1000) +
      `</p:grpSp>`

    const opened = await openPptx(await createBlankPptx())
    const before = opened.deck.slides[0]!.elements.length
    const appended = appendRawElements(opened, 0, [outer])!
    const grpId = appended.elementIds[appended.elementIds.length - 1]!

    const fresh = ungroupElement(opened, 0, grpId)
    expect(fresh).not.toBeNull()
    // 2 direct children: the inner group + sp Y (sp X inside the nested group is not promoted separately)
    expect(fresh!.elements.length).toBe(before + 2)
    const grps = fresh!.elements.filter((e) => e.type === 'group') as GroupElement[]
    expect(grps.length).toBe(1)
    expect(grps[0]!.children.length).toBe(1)
    // chOff/chExt 1:1, coordinates = child coordinates + group offset
    expect(grps[0]!.transform.offset).toMatchObject({ x: 1000, y: 1000, cx: 2000, cy: 2000 })
    const spY = fresh!.elements.find((e) => e.type !== 'group' && e.name === 'Y')!
    expect(spY.transform.offset).toMatchObject({ x: 3000, y: 3000, cx: 1000, cy: 1000 })
  })

  it('roundtrip: no group edits → slide bytes unchanged (fidelity)', async () => {
    const original = fx('01_standard_business.pptx')
    const opened = await openPptx(original)
    const out = await savePptx(opened)
    const origZip = await JSZip.loadAsync(original)
    const outZip = await JSZip.loadAsync(out)
    for (const slide of opened.deck.slides) {
      const a = await origZip.file(slide.path)!.async('uint8array')
      const b = await outZip.file(slide.path)!.async('uint8array')
      expect(Buffer.from(b).equals(Buffer.from(a))).toBe(true)
    }
  })
})

// ── mixed-type group children keep document order ────────────────────────────

const SLIDE_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'

describe('parseGroup document order', () => {
  it('group mixing sp/pic: children keep document order (z-order)', () => {
    const sp = (name: string) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`
    const pic =
      `<p:pic><p:nvPicPr><p:cNvPr id="3" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rId99"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="10" y="10"/><a:ext cx="20" cy="20"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
    const grp =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></p:grpSpPr>` +
      sp('A') +
      pic +
      sp('B') +
      `</p:grpSp>`
    const slideXml =
      `<?xml version="1.0"?><p:sld ${SLIDE_NS}><p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      grp +
      `</p:spTree></p:cSld></p:sld>`
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
    const g = slide.elements[0] as GroupElement
    expect(g.type).toBe('group')
    expect(g.children.map((c) => c.name)).toEqual(['A', 'P', 'B'])
  })
})
