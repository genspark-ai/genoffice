/**
 * In-group editing (patching group children) tests:
 * 1. groupChildSlices: document-order slicing, nvId extraction, nested-group inner children not double-counted.
 * 2. editGroupChildTransform: nvId matching (no mismatch even when children array order ≠ document order),
 *    model and XML updated in sync, nested-group chOff/chExt preserved.
 * 3. patchGroupChildText: child txBody regenerated into the group XML.
 * 4. Integration: groupElements → edit a child → save and reopen with correct coordinates/text.
 */
import { describe, it, expect } from 'vitest'
import {
  groupChildSlices,
  editGroupChildTransform,
  patchGroupChildText,
  groupElements,
  openPptx,
  savePptx,
} from '../src/index'
import type { GroupElement, Slide, SlideElement, TextElement } from '../src/types'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const SP = (id: number, x: number, y: number) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="100" cy="100"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>hi</a:t></a:r></a:p></p:txBody></p:sp>`

const PIC = (id: number) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="p${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rId9"/></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="50" cy="50"/></a:xfrm></p:spPr></p:pic>`

const NESTED_GRP = (id: number, inner: string) =>
  `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="g${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="10" y="10"/><a:ext cx="200" cy="200"/>` +
  `<a:chOff x="0" y="0"/><a:chExt cx="200" cy="200"/></a:xfrm></p:grpSpPr>${inner}</p:grpSp>`

const GRP_XML = (inner: string) =>
  `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="1000" y="1000"/><a:ext cx="5000" cy="5000"/>` +
  `<a:chOff x="1000" y="1000"/><a:chExt cx="5000" cy="5000"/></a:xfrm></p:grpSpPr>${inner}</p:grpSp>`

function makeChild(id: string, nvId: string, type: SlideElement['type'] = 'shape'): SlideElement {
  return {
    id,
    type,
    nvId,
    anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
    transform: { offset: { x: 0, y: 0, cx: 100, cy: 100 }, rot: 0, flipH: false, flipV: false },
  } as SlideElement
}

function makeSlide(grpXml: string, children: SlideElement[]): { slide: Slide; grp: GroupElement } {
  const grp: GroupElement = {
    id: 'grp1',
    type: 'group',
    anchor: { spIndex: 0, originalXml: grpXml, range: [0, grpXml.length] },
    transform: {
      offset: { x: 1000, y: 1000, cx: 5000, cy: 5000 },
      rot: 0,
      flipH: false,
      flipV: false,
    },
    childOffset: { x: 1000, y: 1000, cx: 5000, cy: 5000 },
    children,
  }
  return { slide: { elements: [grp] } as unknown as Slide, grp }
}

describe('groupChildSlices', () => {
  it('extracts children in document order with nvId', () => {
    const slices = groupChildSlices(GRP_XML(SP(2, 0, 0) + PIC(3) + SP(4, 9, 9)))
    expect(slices.map((s) => s.nvId)).toEqual(['2', '3', '4'])
    expect(slices[1]!.xml.startsWith('<p:pic')).toBe(true)
  })

  it('nested group counts as one slice (inner children not duplicated)', () => {
    const slices = groupChildSlices(GRP_XML(SP(2, 0, 0) + NESTED_GRP(3, SP(5, 1, 1)) + SP(4, 9, 9)))
    expect(slices.map((s) => s.nvId)).toEqual(['2', '3', '4'])
    expect(slices[1]!.xml).toContain('id="5"')
  })
})

describe('editGroupChildTransform', () => {
  it('patches the nvId-matched child even when children array order differs', () => {
    // children array order deliberately reversed vs document order (parseGroup's per-tag traversal does this)
    const { slide, grp } = makeSlide(GRP_XML(SP(2, 0, 0) + PIC(3) + SP(4, 900, 900)), [
      makeChild('c4', '4'),
      makeChild('c2', '2'),
      makeChild('c3', '3', 'picture'),
    ])
    const ok = editGroupChildTransform(
      slide,
      'grp1',
      'c4',
      { x: 2000, y: 3000, cx: 400, cy: 500 },
      45,
    )
    expect(ok).toBe(true)
    const xml = grp.anchor.originalXml
    // The sp with id=4 is rewritten; the sp with id=2 stays as-is
    expect(xml).toContain('<a:off x="2000" y="3000"/><a:ext cx="400" cy="500"/>')
    expect(xml).toContain('rot="2700000"')
    expect(xml).toContain('<a:off x="0" y="0"/>')
    // Model kept in sync
    const c4 = grp.children.find((c) => c.id === 'c4')!
    expect(c4.transform.offset).toEqual({ x: 2000, y: 3000, cx: 400, cy: 500 })
    expect(c4.transform.rot).toBe(2700000)
    expect(slide.structureDirty).toBe(true)
  })

  it('keeps chOff/chExt when the child is a nested group', () => {
    const { slide, grp } = makeSlide(GRP_XML(NESTED_GRP(3, SP(5, 1, 1))), [
      makeChild('c3', '3', 'group'),
    ])
    const ok = editGroupChildTransform(slide, 'grp1', 'c3', { x: 7, y: 8, cx: 300, cy: 300 }, 0)
    expect(ok).toBe(true)
    const xml = grp.anchor.originalXml
    expect(xml).toContain(
      '<a:off x="7" y="8"/><a:ext cx="300" cy="300"/><a:chOff x="0" y="0"/><a:chExt cx="200" cy="200"/>',
    )
  })

  it('returns false when nvId missing or unmatched (model untouched)', () => {
    const { slide, grp } = makeSlide(GRP_XML(SP(2, 0, 0)), [makeChild('c9', '9')])
    expect(editGroupChildTransform(slide, 'grp1', 'c9', { x: 1, y: 1, cx: 1, cy: 1 }, 0)).toBe(
      false,
    )
    expect(grp.children[0]!.transform.offset.x).toBe(0)
    expect(slide.structureDirty).toBeFalsy()
  })
})

describe('patchGroupChildText', () => {
  it('regenerates child txBody inside group xml', () => {
    const { slide, grp } = makeSlide(GRP_XML(SP(2, 0, 0) + SP(4, 9, 9)), [
      makeChild('c2', '2'),
      makeChild('c4', '4'),
    ])
    const child = grp.children[1] as TextElement
    child.text = {
      paragraphs: [{ runs: [{ text: 'edited!' }] }],
      autofit: 'none',
      insets: { l: 0, t: 0, r: 0, b: 0 },
    } as unknown as TextElement['text']
    const ok = patchGroupChildText(slide, 'grp1', child)
    expect(ok).toBe(true)
    const xml = grp.anchor.originalXml
    expect(xml).toContain('edited!')
    // Two sp: only id=4 is regenerated; id=2's "hi" remains
    expect(xml).toContain('hi')
    expect(slide.structureDirty).toBe(true)
  })
})

describe('integration: group → edit child → save → reopen', () => {
  it('child transform edit survives roundtrip', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
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

    const result = groupElements(opened, slideIdx, ids)
    expect(result).not.toBeNull()
    const slide = opened.deck.slides[slideIdx]!
    const grp = slide.elements.find((e) => e.id === result!.groupId) as GroupElement
    expect(grp.children.length).toBe(2)
    const child = grp.children[0]!
    expect(child.nvId).toBeTruthy()

    const ok = editGroupChildTransform(
      slide,
      grp.id,
      child.id,
      { x: 123456, y: 234567, cx: 345678, cy: 456789 },
      0,
    )
    expect(ok).toBe(true)

    const out = await savePptx(opened)
    const reopened = await openPptx(out)
    const freshGrp = reopened.deck.slides[slideIdx]!.elements.find(
      (e) => e.type === 'group',
    ) as GroupElement
    expect(freshGrp).toBeDefined()
    const freshChild = freshGrp.children.find((c) => c.nvId === child.nvId)!
    expect(freshChild.transform.offset).toEqual({ x: 123456, y: 234567, cx: 345678, cy: 456789 })
  })
})

describe('group child format operations (font/fill/stroke)', () => {
  const mk = () =>
    makeSlide(GRP_XML(SP(2, 0, 0) + SP(4, 9, 9)), [makeChild('c2', '2'), makeChild('c4', '4')])

  it('editGroupChildFill: model fill updated + group XML child slice patched', async () => {
    const { editGroupChildFill } = await import('../src/index')
    const { slide, grp } = mk()
    expect(editGroupChildFill(slide, 'grp1', 'c4', '#FF8800')).toBe(true)
    expect((grp.children[1] as any).fill).toEqual({ type: 'solid', color: '#FF8800' })
    // Only the id=4 slice carries the new fill
    const slices = grp.anchor.originalXml.split('<p:sp>')
    expect(slices[2]).toContain('FF8800')
    expect(slices[1]).not.toContain('FF8800')
    expect(slide.structureDirty).toBe(true)
  })

  it('editGroupChildFill: gradient fill patched into the child slice', async () => {
    const { editGroupChildFill } = await import('../src/index')
    const { slide, grp } = mk()
    const stops = [
      { pos: 0, color: '#FF0000' },
      { pos: 1, color: '#0000FF' },
    ]
    expect(editGroupChildFill(slide, 'grp1', 'c4', { stops, angle: 5400000 })).toBe(true)
    expect((grp.children[1] as any).fill).toEqual({ type: 'gradient', stops, angle: 5400000 })
    const slices = grp.anchor.originalXml.split('<p:sp>')
    expect(slices[2]).toContain('<a:gradFill')
    expect(slices[2]).toContain('<a:lin ang="5400000"')
    expect(slices[1]).not.toContain('gradFill')

    expect(editGroupChildFill(slide, 'grp1', 'c2', { stops, radial: true })).toBe(true)
    expect((grp.children[0] as any).fill).toEqual({ type: 'gradient', stops, path: 'circle' })
    expect(grp.anchor.originalXml.split('<p:sp>')[1]).toContain('<a:path path="circle"')
  })

  it('editGroupChildStroke: adding and removing a stroke', async () => {
    const { editGroupChildStroke } = await import('../src/index')
    const { slide, grp } = mk()
    expect(editGroupChildStroke(slide, 'grp1', 'c2', { color: '#112233', widthEmu: 19050 })).toBe(
      true,
    )
    expect(grp.anchor.originalXml).toContain('112233')
    expect(editGroupChildStroke(slide, 'grp1', 'c2', null)).toBe(true)
    expect((grp.children[0] as any).stroke).toBeUndefined()
  })

  it('setGroupChildFont: run font/size updated + txBody regenerated', async () => {
    const { setGroupChildFont } = await import('../src/index')
    const { slide, grp } = mk()
    const child = grp.children[0] as any
    child.text = { paragraphs: [{ runs: [{ text: 'hi', fontSize: 18 }] }] }
    expect(setGroupChildFont(slide, 'grp1', 'c2', { fontFamily: 'Arial', fontSizePt: 30 })).toBe(
      true,
    )
    expect(child.text.paragraphs[0].runs[0].fontFamily).toBe('Arial')
    expect(child.text.paragraphs[0].runs[0].fontSize).toBe(30)
    expect(grp.anchor.originalXml).toContain('typeface="Arial"')
    expect(grp.anchor.originalXml).toContain('sz="3000"')
  })
})
