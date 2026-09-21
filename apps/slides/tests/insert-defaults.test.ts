import { beforeEach, describe, expect, it, vi } from 'vitest'
import {
  EQUATION_BODY_PR,
  EQUATION_BOX_PX,
  EQUATION_FONT_FAMILY,
  EQUATION_FONT_PT,
  equationInsertFrame,
  graphicFrameInsertFrame,
  summaryZoomLayout,
  zoomCascadeFrames,
  zoomInsertFrame,
} from '../src/renderer/insert-defaults'
import {
  insertChart,
  insertEquation,
  insertSmartArt,
  insertZoom,
} from '../src/renderer/insert-actions'
import {
  insertSectionZooms,
  insertSlideZooms,
  insertSummaryZoom,
  summaryZoomSections,
  slideTitleText,
  summaryLayoutPath,
} from '../src/renderer/zoom-actions'
import { CHART_GALLERY, SMARTART_GALLERY } from '../src/renderer/insert-presets'
import { t } from '../src/renderer/i18n/locale'
import { renderSlidesToPngBase64 } from '../src/renderer/export-render'
import type { ActionCtx } from '../src/renderer/action-context'

vi.mock('../src/renderer/export-render', () => ({
  renderSlidesToPngBase64: vi.fn(),
}))

const PX_PER_PT = 96 / 72
const emuPx = (emu: number, scale: number) => (emu / 12700) * PX_PER_PT * scale
// The viewport is always FIT_WIDTH wide; a 4:3 deck is taller and its px are 4/3 the size of a 96 dpi px
const wide = { widthPx: 1280, heightPx: 720, scale: 1 } // 960 x 540 pt
const standard = { widthPx: 1280, heightPx: 960, scale: 4 / 3 } // 720 x 540 pt

/** PowerPoint for Mac measurement in pt vs the px frame we insert (within 1 px). */
function expectPt(
  slide: { scale: number },
  frame: { x: number; y: number; w: number; h: number },
  pt: number[],
) {
  const [x, y, w, h] = pt.map((v) => v * PX_PER_PT * slide.scale)
  expect(Math.abs(frame.x - x!)).toBeLessThan(1)
  expect(Math.abs(frame.y - y!)).toBeLessThan(1)
  expect(Math.abs(frame.w - w!)).toBeLessThan(1)
  expect(Math.abs(frame.h - h!)).toBeLessThan(1)
}

describe('insert default frames (PowerPoint for Mac measurements)', () => {
  it('equation: a 1 in square centered, Cambria Math italic 18 pt, wrap="none" + spAutoFit', () => {
    expect(EQUATION_BOX_PX).toBe(96)
    expect(EQUATION_FONT_PT).toBe(18)
    expect(EQUATION_FONT_FAMILY).toBe('Cambria Math')
    expect(EQUATION_BODY_PR).toEqual({ wrap: 'none', autoFit: 'resize' })
    expectPt(wide, equationInsertFrame(wide), [444, 234, 72, 72])
    expectPt(standard, equationInsertFrame(standard), [324, 234, 72, 72])
  })

  it('chart / SmartArt: 2/3 of the slide width in a 3:2 aspect, centered on both deck sizes', () => {
    expectPt(wide, graphicFrameInsertFrame(wide), [160, 56.67, 640, 426.67])
    expectPt(standard, graphicFrameInsertFrame(standard), [120, 110, 480, 320])
  })

  it('slide zoom: a quarter-size thumbnail centered', () => {
    expect(zoomInsertFrame(wide)).toEqual({ x: 480, y: 270, w: 320, h: 180 })
    expectPt(standard, zoomInsertFrame(standard), [270, 202.5, 180, 135])
  })

  it('several zooms cascade 12 pt apart around the centered frame, later tiles on top', () => {
    expect(zoomCascadeFrames(wide, 1)).toEqual([zoomInsertFrame(wide)])
    const three = zoomCascadeFrames(wide, 3)
    expectPt(wide, three[0]!, [348, 190.5, 240, 135])
    expectPt(wide, three[1]!, [360, 202.5, 240, 135])
    expectPt(wide, three[2]!, [372, 214.5, 240, 135])
    const two = zoomCascadeFrames(standard, 2)
    expectPt(standard, two[0]!, [264, 196.5, 180, 135])
    expectPt(standard, two[1]!, [276, 208.5, 180, 135])
  })

  describe('summary zoom grid (Title and Content body box, EMU measured on 960 x 540)', () => {
    const box = {
      x: emuPx(838200, 1),
      y: emuPx(1825625, 1),
      w: emuPx(10515600, 1),
      h: emuPx(4351338, 1),
    }
    const expectEmu = (frame: { x: number; y: number; w: number; h: number }, emu: number[]) => {
      const [x, y, w, h] = emu.map((v) => emuPx(v, 1))
      expect(Math.abs(frame.x - x!)).toBeLessThan(1)
      expect(Math.abs(frame.y - y!)).toBeLessThan(1)
      expect(Math.abs(frame.w - w!)).toBeLessThan(1)
      expect(Math.abs(frame.h - h!)).toBeLessThan(1)
    }

    it('n=3: 2 x 2, last row left-aligned', () => {
      const tiles = summaryZoomLayout(3, box, 16 / 9)
      expect(tiles).toHaveLength(3)
      expectEmu(tiles[0]!, [838200 + 1711460, 1825625 + 152297, 3481070, 1958102])
      expectEmu(tiles[1]!, [838200 + 5323070, 1825625 + 152297, 3481070, 1958102])
      expectEmu(tiles[2]!, [838200 + 1711460, 1825625 + 2240939, 3481070, 1958102])
    })

    it('n=7: 4 x 2 with a three-tile second row', () => {
      const tiles = summaryZoomLayout(7, box, 16 / 9)
      expect(tiles).toHaveLength(7)
      const x0 = 838200 + 392692
      const y0 = 1825625 + 800426
      tiles.forEach((tile, i) => {
        expectEmu(tile, [
          x0 + (i % 4) * 2454735,
          y0 + Math.floor(i / 4) * 1419605,
          2366010,
          1330880,
        ])
      })
    })

    it('tiles keep the deck aspect and stay inside the box on a 4:3 deck', () => {
      const b = { x: 120, y: 200, w: 1040, h: 560 }
      for (const n of [1, 2, 5, 12]) {
        for (const tile of summaryZoomLayout(n, b, 4 / 3)) {
          expect(Math.abs(tile.w / tile.h - 4 / 3)).toBeLessThan(0.02)
          expect(tile.x).toBeGreaterThanOrEqual(b.x)
          expect(tile.y).toBeGreaterThanOrEqual(b.y)
          expect(tile.x + tile.w).toBeLessThanOrEqual(b.x + b.w + 1)
          expect(tile.y + tile.h).toBeLessThanOrEqual(b.y + b.h + 1)
        }
      }
      expect(summaryZoomLayout(0, b, 4 / 3)).toEqual([])
    })
  })
})

describe('insert actions use the measured frames', () => {
  const api = {
    addElement: vi.fn(),
    addChart: vi.fn(),
    addSmartArt: vi.fn(),
    addImageBytes: vi.fn(),
    setLink: vi.fn(),
    beginHistoryBatch: vi.fn(),
    endHistoryBatch: vi.fn(),
    addSlideWithLayout: vi.fn(),
    addBlankSlide: vi.fn(),
    moveSlide: vi.fn(),
    deleteElement: vi.fn(),
    setSections: vi.fn(),
    getLayouts: vi.fn(),
  }
  const applySlide = vi.fn()
  const setSelectedIds = vi.fn()
  const setEditing = vi.fn()
  const setStatus = vi.fn()
  const setEqDialogOpen = vi.fn()
  const setSlides = vi.fn()
  const setSections = vi.fn()
  const setCurrent = vi.fn()
  const setSelectedSlides = vi.fn()
  const setDirty = vi.fn()
  const titled = (text: string) => ({
    type: 'text',
    placeholder: 'title',
    sourceId: 't1',
    box: { x: 0, y: 0, w: 100, h: 40 },
    text: { lines: [{ runs: [{ text: '\u2022', isBullet: true }, { text }] }] },
  })
  const slides = [
    { ...wide, id: 'a', nodes: [] },
    { ...wide, id: 'b', nodes: [titled('Agenda')] },
    { ...wide, id: 'c', nodes: [] },
    { ...wide, id: 'd', nodes: [] },
  ]
  const layouts = [
    { path: 'L1', name: 'Title Slide', layoutType: 'title', placeholders: [] },
    { path: 'L2', name: 'Title and Content', layoutType: 'obj', placeholders: [] },
  ]
  const ctx = {
    slide: slides[0],
    slides,
    images: new Map(),
    current: 0,
    sections: [],
    layouts,
    applySlide,
    setSelectedIds,
    setEditing,
    setStatus,
    setEqDialogOpen,
    setSlides,
    setSections,
    setCurrent,
    setSelectedSlides,
    setDirty,
  } as unknown as ActionCtx

  beforeEach(() => {
    vi.clearAllMocks()
    ;(window as unknown as { slidesApi: unknown }).slidesApi = api
    const result = { slide: slides[0], sourceId: 'sp9' }
    api.addElement.mockResolvedValue(result)
    api.addChart.mockResolvedValue(result)
    api.addSmartArt.mockResolvedValue(result)
    api.addImageBytes.mockResolvedValue(result)
    api.setLink.mockResolvedValue(slides[0])
    api.beginHistoryBatch.mockResolvedValue(true)
    api.endHistoryBatch.mockResolvedValue(null)
    // The engine mints a GUID for every section written without an id
    api.setSections.mockImplementation(async (list: Array<{ id: string; name: string }>) =>
      list.map((sec) => ({ ...sec, id: sec.id || `{${sec.name}}` })),
    )
    vi.mocked(renderSlidesToPngBase64).mockResolvedValue(['UE5H'])
  })

  it('insertEquation drops the auto-fit square and opens it for editing', async () => {
    await insertEquation(ctx, 'E = mc²')
    expect(setEqDialogOpen).toHaveBeenCalledWith(false)
    expect(api.addElement).toHaveBeenCalledWith({
      slideIndex: 0,
      kind: 'textbox',
      xPx: 592,
      yPx: 312,
      wPx: 96,
      hPx: 96,
      fitWidthPx: 1280,
      bodyPr: { wrap: 'none', autoFit: 'resize' },
      paragraphs: [
        { runs: [{ text: 'E = mc²', fontSize: 18, italic: true, fontFamily: 'Cambria Math' }] },
      ],
    })
    expect(setSelectedIds).toHaveBeenCalledWith(['sp9'])
    expect(setEditing).toHaveBeenCalledWith({ sourceId: 'sp9' })
  })

  it('insertChart and insertSmartArt share the 2/3-width 3:2 frame', async () => {
    await insertChart(ctx, CHART_GALLERY[0]!.kind)
    expect(api.addChart).toHaveBeenCalledWith(
      expect.objectContaining({ slideIndex: 0, xPx: 214, yPx: 76, wPx: 853, hPx: 569 }),
    )
    await insertSmartArt(ctx, SMARTART_GALLERY[0]!)
    expect(api.addSmartArt).toHaveBeenCalledWith(
      expect.objectContaining({ slideIndex: 0, xPx: 214, yPx: 76, wPx: 853, hPx: 569 }),
    )
  })

  it('insertZoom inserts a quarter-size snapshot of the target slide linked to it', async () => {
    await insertZoom(ctx, 2)
    expect(renderSlidesToPngBase64).toHaveBeenCalledWith([slides[2]], ctx.images, 1)
    expect(api.addImageBytes).toHaveBeenCalledWith({
      slideIndex: 0,
      base64: 'UE5H',
      ext: 'png',
      fitWidthPx: 1280,
      xPx: 480,
      yPx: 270,
      wPx: 320,
      hPx: 180,
      name: 'Slide Zoom 3',
    })
    expect(api.setLink).toHaveBeenCalledWith({
      slideIndex: 0,
      sourceId: 'sp9',
      target: { kind: 'slide', slideIndex: 2 },
    })
    expect(applySlide).toHaveBeenCalledWith(0, slides[0])
    expect(setSelectedIds).toHaveBeenCalledWith(['sp9'])
  })

  it('insertZoom does nothing for a missing target slide', async () => {
    await insertZoom(ctx, 7)
    expect(renderSlidesToPngBase64).not.toHaveBeenCalled()
    expect(api.addImageBytes).not.toHaveBeenCalled()
    expect(api.beginHistoryBatch).not.toHaveBeenCalled()
  })

  it('insertSlideZooms cascades one tile per slide in one undo step, later tiles on top', async () => {
    vi.mocked(renderSlidesToPngBase64).mockResolvedValue(['P1', 'P2', 'P3'])
    let n = 0
    api.addImageBytes.mockImplementation(async () => ({ slide: slides[0], sourceId: `sp${++n}` }))
    await insertSlideZooms(ctx, [3, 1, 2, 1])
    expect(renderSlidesToPngBase64).toHaveBeenCalledWith(
      [slides[1], slides[2], slides[3]],
      ctx.images,
      1,
    )
    const frames = zoomCascadeFrames(wide, 3)
    expect(api.addImageBytes.mock.calls.map((c) => c[0])).toEqual(
      [1, 2, 3].map((target, i) => ({
        slideIndex: 0,
        base64: `P${i + 1}`,
        ext: 'png',
        fitWidthPx: 1280,
        xPx: frames[i]!.x,
        yPx: frames[i]!.y,
        wPx: frames[i]!.w,
        hPx: frames[i]!.h,
        name: `Slide Zoom ${target + 1}`,
      })),
    )
    expect(api.setLink.mock.calls.map((c) => c[0].target.slideIndex)).toEqual([1, 2, 3])
    expect(api.beginHistoryBatch).toHaveBeenCalledTimes(1)
    expect(api.endHistoryBatch).toHaveBeenCalledTimes(1)
    expect(setSelectedIds).toHaveBeenCalledWith(['sp1', 'sp2', 'sp3'])
    expect(setStatus).toHaveBeenCalledWith(t('appStatusZoomsInserted', { count: 3 }))
  })

  it('insertSectionZooms snapshots each section head slide (lead group = Default Section)', async () => {
    const sectioned = {
      ...ctx,
      sections: [
        { id: '{A}', name: 'Intro', slideIndices: [1] },
        { id: '{B}', name: 'Body', slideIndices: [2, 3] },
      ],
    } as unknown as ActionCtx
    vi.mocked(renderSlidesToPngBase64).mockResolvedValue(['P1', 'P2'])
    await insertSectionZooms(sectioned, [2, 0])
    expect(renderSlidesToPngBase64).toHaveBeenCalledWith([slides[0], slides[2]], ctx.images, 1)
    expect(api.addImageBytes.mock.calls.map((c) => c[0].name)).toEqual([
      'Section Zoom 1',
      'Section Zoom 3',
    ])
    expect(api.setLink.mock.calls.map((c) => c[0].target.slideIndex)).toEqual([0, 2])
  })

  it('summaryLayoutPath prefers layoutType obj, then a title + one body layout', () => {
    expect(summaryLayoutPath(layouts)).toBe('L2')
    const custom = [
      { path: 'X', name: 'Cover', layoutType: 'custom', placeholders: [{ type: 'ctrTitle' }] },
      {
        path: 'Y',
        name: 'Text',
        layoutType: 'custom',
        placeholders: [{ type: 'title' }, { type: '' }],
      },
    ] as unknown as typeof layouts
    expect(summaryLayoutPath(custom)).toBe('Y')
    expect(summaryLayoutPath(custom.slice(0, 1))).toBeNull()
    expect(summaryLayoutPath(null)).toBeNull()
  })

  it('slideTitleText joins the title placeholder runs without bullets', () => {
    expect(slideTitleText(slides[1] as never)).toBe('Agenda')
    expect(slideTitleText(slides[0] as never)).toBe('')
  })

  it('insertSummaryZoom adds a Title and Content slide before the first pick, fills the body box, sections the picks', async () => {
    const body = {
      type: 'shape',
      placeholder: 'body',
      sourceId: 'ph2',
      box: { x: 88, y: 192, w: 1104, h: 457 },
    }
    const fresh = {
      ...wide,
      id: 'new',
      nodes: [{ ...body, placeholder: 'title', sourceId: 'ph1' }, body],
    }
    const cleared = { ...fresh, nodes: [fresh.nodes[0]] }
    const after = [slides[0], fresh, slides[1], slides[2], slides[3]]
    api.addSlideWithLayout.mockResolvedValue({ slides: after, index: 1 })
    api.deleteElement.mockResolvedValue(cleared)
    vi.mocked(renderSlidesToPngBase64).mockResolvedValue(['P1', 'P2'])
    let n = 0
    api.addImageBytes.mockImplementation(async () => ({ slide: cleared, sourceId: `sp${++n}` }))
    api.setLink.mockResolvedValue(cleared)

    await insertSummaryZoom(ctx, [3, 1])

    expect(api.addSlideWithLayout).toHaveBeenCalledWith({
      sourceIndex: 0,
      layoutPath: 'L2',
      fitWidthPx: 1280,
    })
    expect(api.moveSlide).not.toHaveBeenCalled()
    expect(api.deleteElement).toHaveBeenCalledWith({ slideIndex: 1, sourceId: 'ph2' })
    // Picks moved one slot down; snapshots come from the post-insert deck
    expect(renderSlidesToPngBase64).toHaveBeenCalledWith([slides[1], slides[3]], ctx.images, 1)
    const frames = summaryZoomLayout(2, body.box, 16 / 9)
    expect(api.addImageBytes.mock.calls.map((c) => c[0])).toEqual(
      ['P1', 'P2'].map((base64, i) => ({
        slideIndex: 1,
        base64,
        ext: 'png',
        fitWidthPx: 1280,
        xPx: frames[i]!.x,
        yPx: frames[i]!.y,
        wPx: frames[i]!.w,
        hPx: frames[i]!.h,
        name: 'Summary Zoom',
      })),
    )
    expect(api.setLink.mock.calls.map((c) => c[0].target)).toEqual([
      { kind: 'slide', slideIndex: 2 },
      { kind: 'slide', slideIndex: 4 },
    ])
    // Slide 1 stays unsectioned in front (PowerPoint's Default Section); one write for the rest
    expect(api.setSections).toHaveBeenCalledTimes(1)
    expect(api.setSections).toHaveBeenCalledWith([
      { id: '', name: t('appSectionSummary'), slideIndices: [1] },
      { id: '', name: 'Agenda', slideIndices: [2, 3] },
      { id: '', name: t('appSectionN', { n: 5 }), slideIndices: [4] },
    ])
    expect(setSections).toHaveBeenCalledWith([
      { id: `{${t('appSectionSummary')}}`, name: t('appSectionSummary'), slideIndices: [1] },
      { id: '{Agenda}', name: 'Agenda', slideIndices: [2, 3] },
      expect.objectContaining({ slideIndices: [4] }),
    ])
    expect(api.beginHistoryBatch).toHaveBeenCalledTimes(1)
    expect(api.endHistoryBatch).toHaveBeenCalledTimes(1)
    expect(setSlides).toHaveBeenCalledWith([slides[0], cleared, slides[1], slides[2], slides[3]])
    expect(setCurrent).toHaveBeenCalledWith(1)
    expect(setSelectedSlides).toHaveBeenCalledWith([1])
    expect(setSelectedIds).toHaveBeenCalledWith([])
    expect(setDirty).toHaveBeenCalledWith(true)
    expect(setStatus).toHaveBeenCalledWith(t('appStatusSummaryZoomInserted', { count: 2 }))
  })

  /** Deck sectioned as Intro = slides 1-2, Body = slides 3-4 (both heads: 0 and 2). */
  const sectioned = [
    { id: '{A}', name: 'Intro', slideIndices: [0, 1] },
    { id: '{B}', name: 'Body', slideIndices: [2, 3] },
  ]

  it('summaryZoomSections: existing sections keep id/name/slides, chosen non-heads start one, summary stands alone', () => {
    const nameOf = (pos: number) => `S${pos + 1}`
    // Default scenario: every head chosen, slide 1 included -> summary lands at 0
    expect(summaryZoomSections(sectioned, 4, 0, [0, 2], nameOf)).toEqual([
      { id: '', name: t('appSectionSummary'), slideIndices: [0] },
      { id: '{A}', name: 'Intro', slideIndices: [1, 2] },
      { id: '{B}', name: 'Body', slideIndices: [3, 4] },
    ])
    // Middle pick: Intro shrinks to the slide before the summary, Body shifts, non-heads split
    expect(summaryZoomSections(sectioned, 4, 1, [1, 3], nameOf)).toEqual([
      { id: '{A}', name: 'Intro', slideIndices: [0] },
      { id: '', name: t('appSectionSummary'), slideIndices: [1] },
      { id: '', name: 'S3', slideIndices: [2] },
      { id: '{B}', name: 'Body', slideIndices: [3] },
      { id: '', name: 'S5', slideIndices: [4] },
    ])
    // Unsectioned lead slides stay unsectioned; a chosen lead slide heads a new section
    expect(
      summaryZoomSections(
        [{ id: '{B}', name: 'Body', slideIndices: [2, 3] }],
        4,
        1,
        [1, 2],
        nameOf,
      ),
    ).toEqual([
      { id: '', name: t('appSectionSummary'), slideIndices: [1] },
      { id: '', name: 'S3', slideIndices: [2] },
      { id: '{B}', name: 'Body', slideIndices: [3, 4] },
    ])
  })

  it('insertSummaryZoom with slide 1 (a section head) chosen lands at 0 without splitting or emptying sections', async () => {
    const fresh = { ...wide, id: 'new', nodes: [] }
    api.addSlideWithLayout.mockResolvedValue({
      slides: [slides[0], fresh, slides[1], slides[2], slides[3]],
      index: 1,
    })
    const moved = [fresh, slides[0], slides[1], slides[2], slides[3]]
    // What the engine really reports after the move: the summary slide folded into Intro
    api.moveSlide.mockResolvedValue({
      slides: moved,
      sections: [
        { id: '{A}', name: 'Intro', slideIndices: [0, 1, 2] },
        { id: '{B}', name: 'Body', slideIndices: [3, 4] },
      ],
    })
    api.addImageBytes.mockResolvedValue({ slide: fresh, sourceId: 'sp1' })
    api.setLink.mockResolvedValue(fresh)
    const sctx = { ...ctx, sections: sectioned } as unknown as ActionCtx

    await insertSummaryZoom(sctx, [0, 2])

    expect(api.moveSlide).toHaveBeenCalledWith({ fromIndex: 1, toIndex: 0 })
    expect(api.deleteElement).not.toHaveBeenCalled()
    // No body placeholder: the grid falls back to the chart/SmartArt frame
    const frames = summaryZoomLayout(2, graphicFrameInsertFrame(wide), 16 / 9)
    expect(api.addImageBytes.mock.calls[0]![0]).toEqual(
      expect.objectContaining({ slideIndex: 0, xPx: frames[0]!.x, yPx: frames[0]!.y }),
    )
    expect(api.setSections).toHaveBeenCalledTimes(1)
    const written = api.setSections.mock.calls[0]![0] as Array<{
      id: string
      name: string
      slideIndices: number[]
    }>
    expect(written).toEqual([
      { id: '', name: t('appSectionSummary'), slideIndices: [0] },
      { id: '{A}', name: 'Intro', slideIndices: [1, 2] },
      { id: '{B}', name: 'Body', slideIndices: [3, 4] },
    ])
    expect(written.every((sec) => sec.slideIndices.length > 0)).toBe(true)
    expect(new Set(written.map((sec) => sec.slideIndices[0])).size).toBe(written.length)
    expect(setCurrent).toHaveBeenCalledWith(0)
  })

  it('insertSummaryZoom with a mid-deck first pick keeps the surrounding sections intact', async () => {
    const fresh = { ...wide, id: 'new', nodes: [] }
    api.addSlideWithLayout.mockResolvedValue({
      slides: [slides[0], fresh, slides[1], slides[2], slides[3]],
      index: 1,
    })
    api.addImageBytes.mockResolvedValue({ slide: fresh, sourceId: 'sp1' })
    api.setLink.mockResolvedValue(fresh)
    const sctx = { ...ctx, sections: sectioned } as unknown as ActionCtx

    await insertSummaryZoom(sctx, [1, 3])

    expect(api.moveSlide).not.toHaveBeenCalled()
    expect(api.setSections).toHaveBeenCalledWith([
      { id: '{A}', name: 'Intro', slideIndices: [0] },
      { id: '', name: t('appSectionSummary'), slideIndices: [1] },
      { id: '', name: 'Agenda', slideIndices: [2] },
      { id: '{B}', name: 'Body', slideIndices: [3] },
      { id: '', name: t('appSectionN', { n: 5 }), slideIndices: [4] },
    ])
  })
})
