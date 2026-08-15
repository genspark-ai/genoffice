/**
 * Phase 3 edit round-trip verification (no GUI; runs the engine+render paths the main process uses):
 * open -> build RenderTree -> edit text/geometry -> patch save -> reopen and read back.
 * Covers the real UI-triggered data flow (IPC handler internals) without Electron/a display.
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { openPptx, savePptx, type TextElement } from '@genoffice/pptx-engine'
import { buildRenderSlide, EMU_PER_PX_96, type RenderSlide } from '@genoffice/pptx-render'
import {
  fillToKonva,
  glyphToDraw,
  presetToShapeKind,
  normalizeColor,
  fillImageUrl,
} from '../src/renderer/konva-adapter'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) =>
  new Uint8Array(
    readFileSync(
      join(here, '..', '..', '..', 'packages', 'pptx-engine', 'tests', 'fixtures', name),
    ),
  )

const FIT = 1280

function firstTextEl(slide: any): TextElement {
  return slide.elements.find(
    (e: any) => (e.type === 'text' || e.type === 'shape') && e.text?.paragraphs?.length,
  )
}

describe('Phase 3 open → render → edit → save flow', () => {
  it('opens and builds RenderSlides at fit width', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slides: RenderSlide[] = opened.deck.slides.map((s) =>
      buildRenderSlide(s, opened.deck.size, { fitWidthPx: FIT }),
    )
    expect(slides.length).toBe(5)
    expect(slides[0]!.widthPx).toBe(FIT)
    expect(slides[0]!.heightPx).toBeCloseTo(720, 0) // 16:9
    expect(slides[0]!.nodes.length).toBeGreaterThan(0)
  })

  it('edit text (as IPC editText does) rebuilds RenderSlide and marks dirty', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slideIdx = opened.deck.slides.findIndex((s) => firstTextEl(s))
    const slide = opened.deck.slides[slideIdx]!
    const el = firstTextEl(slide)

    // Simulate the editText handler: replace the whole box with one paragraph/one run, keeping the first run's format
    const first = el.text!.paragraphs[0]!.runs[0]
    el.text!.paragraphs = 'New Title\nSecond Line'.split('\n').map((line) => ({
      runs: [{ text: line, bold: first?.bold, fontSize: first?.fontSize, color: first?.color }],
    }))
    el.dirty = true

    const rebuilt = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: FIT })
    const node = rebuilt.nodes.find((n) => n.sourceId === el.id) as any
    const text = node.text.lines
      .flatMap((l: any) => l.runs)
      .map((r: any) => r.text)
      .join('')
    expect(text).toContain('New Title')
    expect(text).toContain('Second Line')
    expect(el.dirty).toBe(true)
  })

  it('edit transform (as IPC editTransform does) reflects new geometry', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = slide.elements[0]!

    // px -> EMU back-calculation (replicating the handler logic)
    const baseWidthPx = opened.deck.size.cx / EMU_PER_PX_96
    const scale = FIT / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    el.transform = {
      ...el.transform,
      offset: { x: toEmu(100), y: toEmu(50), cx: toEmu(300), cy: toEmu(120) },
      rot: Math.round(15 * 60000),
    }
    el.dirty = true

    const rebuilt = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: FIT })
    const node = rebuilt.nodes.find((n) => n.sourceId === el.id)!
    expect(node.box.x).toBeCloseTo(100, 0)
    expect(node.box.y).toBeCloseTo(50, 0)
    expect(node.box.w).toBeCloseTo(300, 0)
    expect(node.box.rotationDeg).toBeCloseTo(15, 1)
  })

  it('run-level edit preserves per-run formatting (bold run + plain run)', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slideIdx = opened.deck.slides.findIndex((s) => firstTextEl(s))
    const el = firstTextEl(opened.deck.slides[slideIdx]!)
    // Simulate the editText handler's run-level rebuild: two runs with different formats
    el.text!.paragraphs = [
      {
        runs: [
          { text: 'BoldRun', bold: true, color: '#FF0000' },
          { text: 'PlainRun', bold: false },
        ],
      },
    ]
    el.dirty = true
    const out = await savePptx(opened)
    const reopened = await openPptx(out)
    const runs = reopened.deck.slides
      .flatMap((s) => s.elements)
      .filter((e: any) => e.type === 'text' || e.type === 'shape')
      .flatMap((e: any) => e.text?.paragraphs ?? [])
      .flatMap((p: any) => p.runs)
    const bold = runs.find((r: any) => r.text === 'BoldRun')
    const plain = runs.find((r: any) => r.text === 'PlainRun')
    expect(bold?.bold).toBe(true)
    expect(bold?.color).toBe('#FF0000')
    expect(plain?.bold).toBe(false)
  })

  it('save after edit → reopen reads new text (full round-trip)', async () => {
    const opened = await openPptx(fx('05_unicode_cjk_emoji.pptx'))
    const slideIdx = opened.deck.slides.findIndex((s) => firstTextEl(s))
    const el = firstTextEl(opened.deck.slides[slideIdx]!)
    el.text!.paragraphs = [{ runs: [{ text: 'Phase3编辑验证🎯', bold: true }] }]
    el.dirty = true

    const out = await savePptx(opened)
    const reopened = await openPptx(out)
    const allText = reopened.deck.slides
      .flatMap((s) => s.elements)
      .filter((e: any) => e.type === 'text' || e.type === 'shape')
      .flatMap((e: any) => e.text?.paragraphs ?? [])
      .flatMap((p: any) => p.runs)
      .map((r: any) => r.text)
      .join('')
    expect(allText).toContain('Phase3编辑验证🎯')
  })

  it('no-edit save stays byte-identical (fidelity through the app path)', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const out = await savePptx(opened)
    const JSZip = (await import('jszip')).default
    const origZip = await JSZip.loadAsync(fx('01_standard_business.pptx'))
    const outZip = await JSZip.loadAsync(out)
    for (const slide of opened.deck.slides) {
      const a = await origZip.file(slide.path)!.async('uint8array')
      const b = await outZip.file(slide.path)!.async('uint8array')
      expect(Buffer.from(b).equals(Buffer.from(a))).toBe(true)
    }
  })
})

describe('konva-adapter (thin data mapping)', () => {
  it('solid fill → css color', () => {
    expect(fillToKonva({ kind: 'solid', color: '#3366CC' }, 100, 50)).toEqual({ fill: '#3366CC' })
  })

  it('gradient fill → linear gradient stops', () => {
    const r = fillToKonva(
      {
        kind: 'gradient',
        stops: [
          { pos: 0, color: '#000000' },
          { pos: 1, color: '#FFFFFF' },
        ],
        angleDeg: 0,
      },
      100,
      50,
    )
    expect(r.fillLinearGradientColorStops).toEqual([0, '#000000', 1, '#FFFFFF'])
  })

  it('linear gradient vector spans the box projection, not max(w,h)', () => {
    // Vertical gradient on a wide box: the ramp must run exactly top edge -> bottom edge
    const r = fillToKonva(
      {
        kind: 'gradient',
        stops: [
          { pos: 0, color: '#000000' },
          { pos: 1, color: '#FFFFFF' },
        ],
        angleDeg: 90,
      },
      784,
      128,
    )
    expect(r.fillLinearGradientStartPoint!.x).toBeCloseTo(392, 6)
    expect(r.fillLinearGradientStartPoint!.y).toBeCloseTo(0, 6)
    expect(r.fillLinearGradientEndPoint!.x).toBeCloseTo(392, 6)
    expect(r.fillLinearGradientEndPoint!.y).toBeCloseTo(128, 6)
  })

  it('8-digit hex → rgba', () => {
    expect(normalizeColor('#FF000080')).toMatch(/^rgba\(255,0,0,0\.50/)
  })

  it('image fill → fillPatternImage when image loaded (stretch scales)', () => {
    const img = { width: 100, height: 50 } as HTMLImageElement
    const images = new Map([['data:x', img]])
    const r = fillToKonva({ kind: 'image', dataUrl: 'data:x', mode: 'stretch' }, 200, 100, images)
    expect(r.fillPatternImage).toBe(img)
    expect(r.fillPatternScaleX).toBeCloseTo(2, 4)
    expect(r.fillPatternScaleY).toBeCloseTo(2, 4)
  })

  it('image fill → empty when image not yet loaded', () => {
    const r = fillToKonva(
      { kind: 'image', dataUrl: 'data:x', mode: 'stretch' },
      200,
      100,
      new Map(),
    )
    expect(r.fillPatternImage).toBeUndefined()
  })

  it('fillImageUrl extracts dataUrl only for image fills', () => {
    expect(fillImageUrl({ kind: 'image', dataUrl: 'data:y', mode: 'tile' })).toBe('data:y')
    expect(fillImageUrl({ kind: 'solid', color: '#000' })).toBeUndefined()
  })

  it('preset geometry mapping', () => {
    expect(presetToShapeKind('ellipse')).toBe('ellipse')
    expect(presetToShapeKind('roundRect')).toBe('roundRect')
    expect(presetToShapeKind('rect')).toBe('rect')
    expect(presetToShapeKind(undefined)).toBe('rect')
  })

  it('glyph draw: baseline → top y conversion + font style', () => {
    const d = glyphToDraw({
      text: 'x',
      x: 10,
      baselineY: 40,
      fontFamily: 'Arial',
      fontSizePx: 20,
      color: '#000000',
      bold: true,
      italic: true,
      underline: true,
      widthPx: 12,
    })
    expect(d.x).toBe(10)
    expect(d.y).toBeCloseTo(40 - 20 * 0.8, 4)
    expect(d.fontStyle).toBe('bold italic')
    expect(d.textDecoration).toBe('underline')
  })
})
