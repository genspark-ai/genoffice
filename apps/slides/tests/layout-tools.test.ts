/**
 * Tests for the layout tool trio (benchmarked against the Google Slides plugin approach):
 *  - runLayoutScript: sandbox-execute the AI layout script, collecting setBox operations
 *  - auditSlideLayout: deterministic geometry audit (overlap/out-of-bounds/text overflow)
 *  - execute_layout_script tool chain: script -> one applyEditScript transaction -> audit report back
 */
import { describe, it, expect, vi, beforeEach } from 'vitest'
import type { RenderSlide, RenderNode, ShapeRenderNode, PlacedBox } from '@genoffice/pptx-render'
import { runLayoutScript, type LayoutScriptElement } from '../src/renderer/ai/layout-script'
import { auditSlideLayout } from '../src/renderer/ai/layout-audit'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'

const box = (x: number, y: number, w: number, h: number, rot = 0): PlacedBox => ({
  x,
  y,
  w,
  h,
  rotationDeg: rot,
  flipH: false,
  flipV: false,
  centerX: x + w / 2,
  centerY: y + h / 2,
})

/** Build a shape node with text layout (contentHeight used for overflow detection). */
const textNode = (
  id: string,
  b: PlacedBox,
  text: string,
  contentHeight?: number,
): ShapeRenderNode => ({
  id,
  sourceId: id,
  type: 'shape',
  box: b,
  fill: { kind: 'none' },
  text: {
    lines: [
      {
        runs: [
          {
            text,
            x: 8,
            baselineY: 20,
            fontFamily: 'Arial',
            fontSizePx: 24,
            color: '#000000',
            bold: false,
            italic: false,
            underline: false,
            widthPx: text.length * 12,
          },
        ],
        top: 0,
        height: 28,
      },
    ],
    insets: { l: 8, t: 4, r: 8, b: 4 },
    anchor: 'top',
    fontScale: 1,
    wrap: true,
    contentHeight: contentHeight ?? 28,
  },
})

const slideOf = (nodes: RenderNode[], w = 1280, h = 720): RenderSlide => ({
  widthPx: w,
  heightPx: h,
  scale: 1,
  background: { kind: 'solid', color: '#FFFFFF' },
  nodes,
})

// ── runLayoutScript ──────────────────────────────────────────

describe('runLayoutScript', () => {
  const els: LayoutScriptElement[] = [
    { id: 'a', type: 'shape', text: 'Card A', x: 100, y: 200, w: 300, h: 320, rotation: 0 },
    { id: 'b', type: 'shape', text: 'Card B', x: 480, y: 210, w: 310, h: 300, rotation: 0 },
    { id: 'c', type: 'shape', text: 'Card C', x: 900, y: 190, w: 280, h: 330, rotation: 0 },
    {
      id: 'child',
      type: 'text',
      text: 'In group',
      x: 10,
      y: 10,
      w: 50,
      h: 20,
      rotation: 0,
      inGroup: true,
    },
  ]
  const canvas = { w: 1280, h: 720 }

  it('computes evenly spaced arrangement from real geometry', () => {
    const code = `
      const cards = els.filter(e => /^[abc]$/.test(e.id));
      const margin = 80, gap = 30;
      const w = (canvas.w - 2*margin - (cards.length-1)*gap) / cards.length;
      cards.forEach((c, i) => setBox(c.id, { x: margin + i*(w+gap), y: 200, w, h: 320 }));
      return cards.length + ' cards';
    `
    const r = runLayoutScript(code, els, canvas)
    expect(r.error).toBeUndefined()
    expect(r.returned).toBe('3 cards')
    expect(r.ops).toHaveLength(3)
    const a = r.ops.find((o) => o.id === 'a')!
    const b = r.ops.find((o) => o.id === 'b')!
    // (1280 - 160 - 60) / 3 ≈ 353.33
    expect(a.x).toBe(80)
    expect(b.x).toBeCloseTo(80 + a.w + 30, 5)
    expect(a.h).toBe(320) // patch took effect
    expect(a.rotation).toBe(0) // unspecified fields kept
  })

  it('setBox errors on unknown id', () => {
    const r = runLayoutScript(`setBox('nope', { x: 0 })`, els, canvas)
    expect(r.error).toContain('nope')
    expect(r.ops).toHaveLength(0)
  })

  it('setBox rejects elements inside a group', () => {
    const r = runLayoutScript(`setBox('child', { x: 0 })`, els, canvas)
    expect(r.error).toContain('group')
  })

  it('multiple setBox calls on the same element merge, later ones win', () => {
    const r = runLayoutScript(`setBox('a', { x: 10 }); setBox('a', { y: 20 });`, els, canvas)
    expect(r.ops).toHaveLength(1)
    expect(r.ops[0]).toMatchObject({ x: 10, y: 20, w: 300 })
  })

  it('forbids sandbox-escape primitives', () => {
    const r = runLayoutScript(`const f = [].constructor; setBox('a', {x:0})`, els, canvas)
    expect(r.error).toContain('constructor')
  })

  it('environment globals are shadowed to undefined', () => {
    const r = runLayoutScript(`return typeof fetch;`, els, canvas)
    expect(r.returned).toBe('undefined')
  })

  it('mutating els inside the script does not affect the input', () => {
    runLayoutScript(`els[0].x = 99999;`, els, canvas)
    expect(els[0]!.x).toBe(100)
  })

  it('log output is collected', () => {
    const r = runLayoutScript(`log('hello', {n: 1});`, els, canvas)
    expect(r.logs[0]).toBe('hello {"n":1}')
  })
})

// ── auditSlideLayout ─────────────────────────────────────────

describe('auditSlideLayout', () => {
  it('clean layout passes', () => {
    const slide = slideOf([
      textNode('t1', box(80, 60, 500, 80), 'Title'),
      textNode('t2', box(80, 200, 500, 200), 'Body'),
    ])
    expect(auditSlideLayout(slide)).toHaveLength(0)
  })

  it('detects overlapping text', () => {
    const slide = slideOf([
      textNode('t1', box(100, 100, 400, 200), 'Title One'),
      textNode('t2', box(200, 150, 400, 200), 'Title Two'),
    ])
    const issues = auditSlideLayout(slide)
    expect(
      issues.some(
        (s) => s.includes('Overlap') && s.includes('Title One') && s.includes('Title Two'),
      ),
    ).toBe(true)
  })

  it('detects out-of-bounds elements', () => {
    const slide = slideOf([textNode('t1', box(1100, 60, 400, 80), 'Overwide title')])
    const issues = auditSlideLayout(slide)
    expect(
      issues.some((s) => s.includes('Out of bounds') && s.includes('past the right edge')),
    ).toBe(true)
  })

  it('detects text overflow (contentHeight > box inner height)', () => {
    // Box height 100, insets 4 top and bottom -> inner height 92; content height 160 -> overflow 68px
    const slide = slideOf([textNode('t1', box(80, 60, 400, 100), 'Lots of text', 160)])
    const issues = auditSlideLayout(slide)
    expect(issues.some((s) => s.includes('Text overflow'))).toBe(true)
  })

  it('decoration layers and large background blocks are excluded', () => {
    const bg = textNode('bg', box(0, 0, 1280, 720), 'Background watermark')
    const deco = { ...textNode('deco', box(100, 100, 400, 200), 'Decoration'), decoration: true }
    const t = textNode('t1', box(150, 150, 400, 200), 'Body')
    const issues = auditSlideLayout(slideOf([bg, deco, t]))
    expect(issues).toHaveLength(0)
  })
})

// ── execute_layout_script tool chain ──────────────────────────

describe('execute_layout_script tool', () => {
  let slide: RenderSlide
  let applied: RenderSlide | null

  const access = (): DeckAccess => ({
    getSlides: () => [slide],
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: (_i, updated) => {
      applied = updated
    },
    applyDeck: () => {},
    fitWidthPx: 1280,
  })

  beforeEach(() => {
    applied = null
    slide = slideOf([
      textNode('t1', box(100, 100, 400, 100), 'Title'),
      textNode('t2', box(120, 150, 400, 100), 'Subtitle'),
    ])
    // Simulate the main process's applyEditScript: update boxes per items and return a new RenderSlide
    ;(globalThis as any).window = {
      slidesApi: {
        applyEditScript: vi.fn(async (op: any) => {
          const nodes = slide.nodes.map((n) => {
            const item = op.boxes.find((b: any) => b.id === n.sourceId)
            if (!item) return n
            return { ...n, box: box(item.x, item.y, item.w, item.h, item.rotation) }
          })
          return { slide: { ...slide, nodes } }
        }),
      },
    }
  })

  it('script → atomic apply → audit passes', async () => {
    const skill = createSlidesSkill(access())
    const r = await skill.executeTool({
      id: '1',
      name: 'execute_layout_script',
      input: {
        slideIndex: 0,
        code: `
          const [t1, t2] = els;
          setBox(t1.id, { x: 80, y: 60 });
          setBox(t2.id, { x: 80, y: t1.h + 80 });
          return 'ok';
        `,
      },
    } as any)
    expect((r as any).isError).toBeFalsy()
    expect(r.mutated).toBe(true)
    expect(r.output).toContain('2 element(s)')
    expect(r.output).toContain('✅ Passed')
    expect(applied).not.toBeNull()
    const t2 = applied!.nodes.find((n) => n.sourceId === 't2')!
    expect(t2.box.y).toBe(180)
  })

  it('still overlapping after adjustment → audit reports the issue and requests a fix', async () => {
    const skill = createSlidesSkill(access())
    const r = await skill.executeTool({
      id: '2',
      name: 'execute_layout_script',
      input: { slideIndex: 0, code: `setBox('t2', { x: 110, y: 120 });` },
    } as any)
    expect(r.mutated).toBe(true)
    expect(r.output).toContain('⚠️')
    expect(r.output).toContain('Overlap')
  })

  it('script error → isError and nothing is applied', async () => {
    const skill = createSlidesSkill(access())
    const r = await skill.executeTool({
      id: '3',
      name: 'execute_layout_script',
      input: { slideIndex: 0, code: `setBox('missing', { x: 0 });` },
    } as any)
    expect((r as any).isError).toBe(true)
    expect(applied).toBeNull()
  })

  it('read_slide returns geometry info', async () => {
    const skill = createSlidesSkill(access())
    const r = await skill.executeTool({
      id: '4',
      name: 'read_slide',
      input: { slideIndex: 0 },
    } as any)
    expect(r.output).toContain('Canvas 1280×720px')
    expect(r.output).toContain('pos(100,100) size 400×100')
    expect(r.output).toContain('font 18pt') // 24px → 18pt
  })
})
