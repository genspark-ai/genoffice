/**
 * AI color readback: read_slide lines carry fill/text/stroke hex colors, the deck outline
 * summarizes each page's main fills, and layout-script els expose the same read-only fields.
 */
import { describe, it, expect } from 'vitest'
import type {
  RenderSlide,
  RenderNode,
  ShapeRenderNode,
  PictureRenderNode,
  PlacedBox,
} from '@genoffice/pptx-render'
import { runLayoutScript, type LayoutScriptElement } from '../src/renderer/ai/layout-script'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'

const box = (x: number, y: number, w: number, h: number): PlacedBox => ({
  x,
  y,
  w,
  h,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
  centerX: x + w / 2,
  centerY: y + h / 2,
})

const shape = (
  id: string,
  b: PlacedBox,
  opts: {
    fill?: ShapeRenderNode['fill']
    stroke?: ShapeRenderNode['stroke']
    runs?: Array<{ text: string; color: string }>
  } = {},
): ShapeRenderNode => ({
  id,
  sourceId: id,
  type: 'shape',
  box: b,
  fill: opts.fill ?? { kind: 'none' },
  ...(opts.stroke ? { stroke: opts.stroke } : {}),
  ...(opts.runs
    ? {
        text: {
          lines: [
            {
              runs: opts.runs.map((r, i) => ({
                text: r.text,
                x: 8 + i * 40,
                baselineY: 20,
                fontFamily: 'Arial',
                fontSizePx: 24,
                color: r.color,
                bold: false,
                italic: false,
                underline: false,
                widthPx: r.text.length * 12,
              })),
              top: 0,
              height: 28,
            },
          ],
          insets: { l: 8, t: 4, r: 8, b: 4 },
          anchor: 'top' as const,
          fontScale: 1,
          wrap: true,
          contentHeight: 28,
        },
      }
    : {}),
})

const picture = (id: string, b: PlacedBox): PictureRenderNode => ({
  id,
  sourceId: id,
  type: 'picture',
  box: b,
})

const slideOf = (nodes: RenderNode[]): RenderSlide => ({
  widthPx: 1280,
  heightPx: 720,
  scale: 1,
  background: { kind: 'solid', color: '#FFFFFF' },
  nodes,
})

const access = (slide: RenderSlide, selectedIds: string[] = []): DeckAccess => ({
  getSlides: () => [slide],
  getCurrent: () => 0,
  getSelectedIds: () => selectedIds,
  applySlide: () => {},
  applyDeck: () => {},
  fitWidthPx: 1280,
})

describe('read_slide colors', () => {
  it('output lines carry fill/text/stroke hex colors', async () => {
    const slide = slideOf([
      shape('card', box(100, 100, 300, 200), {
        fill: { kind: 'solid', color: '#4472C4' },
        stroke: { color: '#333333', widthPx: 1, widthPt: 0.75 },
        runs: [{ text: 'Title', color: '#FFFFFF' }],
      }),
    ])
    const skill = createSlidesSkill(access(slide))
    const r = await skill.executeTool({ id: '1', name: 'read_slide', input: { slideIndex: 0 } })
    expect(r.output).toContain('fill#4472C4')
    expect(r.output).toContain('text#FFFFFF')
    expect(r.output).toContain('stroke#333333')
  })

  it('dominant text color chosen by character-count majority, alpha colors normalized to 6 digits', async () => {
    const slide = slideOf([
      shape('t', box(0, 0, 400, 100), {
        runs: [
          { text: 'a much longer body text run', color: '#112233cc' },
          { text: 'x', color: '#FF0000' },
        ],
      }),
    ])
    const skill = createSlidesSkill(access(slide))
    const r = await skill.executeTool({ id: '1', name: 'read_slide', input: { slideIndex: 0 } })
    expect(r.output).toContain('text#112233')
    expect(r.output).not.toContain('#112233CC')
  })

  it('gradient fill yields no fill field; image elements declare unreadable color at the tail', async () => {
    const slide = slideOf([
      shape('g', box(0, 0, 200, 100), {
        fill: { kind: 'gradient', stops: [{ pos: 0, color: '#AA0000' }], angleDeg: 0 },
      }),
      picture('pic', box(300, 0, 200, 100)),
    ])
    const skill = createSlidesSkill(access(slide))
    const r = await skill.executeTool({ id: '1', name: 'read_slide', input: { slideIndex: 0 } })
    expect(r.output).not.toContain('fill#AA0000')
    expect(r.output).toContain('picture colors not available')
  })
})

describe('deck outline main fills', () => {
  it('tallies main fill colors by frequency', () => {
    const slide = slideOf([
      shape('a', box(0, 0, 100, 100), { fill: { kind: 'solid', color: '#4472C4' } }),
      shape('b', box(120, 0, 100, 100), { fill: { kind: 'solid', color: '#4472C4' } }),
      shape('c', box(240, 0, 100, 100), { fill: { kind: 'solid', color: '#F5F5F5' } }),
      shape('d', box(360, 0, 100, 100)),
    ])
    const skill = createSlidesSkill(access(slide))
    const outline = skill.buildContext!()
    expect(outline).toContain('main fills: #4472C4×2 #F5F5F5')
  })

  it('omits the main fills line when there are no readable fill colors', () => {
    const slide = slideOf([picture('pic', box(0, 0, 100, 100))])
    const skill = createSlidesSkill(access(slide))
    expect(skill.buildContext!()).not.toContain('main fills')
  })

  it('reports a canvas source-id selection with the durable id shown in the outline', () => {
    const title = {
      ...shape('source-title', box(0, 0, 500, 100), {
        runs: [{ text: 'Main title', color: '#FFFFFF' }],
      }),
      durableId: 'e_title',
    }
    const subtitle = {
      ...shape('source-subtitle', box(0, 120, 500, 60), {
        runs: [{ text: 'Subtitle below', color: '#FFFFFF' }],
      }),
      durableId: 'e_subtitle',
    }
    const outline = createSlidesSkill(access(slideOf([title, subtitle]), ['source-title']))
      .buildContext!()
    expect(outline).toContain('User selected elements: e_title')
    expect(outline).not.toContain('User selected elements: source-title')
    expect(outline).toContain('- e_title | shape | "Main title"')
    expect(outline).toContain('- e_subtitle | shape | "Subtitle below"')
  })
})

describe('layout-script els colors', () => {
  it('scripts can read fill/textColor/strokeColor fields', () => {
    const els: LayoutScriptElement[] = [
      {
        id: 'a',
        type: 'shape',
        text: 'Card',
        x: 0,
        y: 0,
        w: 100,
        h: 100,
        rotation: 0,
        fill: '#4472C4',
        textColor: '#FFFFFF',
        strokeColor: '#333333',
      },
    ]
    const r = runLayoutScript(
      `const e = els[0]; return e.fill + '/' + e.textColor + '/' + e.strokeColor;`,
      els,
      { w: 1280, h: 720 },
    )
    expect(r.error).toBeUndefined()
    expect(r.returned).toBe('#4472C4/#FFFFFF/#333333')
  })
})
