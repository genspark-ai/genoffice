/** Hard guard against "building from scratch by hand": apply_ops insert ops (addElement/addSmartArt) on an empty deck are refused and redirected to generate_deck. */
import { describe, it, expect, vi, beforeEach } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import type { RenderSlide, PlacedBox, ShapeRenderNode } from '@genoffice/pptx-render'
import type { AgentToolCall } from '../src/shared/ipc'

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
const textNode = (id: string, text: string): ShapeRenderNode =>
  ({
    id,
    sourceId: id,
    type: 'shape',
    box: box(60, 60, 400, 80),
    fill: { kind: 'none' },
    text: {
      lines: [
        {
          runs: [
            {
              text,
              x: 0,
              baselineY: 20,
              fontFamily: 'Arial',
              fontSizePx: 24,
              color: '#000',
              bold: false,
              italic: false,
              underline: false,
              widthPx: 100,
            },
          ],
          top: 0,
          height: 28,
        },
      ],
      insets: { l: 0, t: 0, r: 0, b: 0 },
      anchor: 'top',
      fontScale: 1,
      contentHeight: 28,
    },
  }) as unknown as ShapeRenderNode

// Blank deck (1 empty page, no text content) -> from-scratch scenario
const blankDeck = { widthPx: 1280, heightPx: 720, nodes: [] } as unknown as RenderSlide
// Existing polished deck (multiple elements with text) -> refinement scenario
const richDeck = {
  widthPx: 1280,
  heightPx: 720,
  nodes: [
    textNode('a', 'Title'),
    textNode('b', 'Point 1'),
    textNode('c', 'Point 2'),
    textNode('d', 'Point 3'),
  ],
} as unknown as RenderSlide

function mkAccess(slides: RenderSlide[]): DeckAccess {
  return {
    getSlides: () => slides,
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: () => {},
    applyDeck: () => {},
    fitWidthPx: 1280,
    landGeneratedPages: async () => ({ ok: true, pages: 1 }),
  } as unknown as DeckAccess
}
const insert = (op: 'addElement' | 'addSmartArt'): AgentToolCall => ({
  id: 't',
  name: 'apply_ops',
  input: {
    ops: [
      op === 'addElement'
        ? {
            op,
            target: { slide: 0 },
            kind: 'rect',
            offset: { x: 10, y: 10, cx: 100, cy: 50 },
            paragraphs: [{ runs: [{ text: 'x' }] }],
          }
        : {
            op,
            target: { slide: 0 },
            layout: 'list',
            items: ['a'],
            offset: { x: 10, y: 10, cx: 100, cy: 50 },
          },
    ],
  },
})

beforeEach(() => {
  ;(window as any).slidesApi = {
    applyTxn: vi.fn(async () => ({
      applied: true,
      records: [{ op: 'addElement', target: '0', created: ['e1'] }],
      slides: [richDeck],
    })),
  }
})

describe('anti hand-building from scratch', () => {
  it('empty deck + addElement → refused with guidance toward generate_deck', async () => {
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!(insert('addElement'))
    expect(r.isError).toBe(true)
    expect(r.output).toContain('ops[0] addElement')
    expect(r.output).toContain('generate_deck')
    expect((window as any).slidesApi.applyTxn).not.toHaveBeenCalled()
  })
  it('empty deck + addSmartArt → refused', async () => {
    const r = await createSlidesSkill(mkAccess([blankDeck])).executeTool!(insert('addSmartArt'))
    expect(r.isError).toBe(true)
    expect((window as any).slidesApi.applyTxn).not.toHaveBeenCalled()
  })
  it('existing rich deck (lots of content) + addElement → allowed (fine-tuning is legitimate)', async () => {
    const r = await createSlidesSkill(mkAccess([richDeck])).executeTool!(insert('addElement'))
    expect(r.isError).toBeUndefined()
    expect((window as any).slidesApi.applyTxn).toHaveBeenCalledOnce()
  })
  it('after cloud generation has run, allowed even with an empty deck (tweak scenario)', async () => {
    const access = {
      ...mkAccess([blankDeck]),
      retryBackoffMs: 0,
      isCloudPageGenEnabled: async () => true,
      generatePageCloud: async () => ({ ok: true, marker: 'cloudpptx:/tmp/x.pptx' }),
    } as unknown as DeckAccess
    const skill = createSlidesSkill(access)
    // First run one generate_deck to set htmlGenerated=true
    await skill.executeTool!({
      id: 't',
      name: 'generate_deck',
      input: {
        core_hook: 'h',
        style: 's',
        pages: [{ title: 'T', brief: 'b', layout: 'data', image_queries: [] }],
      },
    })
    const r = await skill.executeTool!(insert('addElement'))
    expect(r.isError).toBeUndefined()
  })
})
