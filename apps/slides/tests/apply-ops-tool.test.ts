/** apply_ops — the AI batch surface over the op transaction executor. */
import { describe, it, expect, vi, beforeEach } from 'vitest'
import type { RenderSlide, ShapeRenderNode } from '@genoffice/pptx-render'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'

const slide = { widthPx: 1280, heightPx: 720, scale: 1, nodes: [] } as unknown as RenderSlide

const textNode = (id: string, text: string, x = 60, w = 400, y = 60) =>
  ({
    id,
    sourceId: id,
    type: 'shape',
    box: {
      x,
      y,
      w,
      h: 80,
      rotationDeg: 0,
      flipH: false,
      flipV: false,
      centerX: x + w / 2,
      centerY: y + 40,
    },
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

/** A deck with real content: the anti-scratch-build guard stays out of the way */
const richSlide = {
  widthPx: 1280,
  heightPx: 720,
  scale: 1,
  // stacked with gaps so the layout audit has nothing to report
  nodes: [
    textNode('a', 'Title', 60, 400, 60),
    textNode('b', 'Point 1', 60, 400, 200),
    textNode('c', 'Point 2', 60, 400, 340),
  ],
} as unknown as RenderSlide

let deckApplied: RenderSlide[] | null
let deckGoTo: number | null
let current = 0

const access = (slides: RenderSlide[] = [richSlide]): DeckAccess =>
  ({
    getSlides: () => slides,
    getCurrent: () => current,
    getSelectedIds: () => [],
    applySlide: () => {},
    applyDeck: (all: RenderSlide[], goTo: number) => {
      deckApplied = all
      deckGoTo = goTo
    },
    fitWidthPx: 1280,
  }) as unknown as DeckAccess

const call = (input: Record<string, unknown>) => ({ id: 't', name: 'apply_ops', input }) as never

beforeEach(() => {
  deckApplied = null
  deckGoTo = null
  current = 0
  ;(globalThis as any).window = { slidesApi: {} }
})

describe('apply_ops', () => {
  it('dry run echoes the plan and rejected ops without mutating', async () => {
    ;(globalThis as any).window.slidesApi.applyTxn = vi.fn(async () => ({
      applied: false,
      dryRun: true,
      plan: ['[0] setFill s0/e_1'],
      failures: [{ index: 1, error: 'op "sparkle": unknown op' }],
    }))
    const r = await createSlidesSkill(access()).executeTool!(
      call({ ops: [{ op: 'setFill' }, { op: 'sparkle' }], dry_run: true }),
    )
    expect((r as any).isError).toBeFalsy()
    expect(r.mutated).toBe(false)
    expect(r.output).toContain('NOT modified')
    expect(r.output).toContain('[0] setFill s0/e_1')
    expect(r.output).toContain('ops[1]')
    expect(deckApplied).toBeNull()
  })

  it('atomic failure surfaces the guided error and applies nothing', async () => {
    ;(globalThis as any).window.slidesApi.applyTxn = vi.fn(async () => ({
      applied: false,
      failures: [
        { index: 0, error: 'op "setFill": no element "ghost" on slide 0. Available: [e_1]' },
      ],
    }))
    const r = await createSlidesSkill(access()).executeTool!(
      call({ ops: [{ op: 'setFill', target: { slide: 0, el: 'ghost' } }] }),
    )
    expect((r as any).isError).toBe(true)
    expect(r.output).toContain('Nothing was applied (atomic)')
    expect(r.output).toContain('Available: [e_1]')
    expect(deckApplied).toBeNull()
  })

  it('success applies the returned deck and reports the journal echo', async () => {
    ;(globalThis as any).window.slidesApi.applyTxn = vi.fn(async (req: any) => {
      expect(req.ops).toHaveLength(2)
      expect(req.dryRun).toBeUndefined()
      return {
        applied: true,
        records: [
          { op: 'setHidden', target: '0' },
          { op: 'addElement', target: '1', created: ['e_9'] },
        ],
        slides: [slide, slide],
      }
    })
    const r = await createSlidesSkill(access()).executeTool!(
      call({
        ops: [
          { op: 'setHidden', target: { slide: 0 }, hidden: true },
          {
            op: 'addElement',
            target: { slide: 1 },
            kind: 'rect',
            offset: { x: 0, y: 0, cx: 1, cy: 1 },
          },
        ],
      }),
    )
    expect((r as any).isError).toBeFalsy()
    expect(r.mutated).toBe(true)
    expect(r.output).toContain('Applied 2 op(s)')
    expect(r.output).toContain('setHidden @0')
    expect(r.output).toContain('New element ids: e_9')
    expect(deckApplied).toHaveLength(2)
  })

  it('clamps the current page index when the batch shrank the deck', async () => {
    current = 1 // the user is on page 2; the batch deletes it
    ;(globalThis as any).window.slidesApi.applyTxn = vi.fn(async () => ({
      applied: true,
      records: [{ op: 'deleteSlide', target: '1' }],
      slides: [slide],
    }))
    const r = await createSlidesSkill(access()).executeTool!(
      call({ ops: [{ op: 'deleteSlide', target: { slide: 1 } }] }),
    )
    expect((r as any).isError).toBeFalsy()
    expect(deckGoTo).toBe(0)
  })

  it('addChart without dataSource is refused before the IPC (figure provenance gate)', async () => {
    const spy = vi.fn()
    ;(globalThis as any).window.slidesApi.applyTxn = spy
    const r = await createSlidesSkill(access()).executeTool!(
      call({
        ops: [
          {
            op: 'addChart',
            target: { slide: 0 },
            kind: 'bar',
            categories: ['Q1'],
            series: [{ name: 'Revenue', values: [12.5] }],
            offset: { x: 0, y: 0, cx: 1, cy: 1 },
          },
        ],
      }),
    )
    expect((r as any).isError).toBe(true)
    expect(r.output).toContain('ops[0] addChart')
    expect(r.output).toContain('dataSource')
    expect(spy).not.toHaveBeenCalled()
  })

  it('strips dataSource before the IPC and discloses sample figures', async () => {
    let sent: any
    ;(globalThis as any).window.slidesApi.applyTxn = vi.fn(async (req: any) => {
      sent = req
      return {
        applied: true,
        records: [{ op: 'addChart', target: '0', created: ['e_5'] }],
        slides: [richSlide],
      }
    })
    const r = await createSlidesSkill(access()).executeTool!(
      call({
        ops: [
          {
            op: 'addChart',
            target: { slide: 0 },
            kind: 'pie',
            categories: ['A', 'B'],
            series: [{ name: 'Share', values: [60, 40] }],
            offset: { x: 0, y: 0, cx: 1, cy: 1 },
            dataSource: 'sample',
          },
        ],
      }),
    )
    expect((r as any).isError).toBeFalsy()
    expect(sent.ops[0]).not.toHaveProperty('dataSource')
    expect(sent.ops[0].kind).toBe('pie')
    expect(r.output).toContain('illustrative')
  })

  it('inserts on a blank deck are redirected to generate_deck (scratch guard)', async () => {
    const spy = vi.fn()
    ;(globalThis as any).window.slidesApi.applyTxn = spy
    const r = await createSlidesSkill(access([slide])).executeTool!(
      call({
        ops: [
          { op: 'setHidden', target: { slide: 0 }, hidden: true },
          {
            op: 'addElement',
            target: { slide: 0 },
            kind: 'rect',
            offset: { x: 0, y: 0, cx: 1, cy: 1 },
          },
        ],
      }),
    )
    expect((r as any).isError).toBe(true)
    expect(r.output).toContain('ops[1] addElement')
    expect(r.output).toContain('generate_deck')
    expect(spy).not.toHaveBeenCalled()
  })

  it('appends a layout audit for touched pages when they have issues', async () => {
    const broken = {
      ...richSlide,
      nodes: [...(richSlide as any).nodes, textNode('z', 'Runs off the page', 1100, 400, 480)],
    } as unknown as RenderSlide
    ;(globalThis as any).window.slidesApi.applyTxn = vi.fn(async () => ({
      applied: true,
      records: [
        { op: 'setTransform', target: '0/z' },
        { op: 'setHidden', target: '1' },
      ],
      slides: [broken, richSlide],
    }))
    const r = await createSlidesSkill(access([richSlide, richSlide])).executeTool!(
      call({
        ops: [
          { op: 'setTransform', target: { slide: 0, el: 'z' }, box: { x: 0, y: 0, cx: 1, cy: 1 } },
          { op: 'setHidden', target: { slide: 1 }, hidden: true },
        ],
      }),
    )
    expect((r as any).isError).toBeFalsy()
    expect(r.output).toContain('<layout-audit>')
    expect(r.output).toContain('page 1:')
    expect(r.output).toContain('Out of bounds')
    expect(r.output).not.toContain('page 2:')
  })

  it('omits the audit block when every touched page passes', async () => {
    ;(globalThis as any).window.slidesApi.applyTxn = vi.fn(async () => ({
      applied: true,
      records: [{ op: 'setHidden', target: '0' }],
      slides: [richSlide],
    }))
    const r = await createSlidesSkill(access()).executeTool!(
      call({ ops: [{ op: 'setHidden', target: { slide: 0 }, hidden: true }] }),
    )
    expect(r.output).not.toContain('<layout-audit>')
  })

  it('empty ops fails fast without touching the IPC', async () => {
    const spy = vi.fn()
    ;(globalThis as any).window.slidesApi.applyTxn = spy
    const r = await createSlidesSkill(access()).executeTool!(call({ ops: [] }))
    expect((r as any).isError).toBe(true)
    expect(spy).not.toHaveBeenCalled()
  })
})
