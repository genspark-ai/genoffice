/** apply_ops — the AI batch surface over the op transaction executor. */
import { describe, it, expect, vi, beforeEach } from 'vitest'
import type { RenderSlide } from '@genoffice/pptx-render'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'

const slide = { widthPx: 1280, heightPx: 720, scale: 1, nodes: [] } as unknown as RenderSlide

let deckApplied: RenderSlide[] | null
let deckGoTo: number | null
let current = 0

const access = (): DeckAccess =>
  ({
    getSlides: () => [slide],
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

  it('empty ops fails fast without touching the IPC', async () => {
    const spy = vi.fn()
    ;(globalThis as any).window.slidesApi.applyTxn = spy
    const r = await createSlidesSkill(access()).executeTool!(call({ ops: [] }))
    expect((r as any).isError).toBe(true)
    expect(spy).not.toHaveBeenCalled()
  })
})
