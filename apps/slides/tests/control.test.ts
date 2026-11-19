import { describe, expect, it, vi } from 'vitest'
import type { RenderNode, RenderSlide } from '@genoffice/pptx-render'
import { handleSlidesControl, type SlidesControlState } from '../src/renderer/control'

const node = (id: string, extra: Partial<RenderNode> = {}): RenderNode =>
  ({
    id: `n_${id}`,
    sourceId: `s_${id}`,
    durableId: `e_${id}`,
    type: 'shape',
    ...extra,
  }) as RenderNode

const slide = (nodes: RenderNode[]): RenderSlide => ({ nodes }) as RenderSlide

function state(overrides: Partial<SlidesControlState> = {}): SlidesControlState {
  return {
    slides: [
      slide([node('1'), node('deco', { decoration: true })]),
      slide([node('2'), { ...node('g'), type: 'group', children: [node('inner')] } as RenderNode]),
    ],
    path: '/deck.pptx',
    current: 0,
    selectedIds: [],
    setCurrent: vi.fn(),
    setSelectedIds: vi.fn(),
    clearEditing: vi.fn(),
    ...overrides,
  }
}

describe('slides control hook', () => {
  it('answers not_ready until a deck is loaded', () => {
    expect(handleSlidesControl({ cmd: 'selection' }, state({ slides: [] }))).toEqual({
      status: 'not_ready',
    })
  })

  it('moves to the slide and selects the element by durable id', () => {
    const s = state()
    const reply = handleSlidesControl(
      { cmd: 'goto', target: { kind: 'slide', slide: 1, el: 'e_2' } },
      s,
    )
    expect(reply).toEqual({ status: 'ok', result: { slide: 1, element: 'e_2', type: 'shape' } })
    expect(s.setCurrent).toHaveBeenCalledWith(1)
    expect(s.setSelectedIds).toHaveBeenCalledWith(['s_2'])
    expect(s.clearEditing).toHaveBeenCalled()
  })

  it('selects the enclosing top-level group for an element nested in a group', () => {
    const s = state()
    const reply = handleSlidesControl(
      { cmd: 'goto', target: { kind: 'slide', slide: 1, el: 'e_inner' } },
      s,
    )
    expect(reply).toMatchObject({ status: 'ok', result: { element: 'e_g', type: 'group' } })
    expect(s.setSelectedIds).toHaveBeenCalledWith(['s_g'])
  })

  it('reports the valid range and the selectable ids on a miss', () => {
    const s = state()
    expect(
      handleSlidesControl({ cmd: 'goto', target: { kind: 'slide', slide: 5 } }, s),
    ).toMatchObject({
      status: 'error',
      error: { reason: 'out_of_range', detail: { valid_range: '0-1' } },
    })
    expect(
      handleSlidesControl({ cmd: 'goto', target: { kind: 'slide', slide: 0, el: 'e_9' } }, s),
    ).toMatchObject({
      status: 'error',
      error: { reason: 'target_not_found', detail: { available: ['e_1'] } },
    })
    expect(s.setCurrent).not.toHaveBeenCalled()
  })

  it('reads a selection inside an entered group', () => {
    const s = state({ current: 1, selectedIds: ['s_inner'] })
    expect(handleSlidesControl({ cmd: 'selection' }, s)).toEqual({
      status: 'ok',
      result: { slide: 1, elements: ['e_inner'], types: ['shape'] },
    })
  })

  it('reads the selection back as durable ids', () => {
    const s = state({ current: 1, selectedIds: ['s_2'] })
    expect(handleSlidesControl({ cmd: 'selection' }, s)).toEqual({
      status: 'ok',
      result: { slide: 1, elements: ['e_2'], types: ['shape'] },
    })
  })
})
