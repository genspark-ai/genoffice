/** Figure-provenance gate: chart data and data-dense briefs must declare dataSource; 'search' requires a real web_search first. */
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

// Rich deck so the anti-scratch-build guard stays out of the way
const richDeck = {
  widthPx: 1280,
  heightPx: 720,
  nodes: [textNode('a', 'Title'), textNode('b', 'P1'), textNode('c', 'P2'), textNode('d', 'P3')],
} as unknown as RenderSlide

function mkAccess(extra: Record<string, unknown> = {}): DeckAccess {
  return {
    getSlides: () => [richDeck],
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: () => {},
    applyDeck: () => {},
    fitWidthPx: 1280,
    retryBackoffMs: 0,
    ...extra,
  } as unknown as DeckAccess
}

const chartCall = (extra: Record<string, unknown> = {}): AgentToolCall => ({
  id: 't',
  name: 'add_chart',
  input: {
    slideIndex: 0,
    kind: 'bar',
    categories: ['Q1', 'Q2'],
    series: [{ name: 'Sales', values: [12.5, 48.2] }],
    ...extra,
  },
})

beforeEach(() => {
  ;(window as any).slidesApi = {
    addChart: vi.fn(async () => ({ slide: richDeck, sourceId: 'c1' })),
    editChart: vi.fn(async () => ({ slide: richDeck })),
    webSearch: vi.fn(async () => ({ answer: '', results: [] })),
  }
})

describe('add_chart provenance gate', () => {
  it('refuses without dataSource and names the accepted values', async () => {
    const r = await createSlidesSkill(mkAccess()).executeTool!(chartCall())
    expect(r.isError).toBe(true)
    expect(r.output).toContain('dataSource')
    expect((window as any).slidesApi.addChart).not.toHaveBeenCalled()
  })

  it("refuses dataSource 'search' when no web_search ran in this conversation", async () => {
    const r = await createSlidesSkill(mkAccess()).executeTool!(chartCall({ dataSource: 'search' }))
    expect(r.isError).toBe(true)
    expect(r.output).toContain('web_search')
    expect((window as any).slidesApi.addChart).not.toHaveBeenCalled()
  })

  it("accepts dataSource 'search' after a real web_search", async () => {
    const skill = createSlidesSkill(mkAccess())
    await skill.executeTool!({ id: 's', name: 'web_search', input: { query: 'heytea stores' } })
    const r = await skill.executeTool!(chartCall({ dataSource: 'search' }))
    expect(r.isError).toBeUndefined()
    expect((window as any).slidesApi.addChart).toHaveBeenCalledOnce()
  })

  it("accepts 'sample' but forces disclosure in the output", async () => {
    const r = await createSlidesSkill(mkAccess()).executeTool!(chartCall({ dataSource: 'sample' }))
    expect(r.isError).toBeUndefined()
    expect(r.output).toContain('illustrative')
  })

  it("accepts 'user' with no disclosure note", async () => {
    const r = await createSlidesSkill(mkAccess()).executeTool!(chartCall({ dataSource: 'user' }))
    expect(r.isError).toBeUndefined()
    expect(r.output).not.toContain('illustrative')
  })
})

describe('edit_chart provenance gate', () => {
  it('gates only calls that pass series data', async () => {
    const skill = createSlidesSkill(mkAccess())
    const withData = await skill.executeTool!({
      id: 't',
      name: 'edit_chart',
      input: { slideIndex: 0, sourceId: 'c1', series: [{ name: 's', values: [1.5, 2.5] }] },
    })
    expect(withData.isError).toBe(true)
    expect(withData.output).toContain('dataSource')

    const styleOnly = await skill.executeTool!({
      id: 't',
      name: 'edit_chart',
      input: { slideIndex: 0, sourceId: 'c1', title: 'New title' },
    })
    expect(styleOnly.isError).toBeUndefined()
  })
})

describe('brief provenance gate (regenerate_slide / generate_deck)', () => {
  const cloudAccess = () =>
    mkAccess({
      regenerateSlide: async () => null,
      generatePageCloud: async () => ({ ok: false, error: 'cloud down' }),
      isCloudPageGenEnabled: async () => true,
      landGeneratedPages: async () => ({ ok: true, pages: 1 }),
    })

  it('regenerate_slide with a figure-dense brief refuses without dataSource', async () => {
    const r = await createSlidesSkill(cloudAccess()).executeTool!({
      id: 't',
      name: 'regenerate_slide',
      input: { slideIndex: 0, brief: '2023 年营收 48 亿，同比增长 12.5%，客单价 ¥21.8' },
    })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('dataSource')
  })

  it('regenerate_slide with a figure-free brief is not gated', async () => {
    const r = await createSlidesSkill(cloudAccess()).executeTool!({
      id: 't',
      name: 'regenerate_slide',
      input: { slideIndex: 0, brief: 'Redo this page as a three column card layout' },
    })
    // Fails later in the cloud pipeline (mocked down), not at the provenance gate
    expect(r.output).not.toContain('dataSource')
  })

  it('generate_deck with figure-dense briefs refuses without dataSource', async () => {
    const r = await createSlidesSkill(cloudAccess()).executeTool!({
      id: 't',
      name: 'generate_deck',
      input: {
        core_hook: 'h',
        style: 's',
        pages: [{ title: 'Key metrics', brief: 'Revenue ¥4.8bn, growth 12.5%, ARPU $21.8' }],
      },
    })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('dataSource')
  })
})
