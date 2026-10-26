/**
 * Deck-picture vision injection on the AiPanel send path: vision-capable
 * settings make the panel collect the deck's pictures and pass them to
 * loop.run after the attachment images (outline then annotates "image #K",
 * covered in deck-images.test.ts); text-only settings leave the call
 * untouched — no collection, no IPC, zero behavior change.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'

const agentHarness = vi.hoisted(() => ({
  events: null as null | {
    onDone?: (result: { text: string; cancelled: boolean; turnLimit: boolean }) => void
  },
  run: vi.fn(),
  cancel: vi.fn(),
  reset: vi.fn(),
  restore: vi.fn(),
}))

vi.mock('@genoffice/agent-core', async () => {
  const actual =
    await vi.importActual<typeof import('@genoffice/agent-core')>('@genoffice/agent-core')
  return {
    ...actual,
    AgentLoop: class MockAgentLoop {
      busy = false

      constructor(options: { events?: typeof agentHarness.events }) {
        agentHarness.events = options.events ?? null
      }

      run(...args: unknown[]) {
        agentHarness.run(...args)
      }

      cancel() {
        agentHarness.cancel()
      }

      reset() {
        agentHarness.reset()
      }

      restore(...args: unknown[]) {
        agentHarness.restore(...args)
      }
    },
  }
})

// react-konva's node entry requires the native 'canvas' package; these tests do not draw.
vi.mock('react-konva', () => {
  const stub = () => null
  return {
    Stage: stub,
    Layer: stub,
    Rect: stub,
    Group: stub,
    Transformer: stub,
    Line: stub,
    Arrow: stub,
    Text: stub,
    Ellipse: stub,
    Image: stub,
    Path: stub,
    Circle: stub,
    Arc: stub,
  }
})

import { AiPanel } from '../src/renderer/ai/AiPanel'
import {
  createSlidesSkill,
  setDeckImageAttachment,
  type DeckAccess,
} from '../src/renderer/ai/slides-skill'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'
import type { RenderSlide } from '@genoffice/pptx-render'

/** Vision-capable: anthropic's default model takes images */
const visionSettings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((provider) => [provider.id, { apiKey: '', model: provider.defaultModel }]),
  ) as AiSettings['providers'],
}

/** Text-only: glm's provider capability is vision:false, whatever the model */
const textOnlySettings: AiSettings = {
  ...visionSettings,
  provider: 'glm',
}

const picSlide = {
  widthPx: 1280,
  heightPx: 720,
  scale: 1,
  background: { kind: 'solid', color: '#FFFFFF' },
  nodes: [
    {
      id: 'r_pic1',
      sourceId: 'pic1',
      type: 'picture',
      box: { x: 0, y: 0, w: 100, h: 100, rotationDeg: 0, flipH: false, flipV: false, centerX: 50, centerY: 50 },
      dataUrl: 'data:image/png;base64,QUJD',
    },
    {
      id: 'r_pic2',
      sourceId: 'pic2',
      type: 'picture',
      box: { x: 0, y: 200, w: 100, h: 100, rotationDeg: 0, flipH: false, flipV: false, centerX: 50, centerY: 250 },
      dataUrl: 'data:image/gif;base64,R0w=',
    },
  ],
} as unknown as RenderSlide

const mountedRoots: Array<{ root: Root; container: HTMLElement }> = []

function mount(element: React.ReactElement): { root: Root; container: HTMLElement } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  mountedRoots.push({ root, container })
  return { root, container }
}

function panelProps(overrides: Record<string, unknown> = {}) {
  return {
    slides: [],
    current: 0,
    selectedIds: [],
    images: new Map<string, HTMLImageElement>(),
    applySlide: () => {},
    applyDeck: () => {},
    fitWidthPx: 960,
    settings: visionSettings,
    open: true,
    onExpand: () => {},
    onCollapse: () => {},
    ...overrides,
  }
}

async function flushEffects(): Promise<void> {
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
    await Promise.resolve()
  })
}

function installSlidesApi(readDeckImage: (op: { slideIndex: number; elementId: string }) => Promise<unknown>) {
  Object.defineProperty(window, 'slidesApi', {
    configurable: true,
    value: {
      aiGskStatus: vi.fn(async () => ({ loggedIn: true })),
      beginHistoryBatch: vi.fn(async () => false),
      endHistoryBatch: vi.fn(async () => null),
      aiLogRunFailure: vi.fn(async () => undefined),
      readDeckImage,
    },
  })
}

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

beforeEach(() => {
  agentHarness.events = null
  agentHarness.run.mockReset()
  agentHarness.cancel.mockReset()
  agentHarness.reset.mockReset()
  agentHarness.restore.mockReset()
  setDeckImageAttachment(null)
  Object.defineProperty(window, 'projectApi', { configurable: true, value: undefined })
  Object.defineProperty(window, 'desktop', {
    configurable: true,
    value: { readAttachmentImage: vi.fn(async () => ({ ok: false })) },
  })
})

afterEach(() => {
  for (const { root, container } of mountedRoots.splice(0)) {
    act(() => root.unmount())
    container.remove()
  }
  vi.restoreAllMocks()
})

describe('AiPanel deck-picture vision injection (slides)', () => {
  it('vision on: collects deck pictures and sends them after attachment images', async () => {
    const readDeckImage = vi.fn(async () => ({ ok: false, error: 'not found' }))
    installSlidesApi(readDeckImage)
    Object.defineProperty(window, 'desktop', {
      configurable: true,
      value: {
        readAttachmentImage: vi.fn(async () => ({ ok: true, base64: 'AAAA', mime: 'image/png' })),
      },
    })
    const attachments = [
      { path: '/tmp/a.png', name: 'a.png', ext: 'png', sizeBytes: 128 },
    ]

    mount(
      createElement(
        AiPanel,
        panelProps({
          slides: [picSlide],
          preset: { text: 'Polish this deck', nonce: 1, autoRun: true, attachments },
        }),
      ),
    )
    await flushEffects()

    // dataUrls are present on the nodes, so the bridge is never called
    expect(readDeckImage).not.toHaveBeenCalled()
    expect(agentHarness.run).toHaveBeenCalledWith('Polish this deck', [
      { base64: 'AAAA', mime: 'image/png' },
      { base64: 'QUJD', mime: 'image/png' },
      { base64: 'R0w=', mime: 'image/gif' },
    ])
  })

  it('vision off: no deck images, no readDeckImage call, run args unchanged', async () => {
    const readDeckImage = vi.fn(async () => ({ ok: false, error: 'not found' }))
    installSlidesApi(readDeckImage)

    mount(
      createElement(
        AiPanel,
        panelProps({
          slides: [picSlide],
          settings: textOnlySettings,
          preset: { text: 'Polish this deck', nonce: 2, autoRun: true },
        }),
      ),
    )
    await flushEffects()

    expect(readDeckImage).not.toHaveBeenCalled()
    expect(agentHarness.run).toHaveBeenCalledTimes(1)
    // text-only: loop.run received no images argument at all
    expect(agentHarness.run).toHaveBeenCalledWith('Polish this deck', [])
    expect(agentHarness.run.mock.calls[0]![1]).toEqual([])
  })

  it('resolves dataUrl-less pictures through the bridge', async () => {
    const bareSlide = {
      ...picSlide,
      nodes: [
        {
          id: 'r_pic1',
          sourceId: 'pic1',
          type: 'picture',
          box: { x: 0, y: 0, w: 100, h: 100, rotationDeg: 0, flipH: false, flipV: false, centerX: 50, centerY: 50 },
        },
      ],
    } as unknown as RenderSlide
    const readDeckImage = vi.fn(async (op: { slideIndex: number; elementId: string }) => {
      expect(op.slideIndex).toBe(0)
      expect(op.elementId).toBe('pic1')
      return { ok: true, base64: 'QkNF', mime: 'image/png' }
    })
    installSlidesApi(readDeckImage)

    mount(
      createElement(
        AiPanel,
        panelProps({
          slides: [bareSlide],
          preset: { text: 'Polish this deck', nonce: 3, autoRun: true },
        }),
      ),
    )
    await flushEffects()

    expect(readDeckImage).toHaveBeenCalledWith({ slideIndex: 0, elementId: 'pic1' })
    expect(agentHarness.run).toHaveBeenCalledWith('Polish this deck', [
      { base64: 'QkNF', mime: 'image/png' },
    ])
  })

  it('clears the outline attachment state when the run settles (onDone)', async () => {
    installSlidesApi(vi.fn(async () => ({ ok: false })))
    mount(
      createElement(
        AiPanel,
        panelProps({
          slides: [picSlide],
          preset: { text: 'Polish this deck', nonce: 4, autoRun: true },
        }),
      ),
    )
    await flushEffects()
    // The run started with the attachment state set (run received the images)
    expect(agentHarness.run.mock.calls[0]![1]!.length).toBeGreaterThan(0)
    // Settle the run: the images are gone, the outline must not quote ordinals
    act(() => {
      agentHarness.events!.onDone?.({ text: 'done', cancelled: false, turnLimit: false })
    })
    await flushEffects()
    // buildContext of a fresh skill reads the cleared module state
    const qcAccess: DeckAccess = {
      getSlides: () => [picSlide],
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => {},
      applyDeck: () => {},
      fitWidthPx: 960,
    }
    const ctx = createSlidesSkill(qcAccess).buildContext!()
    expect(ctx).not.toContain('attached as image')
  })
})
