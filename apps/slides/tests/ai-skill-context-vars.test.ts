// Slides: the context vars an attached skill's body expands at send time.
// ${APP}/${LANGUAGE}/${TODAY} come from the UI language, ${FILE_NAME} from the
// open deck's path, ${SELECTION} from the selected elements on the current
// slide (the redacted model view, capped at 500 chars) — and $ARGUMENTS/$N
// from the request itself. Anything with no value stays literal.
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'

const agentHarness = vi.hoisted(() => ({ run: vi.fn() }))

vi.mock('@genoffice/agent-core', async () => {
  const actual =
    await vi.importActual<typeof import('@genoffice/agent-core')>('@genoffice/agent-core')
  return {
    ...actual,
    AgentLoop: class MockAgentLoop {
      busy = false

      run(...args: unknown[]) {
        agentHarness.run(...args)
      }

      cancel() {}

      reset() {}

      restore() {}
    },
  }
})

// react-konva's node entry requires the native 'canvas' package; nothing here draws
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
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'
import type {
  GroupRenderNode,
  RenderNode,
  RenderSlide,
  ShapeRenderNode,
} from '@genoffice/pptx-render'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((provider) => [provider.id, { apiKey: '', model: provider.defaultModel }]),
  ) as AiSettings['providers'],
}

const glyph = (text: string, redact?: string) => ({
  text,
  x: 0,
  baselineY: 20,
  fontFamily: 'Arial',
  fontSizePx: 24,
  color: '#000',
  bold: false,
  italic: false,
  underline: !!redact,
  widthPx: text.length * 12,
  ...(redact ? { redact } : {}),
})

const shape = (id: string, runs: Array<ReturnType<typeof glyph>>): ShapeRenderNode =>
  ({
    id,
    sourceId: id,
    type: 'shape',
    box: { x: 0, y: 0, w: 400, h: 60 },
    fill: { kind: 'none' },
    text: {
      lines: [{ runs, top: 0, height: 28 }],
      insets: { l: 0, t: 0, r: 0, b: 0 },
      anchor: 'top',
      fontScale: 1,
      contentHeight: 28,
    },
  }) as unknown as ShapeRenderNode

const group = (id: string, children: ShapeRenderNode[]): GroupRenderNode =>
  ({
    id,
    sourceId: id,
    type: 'group',
    box: { x: 0, y: 0, w: 400, h: 200 },
    children,
  }) as unknown as GroupRenderNode

const deck = (...nodes: RenderNode[]): RenderSlide =>
  ({
    index: 0,
    widthPx: 1280,
    heightPx: 720,
    scale: 1,
    nodes,
    background: null,
  }) as unknown as RenderSlide

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
    settings,
    open: true,
    onExpand: () => {},
    onCollapse: () => {},
    ...overrides,
  }
}

/** The body `window.slidesApi.skillText` hands out for the one skill in play */
let skillBody = ''

/** Simulate typing into React's controlled textarea (the caret lands at the end) */
function typeInto(textarea: HTMLTextAreaElement, text: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
  act(() => {
    setter.call(textarea, text)
    textarea.setSelectionRange(text.length, text.length)
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function press(textarea: HTMLTextAreaElement, key: string) {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })
  act(() => textarea.dispatchEvent(event))
  return event
}

async function flushEffects(): Promise<void> {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 0))
    await Promise.resolve()
  })
}

/** Attach the skill through the composer's slash palette, then send the request */
async function sendWithSkill(container: HTMLElement, request: string): Promise<void> {
  const textarea = container.querySelector<HTMLTextAreaElement>('.ai-input-box textarea')!
  typeInto(textarea, '/proof')
  await flushEffects()
  press(textarea, 'Enter')
  typeInto(textarea, request)
  press(textarea, 'Enter')
  await flushEffects()
}

/** The model-facing instruction of the (single) send this test made */
function sentInstruction(): string {
  expect(agentHarness.run).toHaveBeenCalledTimes(1)
  return agentHarness.run.mock.calls[0]![0] as string
}

beforeAll(() => {
  // jsdom has no scrollTo; the panel auto-scrolls its chat log
  Element.prototype.scrollTo ??= () => {}
})

beforeEach(() => {
  agentHarness.run.mockReset()
  skillBody = ''
  Object.defineProperty(window, 'slidesApi', {
    configurable: true,
    value: {
      aiGskStatus: vi.fn(async () => ({ loggedIn: true })),
      beginHistoryBatch: vi.fn(async () => false),
      endHistoryBatch: vi.fn(async () => null),
      aiLogRunFailure: vi.fn(async () => undefined),
      skillsList: vi.fn(async () => [{ name: 'proof', description: 'Tighten the wording' }]),
      skillText: vi.fn(async () => skillBody),
      readDeckImage: vi.fn(async () => ({ ok: false })),
    },
  })
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

describe('AiPanel attached-skill context vars (slides)', () => {
  it("expands every var from the send's own context", async () => {
    skillBody =
      'App: ${APP}\nLang: ${LANGUAGE}\nToday: ${TODAY}\nFile: ${FILE_NAME}\n' +
      'Sel: ${SELECTION}\nFirst: $0\nAll: $ARGUMENTS'
    const { container } = mount(
      createElement(
        AiPanel,
        panelProps({
          slides: [deck(shape('e_1', [glyph('Ship it faster')]))],
          selectedIds: ['e_1'],
          currentFilePath: 'C:\\Decks\\Q3 review.pptx',
        }),
      ),
    )

    await sendWithSkill(container, 'make this snappier')

    const instruction = sentInstruction()
    expect(instruction).toContain('<skill name="proof">')
    // the panel runs in the default UI language (zh) with no provider mounted
    expect(instruction).toContain('App: 幻灯片')
    expect(instruction).toContain('Lang: 中文')
    expect(instruction).toContain(`Today: ${new Date().toLocaleDateString('zh')}`)
    expect(instruction).toContain('File: Q3 review.pptx')
    expect(instruction).toContain('Sel: Ship it faster')
    expect(instruction).toContain('First: make')
    expect(instruction).toContain('All: make this snappier')
    expect(instruction).toContain('User request:\nmake this snappier')
  })

  it('leaves ${FILE_NAME}/${SELECTION} literal with no file and nothing selected', async () => {
    skillBody = 'File: ${FILE_NAME}\nSel: ${SELECTION}\nNope: ${MISSING}'
    const { container } = mount(
      createElement(AiPanel, panelProps({ slides: [deck(shape('e_1', [glyph('Orphan')]))] })),
    )

    await sendWithSkill(container, 'hello world')

    const instruction = sentInstruction()
    expect(instruction).toContain('File: ${FILE_NAME}')
    expect(instruction).toContain('Sel: ${SELECTION}')
    expect(instruction).toContain('Nope: ${MISSING}')
  })

  it('caps ${SELECTION} at 500 chars', async () => {
    skillBody = 'Sel: ${SELECTION}'
    const long = 'x'.repeat(600)
    const { container } = mount(
      createElement(
        AiPanel,
        panelProps({ slides: [deck(shape('e_1', [glyph(long)]))], selectedIds: ['e_1'] }),
      ),
    )

    await sendWithSkill(container, 'trim it')

    const instruction = sentInstruction()
    expect(instruction).toContain(`Sel: ${'x'.repeat(499)}…`)
    expect(instruction).not.toContain('x'.repeat(500))
  })

  it('gives the skill the redacted view of a withheld selection', async () => {
    skillBody = 'Sel: ${SELECTION}'
    const { container } = mount(
      createElement(
        AiPanel,
        panelProps({
          slides: [deck(shape('e_1', [glyph('Call '), glyph('555-0123', 'client phone')]))],
          selectedIds: ['e_1'],
        }),
      ),
    )

    await sendWithSkill(container, 'proofread')

    const instruction = sentInstruction()
    expect(instruction).toContain('Sel: Call {{client phone}}')
    expect(instruction).not.toContain('555-0123')
  })

  it('reads a selected group as its children\u2019s text', async () => {
    skillBody = 'Sel: ${SELECTION}'
    const { container } = mount(
      createElement(
        AiPanel,
        panelProps({
          slides: [deck(group('e_9', [shape('e_10', [glyph('grouped text')])]))],
          selectedIds: ['e_9'],
        }),
      ),
    )

    await sendWithSkill(container, 'tighten')

    expect(sentInstruction()).toContain('Sel: grouped text')
  })
})
