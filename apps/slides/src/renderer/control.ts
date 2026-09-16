import type { RenderNode, RenderSlide } from '@genoffice/pptx-render'

/** Request/reply the shell relays from `genoffice open --slide/--el` and `genoffice selection`. */
export type ControlRequest =
  { cmd: 'goto'; target: { kind: string; slide?: number; el?: string } } | { cmd: 'selection' }

export type ControlReply =
  | { status: 'ok'; result: Record<string, unknown> }
  | { status: 'not_ready' }
  | {
      status: 'error'
      error: { reason: string; message: string; detail?: Record<string, unknown> }
    }

export interface SlidesControlState {
  slides: RenderSlide[]
  path: string | null
  current: number
  selectedIds: string[]
  setCurrent: (index: number) => void
  setSelectedIds: (ids: string[]) => void
  clearEditing: () => void
}

const publicId = (node: RenderNode) => node.durableId ?? node.sourceId

/** Selected ids are child sourceIds once the user has entered a group. */
function findNode(nodes: RenderNode[], id: string): RenderNode | null {
  for (const n of nodes) {
    if (n.sourceId === id || n.id === id) return n
    if (n.type === 'group') {
      const inner = findNode(n.children, id)
      if (inner) return inner
    }
  }
  return null
}

/** The top-level node that is or contains the element with this id. */
function topLevelNodeFor(slide: RenderSlide, id: string): RenderNode | null {
  const matches = (n: RenderNode) => n.durableId === id || n.sourceId === id || n.id === id
  const contains = (n: RenderNode): boolean =>
    matches(n) || (n.type === 'group' && n.children.some(contains))
  return slide.nodes.find((n) => !n.decoration && contains(n)) ?? null
}

export function handleSlidesControl(req: ControlRequest, state: SlidesControlState): ControlReply {
  if (!state.path || state.slides.length === 0) return { status: 'not_ready' }
  if (req.cmd === 'selection') {
    const slide = state.slides[state.current]
    const selected = state.selectedIds
      .map((id) => (slide ? findNode(slide.nodes, id) : null))
      .filter((n): n is RenderNode => Boolean(n))
    return {
      status: 'ok',
      result: {
        slide: state.current,
        elements: selected.map(publicId),
        types: selected.map((n) => n.type),
      },
    }
  }
  const { slide, el } = req.target
  const count = state.slides.length
  if (slide === undefined || !Number.isInteger(slide) || slide < 0 || slide >= count) {
    return {
      status: 'error',
      error: {
        reason: 'out_of_range',
        message: `slide ${slide} is out of range (the deck has ${count} slides)`,
        detail: { valid_range: `0-${count - 1}` },
      },
    }
  }
  const target = state.slides[slide]!
  if (el !== undefined) {
    const node = topLevelNodeFor(target, el)
    if (!node) {
      return {
        status: 'error',
        error: {
          reason: 'target_not_found',
          message: `no element ${el} on slide ${slide}`,
          detail: { available: target.nodes.filter((n) => !n.decoration).map(publicId) },
        },
      }
    }
    state.clearEditing()
    if (slide !== state.current) state.setCurrent(slide)
    state.setSelectedIds([node.sourceId])
    return { status: 'ok', result: { slide, element: publicId(node), type: node.type } }
  }
  state.clearEditing()
  if (slide !== state.current) state.setCurrent(slide)
  state.setSelectedIds([])
  return { status: 'ok', result: { slide } }
}
