/**
 * Arrangement ops: align / distribute a set of top-level elements and add a
 * connector glued to two shapes. Geometry is document-space EMU on the
 * axis-aligned frame (rotation is ignored, as PowerPoint's Align menu does).
 */
import {
  addElement,
  alignRects,
  connectionSiteForSide,
  distributeRects,
  elementDurableId,
  elementSpid,
  matchesElementRef,
  setElementConnection,
  updateConnectorsForMoved,
  type AlignKind,
  type ConnectionSide,
  type Slide,
  type SlideElement,
} from '@genoffice/pptx-engine'
import {
  GuidedError,
  register,
  requireFinite,
  requireHexColor,
  resolveSlide,
  type Op,
  type OpContext,
  type OpRecord,
} from './registry'

const ALIGN_MODES: Record<string, AlignKind> = {
  left: 'left',
  centerH: 'center-h',
  right: 'right',
  top: 'top',
  centerV: 'center-v',
  bottom: 'bottom',
}

function availableIds(slide: Slide): string {
  return slide.elements
    .map((x) => {
      const durable = elementDurableId(x)
      return durable ? `${x.id} (${durable})` : x.id
    })
    .join(', ')
}

/** Top-level elements for `els`, in the caller's order; group children are refused (their frame is group-local). */
function resolveTopLevel(
  op: Op,
  ctx: OpContext,
  min: number,
): { slide: Slide; els: SlideElement[] } {
  const { index, slide } = resolveSlide(ctx, op)
  const ids = op.els
  if (!Array.isArray(ids) || ids.length < min) {
    throw new GuidedError(`op "${op.op}" needs "els": at least ${min} element ids.`)
  }
  const els = (ids as unknown[]).map((raw) => {
    const id = String(raw)
    const el = slide.elements.find((x) => matchesElementRef(x, id))
    if (!el) {
      const inGroup = slide.elements.some(
        (g) => g.type === 'group' && g.children.some((c) => matchesElementRef(c, id)),
      )
      throw new GuidedError(
        inGroup
          ? `op "${op.op}": "${id}" is inside a group — arrange the group itself, or ungroupElement first.`
          : `op "${op.op}": no element "${id}" on slide ${index}. Available: [${availableIds(slide)}].`,
      )
    }
    return el
  })
  if (new Set(els).size !== els.length) {
    throw new GuidedError(`op "${op.op}": "els" lists the same element twice.`)
  }
  return { slide, els }
}

function toRef(op: Op): 'selection' | 'slide' {
  if (op.to === undefined || op.to === 'selection') return 'selection'
  if (op.to === 'slide') return 'slide'
  throw new GuidedError(`op "${op.op}": "to" must be "selection" (default) or "slide".`)
}

function moveTo(slide: Slide, els: SlideElement[], pos: Array<{ x: number; y: number }>): number {
  const moved: string[] = []
  els.forEach((el, i) => {
    const x = Math.round(pos[i]!.x)
    const y = Math.round(pos[i]!.y)
    if (x === el.transform.offset.x && y === el.transform.offset.y) return
    el.transform = { ...el.transform, offset: { ...el.transform.offset, x, y } }
    el.dirtyTransform = true
    moved.push(el.id)
  })
  if (moved.length) updateConnectorsForMoved(slide, moved)
  return moved.length
}

const rectsOf = (els: SlideElement[]) =>
  els.map((el) => {
    const o = el.transform.offset
    return { x: o.x, y: o.y, w: o.cx, h: o.cy }
  })

// ── alignElements ───────────────────────────────────────────────────────
register({
  name: 'alignElements',
  validate(op, ctx) {
    if (!ALIGN_MODES[String(op.mode)]) {
      throw new GuidedError(
        `op "alignElements" needs "mode": ${Object.keys(ALIGN_MODES).join('/')}.`,
      )
    }
    resolveTopLevel(op, ctx, toRef(op) === 'slide' ? 1 : 2)
  },
  apply(op, ctx): OpRecord {
    const to = toRef(op)
    const { slide, els } = resolveTopLevel(op, ctx, to === 'slide' ? 1 : 2)
    const size = ctx.opened.deck.size
    const container = to === 'slide' ? { x: 0, y: 0, w: size.cx, h: size.cy } : null
    const before = els.map((el) => ({ id: el.id, ...el.transform.offset }))
    const pos = alignRects(rectsOf(els), ALIGN_MODES[String(op.mode)]!, container)
    const moved = moveTo(slide, els, pos)
    return { op, before, after: { mode: op.mode, to, moved } }
  },
})

// ── distributeElements ──────────────────────────────────────────────────
// selection: outer elements stay, inner ones get even gaps (PowerPoint);
// slide: even gaps between the slide edges and every element.
register({
  name: 'distributeElements',
  validate(op, ctx) {
    if (op.axis !== 'horizontal' && op.axis !== 'vertical') {
      throw new GuidedError('op "distributeElements" needs "axis": "horizontal" or "vertical".')
    }
    resolveTopLevel(op, ctx, toRef(op) === 'slide' ? 1 : 3)
  },
  apply(op, ctx): OpRecord {
    const to = toRef(op)
    const { slide, els } = resolveTopLevel(op, ctx, to === 'slide' ? 1 : 3)
    const horizontal = op.axis === 'horizontal'
    const rects = rectsOf(els)
    const before = els.map((el) => ({ id: el.id, ...el.transform.offset }))
    let pos: Array<{ x: number; y: number }>
    if (to === 'selection') {
      pos = distributeRects(rects, horizontal ? 'horizontal' : 'vertical')
    } else {
      const size = ctx.opened.deck.size
      const span = horizontal ? size.cx : size.cy
      const total = rects.reduce((s, r) => s + (horizontal ? r.w : r.h), 0)
      const gap = (span - total) / (rects.length + 1)
      const order = rects
        .map((r, i) => ({ r, i }))
        .sort((a, b) => (horizontal ? a.r.x - b.r.x : a.r.y - b.r.y))
      pos = rects.map((r) => ({ x: r.x, y: r.y }))
      let cursor = gap
      for (const { r, i } of order) {
        pos[i] = horizontal ? { x: cursor, y: r.y } : { x: r.x, y: cursor }
        cursor += (horizontal ? r.w : r.h) + gap
      }
    }
    const moved = moveTo(slide, els, pos)
    return { op, before, after: { axis: op.axis, to, moved } }
  },
})

// ── addConnector ────────────────────────────────────────────────────────
const CONNECTOR_KINDS: Record<string, { none: string; arrow: string }> = {
  straight: { none: 'line', arrow: 'lineArrow' },
  elbow: { none: 'lineBent', arrow: 'lineBent' },
  curved: { none: 'lineCurved', arrow: 'lineCurved' },
}
const SIDES: ConnectionSide[] = ['top', 'left', 'bottom', 'right']
const DASHES = new Set([
  'solid',
  'dash',
  'dot',
  'lgDash',
  'dashDot',
  'lgDashDot',
  'sysDash',
  'sysDot',
  'sysDashDot',
])
const EMU_PER_PT = 12700

function sideMid(el: SlideElement, side: ConnectionSide): { x: number; y: number } {
  const o = el.transform.offset
  switch (side) {
    case 'top':
      return { x: o.x + o.cx / 2, y: o.y }
    case 'left':
      return { x: o.x, y: o.y + o.cy / 2 }
    case 'bottom':
      return { x: o.x + o.cx / 2, y: o.y + o.cy }
    case 'right':
      return { x: o.x + o.cx, y: o.y + o.cy / 2 }
  }
}

function sideOf(op: Op, field: 'fromSide' | 'toSide'): ConnectionSide | undefined {
  const v = op[field]
  if (v === undefined) return undefined
  if (!SIDES.includes(v as ConnectionSide)) {
    throw new GuidedError(`op "addConnector": "${field}" must be ${SIDES.join('/')}.`)
  }
  return v as ConnectionSide
}

/** The side pair whose edge midpoints are closest; a given side pins its end. */
function pickSides(
  from: SlideElement,
  to: SlideElement,
  fromSide?: ConnectionSide,
  toSide?: ConnectionSide,
): [ConnectionSide, ConnectionSide] {
  let best: [ConnectionSide, ConnectionSide] = [fromSide ?? 'right', toSide ?? 'left']
  let bestD = Infinity
  for (const a of fromSide ? [fromSide] : SIDES) {
    for (const b of toSide ? [toSide] : SIDES) {
      const p = sideMid(from, a)
      const q = sideMid(to, b)
      const d = (p.x - q.x) ** 2 + (p.y - q.y) ** 2
      if (d < bestD) {
        bestD = d
        best = [a, b]
      }
    }
  }
  return best
}

function resolveEndpoint(op: Op, slide: Slide, field: 'from' | 'to', index: number): SlideElement {
  const id = op[field]
  if (typeof id !== 'string' || !id) {
    throw new GuidedError(`op "addConnector" needs "${field}": an element id on slide ${index}.`)
  }
  const el = slide.elements.find((x) => matchesElementRef(x, id))
  if (!el) {
    throw new GuidedError(
      `op "addConnector": no element "${id}" on slide ${index}. Available: [${availableIds(slide)}].`,
    )
  }
  if (elementSpid(el) == null) {
    throw new GuidedError(`op "addConnector": "${id}" has no shape id to attach to.`)
  }
  return el
}

function connectorLine(op: Op): { color: string; widthEmu: number; dash?: string } {
  const line = (op.line ?? {}) as {
    color?: unknown
    widthPt?: unknown
    widthEmu?: unknown
    dash?: unknown
  }
  if (line.color !== undefined) requireHexColor(line.color, 'addConnector', 'line.color')
  if (line.widthPt !== undefined) requireFinite(line.widthPt, 'addConnector', 'line.widthPt')
  if (line.widthEmu !== undefined) requireFinite(line.widthEmu, 'addConnector', 'line.widthEmu')
  if (line.dash !== undefined && !DASHES.has(String(line.dash))) {
    throw new GuidedError(`op "addConnector": "line.dash" must be one of ${[...DASHES].join('/')}.`)
  }
  const widthEmu =
    line.widthEmu !== undefined
      ? Math.round(line.widthEmu as number)
      : line.widthPt !== undefined
        ? Math.round((line.widthPt as number) * EMU_PER_PT)
        : EMU_PER_PT
  if (widthEmu <= 0) throw new GuidedError('op "addConnector": line width must be > 0.')
  return {
    color: typeof line.color === 'string' ? line.color : '#000000',
    widthEmu,
    ...(line.dash !== undefined && line.dash !== 'solid' ? { dash: String(line.dash) } : {}),
  }
}

function arrowOf(op: Op): 'none' | 'end' | 'both' {
  if (op.arrow === undefined || op.arrow === 'end') return 'end'
  if (op.arrow === 'none' || op.arrow === 'both') return op.arrow
  throw new GuidedError('op "addConnector": "arrow" must be none/end/both.')
}

register({
  name: 'addConnector',
  validate(op, ctx) {
    const { index, slide } = resolveSlide(ctx, op)
    const from = resolveEndpoint(op, slide, 'from', index)
    const to = resolveEndpoint(op, slide, 'to', index)
    if (from === to) throw new GuidedError('op "addConnector": "from" and "to" must differ.')
    if (op.kind !== undefined && !CONNECTOR_KINDS[String(op.kind)]) {
      throw new GuidedError('op "addConnector": "kind" must be straight/elbow/curved.')
    }
    sideOf(op, 'fromSide')
    sideOf(op, 'toSide')
    arrowOf(op)
    connectorLine(op)
  },
  apply(op, ctx): OpRecord {
    const { index, slide } = resolveSlide(ctx, op)
    const from = resolveEndpoint(op, slide, 'from', index)
    const to = resolveEndpoint(op, slide, 'to', index)
    const kind = CONNECTOR_KINDS[String(op.kind ?? 'straight')]!
    const arrow = arrowOf(op)
    const line = connectorLine(op)
    const [a, b] = pickSides(from, to, sideOf(op, 'fromSide'), sideOf(op, 'toSide'))
    const p1 = sideMid(from, a)
    const p2 = sideMid(to, b)
    const offset = {
      x: Math.round(Math.min(p1.x, p2.x)),
      y: Math.round(Math.min(p1.y, p2.y)),
      cx: Math.round(Math.abs(p2.x - p1.x)),
      cy: Math.round(Math.abs(p2.y - p1.y)),
    }
    const el = addElement(slide, {
      kind: arrow === 'none' ? kind.none : kind.arrow,
      offset,
      stroke: { color: line.color, widthEmu: line.widthEmu },
    })
    // addElement only knows straight arrow variants; write the ends (and dash) directly
    const head = arrow === 'both' ? '<a:headEnd type="triangle" w="med" len="med"/>' : ''
    const tail = arrow === 'none' ? '' : '<a:tailEnd type="triangle" w="med" len="med"/>'
    const dash = line.dash ? `<a:prstDash val="${line.dash}"/>` : ''
    el.anchor.originalXml = el.anchor.originalXml
      .replace(/<a:headEnd\b[^>]*\/>|<a:tailEnd\b[^>]*\/>/g, '')
      .replace('</a:ln>', `${dash}${head}${tail}</a:ln>`)
    // direction lives in the xfrm flips; dirtyTransform makes setElementConnection bake them in
    el.transform = { ...el.transform, flipH: p1.x > p2.x, flipV: p1.y > p2.y }
    el.dirtyTransform = true
    setElementConnection(slide, el.id, {
      start: { id: elementSpid(from)!, idx: connectionSiteForSide(from, a) },
      end: { id: elementSpid(to)!, idx: connectionSiteForSide(to, b) },
    })
    return { op, created: [el.id], after: { fromSide: a, toSide: b, offset } }
  },
})
