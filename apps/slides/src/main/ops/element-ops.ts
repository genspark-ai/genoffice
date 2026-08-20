/**
 * Element-family ops: geometry (transform/flip/connectors/reorder/grouping)
 * and element-level appearance (crop/opacity/geometry preset/anchor/link/
 * image fill). All coordinates are document-space EMU — px→EMU conversion is
 * surface translation and stays in the IPC shims.
 */
import {
  editGroupChildTransform,
  editPictureSrcRect,
  elementSpid,
  findGroupChild,
  groupElements,
  reorderElement,
  resizeTable,
  setElementConnection,
  setElementImageFill,
  setElementLink,
  setElementTextAnchor,
  setPictureOpacity,
  setShapePresetGeometry,
  setGroupChildShapePresetGeometry,
  setShapeAdjustValues,
  setGroupChildShapeAdjustValues,
  ungroupElement,
  updateConnectorsForMoved,
  type EmuRect,
  type ReorderDirection,
  type SlideElement,
} from '@genoffice/pptx-engine'
import {
  GuidedError,
  matchesElementRef,
  register,
  resolveElement,
  resolveGroup,
  resolveGroupChildId,
  resolveSlide,
  type Op,
  type OpRecord,
} from './registry'

function emuRect(op: Op, field = 'box'): EmuRect {
  const b = op[field] as Partial<EmuRect> | undefined
  if (
    !b ||
    typeof b.x !== 'number' ||
    typeof b.y !== 'number' ||
    typeof b.cx !== 'number' ||
    typeof b.cy !== 'number'
  ) {
    throw new GuidedError(`op "${op.op}" needs "${field}": an EMU rect {x, y, cx, cy}.`)
  }
  return { x: b.x, y: b.y, cx: b.cx, cy: b.cy }
}

// ── setTransform ────────────────────────────────────────────────────────
// Top-level: tables optionally redistribute their grid to the new frame
// (legacy parity: the single-element edit path resizes the grid, the batch
// path historically did not — resizeTableGrid keeps both behaviors intact).
// Group children take either a child-space EMU `box` (interactive drags — the
// shim owns that px translation) or a document-space `absBox` converted here
// at apply time (edit scripts — see the comment in apply).
register({
  name: 'setTransform',
  validate(op, ctx) {
    emuRect(op, op.group && op.absBox !== undefined ? 'absBox' : 'box')
    if (op.group) {
      const { index, slide } = resolveSlide(ctx, op)
      resolveGroup(op, index, slide.elements)
      return
    }
    resolveElement(ctx, op, { allowPart: true })
  },
  apply(op, ctx): OpRecord {
    const rotDeg = typeof op.rotDeg === 'number' ? op.rotDeg : 0
    if (op.group) {
      const { index, slide } = resolveSlide(ctx, op)
      const groupId = resolveGroup(op, index, slide.elements)
      const id = resolveGroupChildId(slide, groupId, String(op.target?.el ?? ''))
      // absBox = document-space EMU, converted into the child coordinate system HERE,
      // against the group's live state — earlier ops in the same transaction may have
      // moved/resized the group, and a pre-transaction conversion would double-shift
      // the child. (`box` stays for callers that already speak child space, e.g. the
      // interactive drag shim.)
      let box: EmuRect
      if (op.absBox !== undefined) {
        const found = findGroupChild(slide, groupId, id)
        if (!found) {
          throw new GuidedError(`op "setTransform": no child "${id}" in group "${groupId}".`)
        }
        const abs = emuRect(op, 'absBox')
        const ch = found.grp.childOffset
        const gExt = found.grp.transform.offset
        const chX = ch?.x ?? gExt.x
        const chY = ch?.y ?? gExt.y
        const gsx = ch?.cx ? gExt.cx / ch.cx : 1
        const gsy = ch?.cy ? gExt.cy / ch.cy : 1
        box = {
          x: Math.round((abs.x - gExt.x) / gsx) + chX,
          y: Math.round((abs.y - gExt.y) / gsy) + chY,
          cx: Math.round(abs.cx / gsx),
          cy: Math.round(abs.cy / gsy),
        }
      } else {
        box = emuRect(op)
      }
      if (!editGroupChildTransform(slide, groupId, id, box, rotDeg)) {
        throw new GuidedError(
          `op "setTransform": the child slice for "${id}" could not be located inside group "${groupId}".`,
        )
      }
      return { op, after: { ...box, rotDeg } }
    }
    const box = emuRect(op)
    const { slide, el } = resolveElement(ctx, op, { allowPart: true })
    const before = { ...el.transform.offset, rot: el.transform.rot }
    const isTable = el.type === 'table' && op.resizeTableGrid === true
    if (isTable) resizeTable(slide, el.id, box.cx, box.cy)
    el.transform = {
      ...el.transform,
      offset: {
        x: box.x,
        y: box.y,
        // resizeTable synced cx/cy to the redistributed sums
        cx: isTable ? el.transform.offset.cx : box.cx,
        cy: isTable ? el.transform.offset.cy : box.cy,
      },
      rot: Math.round(rotDeg * 60000),
    }
    el.dirtyTransform = true
    updateConnectorsForMoved(slide, [el.id])
    return { op, before, after: { ...box, rotDeg } }
  },
})

// ── setConnectorEndpoints ───────────────────────────────────────────────
// Box+flip re-derived from the two endpoints; attach/detach writes
// a:stCxn/a:endCxn so the connector follows later shape moves.
register({
  name: 'setConnectorEndpoints',
  validate(op, ctx) {
    resolveElement(ctx, op)
    for (const k of ['p1', 'p2'] as const) {
      const p = op[k] as { x?: unknown; y?: unknown } | undefined
      if (!p || typeof p.x !== 'number' || typeof p.y !== 'number') {
        throw new GuidedError(`op "setConnectorEndpoints" needs "${k}": an EMU point {x, y}.`)
      }
    }
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op)
    const p1 = op.p1 as { x: number; y: number }
    const p2 = op.p2 as { x: number; y: number }
    const before = { ...el.transform.offset }
    el.transform = {
      ...el.transform,
      offset: {
        x: Math.min(p1.x, p2.x),
        y: Math.min(p1.y, p2.y),
        cx: Math.abs(p2.x - p1.x),
        cy: Math.abs(p2.y - p1.y),
      },
      rot: 0,
      flipH: p1.x > p2.x,
      flipV: p1.y > p2.y,
    }
    el.dirtyTransform = true
    const toRef = (
      v: { targetId: string; idx: number } | null | undefined,
    ): { id: number; idx: number } | null | undefined => {
      if (v === undefined) return undefined
      if (v === null) return null
      const target = slide.elements.find((x) => x.id === v.targetId)
      const spid = target ? elementSpid(target) : null
      return spid != null ? { id: spid, idx: v.idx } : null
    }
    setElementConnection(slide, el.id, {
      start: toRef(op.start as { targetId: string; idx: number } | null | undefined),
      end: toRef(op.end as { targetId: string; idx: number } | null | undefined),
    })
    return { op, before, after: { p1, p2 } }
  },
})

// ── flipElements ────────────────────────────────────────────────────────
register({
  name: 'flipElements',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    if (!Array.isArray(op.els) || op.els.length === 0) {
      throw new GuidedError('op "flipElements" needs "els": a non-empty element id array.')
    }
    if (op.axis !== 'h' && op.axis !== 'v') {
      throw new GuidedError('op "flipElements" needs "axis": "h" or "v".')
    }
  },
  apply(op, ctx): OpRecord {
    const { index, slide } = resolveSlide(ctx, op)
    const ids = (op.els as unknown[]).map(String)
    const groupId = op.group ? resolveGroup(op, index, slide.elements) : undefined
    // Any element with an a:xfrm can flip (picture/shape/text/group) — a
    // text/shape filter would silently drop pictures
    const targets = ids
      .map((id) =>
        groupId
          ? findGroupChild(slide, groupId, resolveGroupChildId(slide, groupId, id))?.child
          : slide.elements.find((x) => matchesElementRef(x, id)),
      )
      .filter((el): el is SlideElement => !!el && !!el.transform)
    if (targets.length === 0) {
      throw new GuidedError(
        `op "flipElements": none of [${ids.join(', ')}] are flippable on slide ${index}.`,
      )
    }
    for (const el of targets) {
      const t = el.transform
      // Rotation pivots on the flip-adjusted box origin, so toggling a flip on a
      // rotated element would move its visual center: origin + R(rot)·(flip-signed
      // half-extent) must stay put — shift the offset by the orbit difference.
      const orbit = () => {
        const rad = (((t.rot ?? 0) / 60000) * Math.PI) / 180
        const bx = t.flipH ? t.offset.cx : 0
        const by = t.flipV ? t.offset.cy : 0
        const vx = ((t.flipH ? -1 : 1) * t.offset.cx) / 2
        const vy = ((t.flipV ? -1 : 1) * t.offset.cy) / 2
        return {
          x: bx + vx * Math.cos(rad) - vy * Math.sin(rad),
          y: by + vx * Math.sin(rad) + vy * Math.cos(rad),
        }
      }
      const before = orbit()
      if (op.axis === 'h') t.flipH = !t.flipH
      else t.flipV = !t.flipV
      const after = orbit()
      t.offset.x += Math.round(before.x - after.x)
      t.offset.y += Math.round(before.y - after.y)
      el.dirtyTransform = true
    }
    updateConnectorsForMoved(
      slide,
      targets.map((el) => el.id),
    )
    return { op, after: { flipped: targets.length } }
  },
})

// ── setPictureSrcRect ───────────────────────────────────────────────────
// Crop confirm may also shrink the element frame to the crop frame — one op,
// so a single undo restores frame and crop together.
register({
  name: 'setPictureSrcRect',
  validate(op, ctx) {
    resolveElement(ctx, op)
    if (typeof op.srcRect !== 'object' || op.srcRect === null) {
      throw new GuidedError('op "setPictureSrcRect" needs "srcRect": fractions {l, t, r, b}.')
    }
    if (op.box !== undefined) emuRect(op)
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op)
    if (
      !editPictureSrcRect(
        slide,
        el.id,
        op.srcRect as { l: number; t: number; r: number; b: number },
      )
    ) {
      throw new GuidedError(
        `op "setPictureSrcRect": element "${el.id}" is not a croppable picture.`,
      )
    }
    if (op.box !== undefined) {
      const box = emuRect(op)
      el.transform = { ...el.transform, offset: box }
      el.dirtyTransform = true
      updateConnectorsForMoved(slide, [el.id])
    }
    return { op, after: op.srcRect }
  },
})

// ── setPictureOpacity ───────────────────────────────────────────────────
register({
  name: 'setPictureOpacity',
  validate(op, ctx) {
    resolveElement(ctx, op)
    if (typeof op.opacity !== 'number') {
      throw new GuidedError('op "setPictureOpacity" needs "opacity": a number 0..1.')
    }
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op)
    if (!setPictureOpacity(slide, el.id, op.opacity as number)) {
      throw new GuidedError(`op "setPictureOpacity": element "${el.id}" is not a picture.`)
    }
    return { op, after: op.opacity }
  },
})

// ── reorderElement ──────────────────────────────────────────────────────
register({
  name: 'reorderElement',
  validate(op, ctx) {
    resolveElement(ctx, op)
    if (!['front', 'back', 'forward', 'backward'].includes(String(op.dir))) {
      throw new GuidedError('op "reorderElement" needs "dir": front/back/forward/backward.')
    }
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op)
    if (!reorderElement(slide, el.id, op.dir as ReorderDirection)) {
      throw new GuidedError(`op "reorderElement": element "${el.id}" is already at that extreme.`)
    }
    return { op, after: op.dir }
  },
})

// ── groupElements / ungroupElement ──────────────────────────────────────
register({
  name: 'groupElements',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    if (!Array.isArray(op.els) || op.els.length < 2) {
      throw new GuidedError('op "groupElements" needs "els": at least two element ids.')
    }
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    const r = groupElements(ctx.opened, index, (op.els as unknown[]).map(String))
    if (!r) {
      throw new GuidedError(
        `op "groupElements": the elements could not be grouped (check the ids exist and are groupable).`,
      )
    }
    return { op, created: [r.groupId] }
  },
})

register({
  name: 'ungroupElement',
  validate(op, ctx) {
    resolveElement(ctx, op, { types: ['group'] })
  },
  apply(op, ctx): OpRecord {
    const { index, el } = resolveElement(ctx, op, { types: ['group'] })
    if (!ungroupElement(ctx.opened, index, el.id)) {
      throw new GuidedError(`op "ungroupElement": group "${el.id}" could not be dissolved.`)
    }
    return { op, before: { group: el.id } }
  },
})

// ── setShapeGeometry / setTextAnchor / setLink ──────────────────────────
register({
  name: 'setShapeGeometry',
  validate(op, ctx) {
    if (typeof op.prst !== 'string' || !op.prst) {
      throw new GuidedError('op "setShapeGeometry" needs "prst": an OOXML preset geometry name.')
    }
    if (op.group) {
      const { index, slide } = resolveSlide(ctx, op)
      resolveGroup(op, index, slide.elements)
      return
    }
    resolveElement(ctx, op)
  },
  apply(op, ctx): OpRecord {
    if (op.group) {
      const { index, slide } = resolveSlide(ctx, op)
      const groupId = resolveGroup(op, index, slide.elements)
      const id = resolveGroupChildId(slide, groupId, String(op.target?.el ?? ''))
      if (!setGroupChildShapePresetGeometry(slide, groupId, id, String(op.prst))) {
        throw new GuidedError(
          `op "setShapeGeometry": no child "${id}" in group "${groupId}", or it has no preset geometry.`,
        )
      }
      return { op, after: op.prst }
    }
    const { slide, el } = resolveElement(ctx, op)
    if (!setShapePresetGeometry(slide, el.id, String(op.prst))) {
      throw new GuidedError(`op "setShapeGeometry": element "${el.id}" has no preset geometry.`)
    }
    return { op, after: op.prst }
  },
})

register({
  name: 'setShapeAdjust',
  validate(op, ctx) {
    const adjust = op.adjust as Record<string, unknown> | undefined
    if (
      !adjust ||
      typeof adjust !== 'object' ||
      Array.isArray(adjust) ||
      Object.keys(adjust).length === 0 ||
      Object.values(adjust).some((v) => typeof v !== 'number' || !Number.isFinite(v))
    ) {
      throw new GuidedError(
        'op "setShapeAdjust" needs "adjust": a non-empty { gdName: number } map of avLst values.',
      )
    }
    if (op.group) {
      const { index, slide } = resolveSlide(ctx, op)
      resolveGroup(op, index, slide.elements)
      return
    }
    resolveElement(ctx, op)
  },
  apply(op, ctx): OpRecord {
    const adjust = op.adjust as Record<string, number>
    if (op.group) {
      const { index, slide } = resolveSlide(ctx, op)
      const groupId = resolveGroup(op, index, slide.elements)
      const id = resolveGroupChildId(slide, groupId, String(op.target?.el ?? ''))
      if (!setGroupChildShapeAdjustValues(slide, groupId, id, adjust)) {
        throw new GuidedError(
          `op "setShapeAdjust": no child "${id}" in group "${groupId}", or it has no preset geometry.`,
        )
      }
      return { op, after: adjust }
    }
    const { slide, el } = resolveElement(ctx, op)
    if (!setShapeAdjustValues(slide, el.id, adjust)) {
      throw new GuidedError(`op "setShapeAdjust": element "${el.id}" has no preset geometry.`)
    }
    return { op, after: adjust }
  },
})

register({
  name: 'setTextAnchor',
  validate(op, ctx) {
    resolveElement(ctx, op)
    if (!['top', 'middle', 'bottom'].includes(String(op.anchor))) {
      throw new GuidedError('op "setTextAnchor" needs "anchor": top/middle/bottom.')
    }
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op)
    if (!setElementTextAnchor(slide, el.id, op.anchor as 'top' | 'middle' | 'bottom')) {
      throw new GuidedError(`op "setTextAnchor": element "${el.id}" has no text body.`)
    }
    return { op, after: op.anchor }
  },
})

register({
  name: 'setLink',
  validate(op, ctx) {
    resolveElement(ctx, op)
  },
  apply(op, ctx): OpRecord {
    const { index, el } = resolveElement(ctx, op)
    if (
      !setElementLink(ctx.opened, index, el.id, op.link as Parameters<typeof setElementLink>[3])
    ) {
      throw new GuidedError(`op "setLink": element "${el.id}" does not accept links.`)
    }
    return { op, after: op.link }
  },
})

// ── setImageFill ────────────────────────────────────────────────────────
// One target per op; the record's after.mediaPath lets the shim reuse the
// landed media part across a multi-target selection (one part, many rels).
register({
  name: 'setImageFill',
  validate(op, ctx) {
    const source = op.source as { bytes?: unknown; mediaPath?: unknown } | undefined
    if (!source || (!source.bytes && !source.mediaPath)) {
      throw new GuidedError(
        'op "setImageFill" needs "source": {bytes, ext} or {mediaPath} of an already-landed image.',
      )
    }
    if (op.group) {
      const { index, slide } = resolveSlide(ctx, op)
      resolveGroup(op, index, slide.elements)
      return
    }
    resolveElement(ctx, op)
  },
  apply(op, ctx): OpRecord {
    const { index, slide } = resolveSlide(ctx, op)
    const groupId = op.group ? resolveGroup(op, index, slide.elements) : undefined
    const rawId = String(op.target?.el ?? '')
    // Non-group refs may be durable too — the engine matches parse-time ids only
    const id = groupId ? resolveGroupChildId(slide, groupId, rawId) : resolveElement(ctx, op).el.id
    const used = setElementImageFill(
      ctx.opened,
      slide,
      id,
      op.source as Parameters<typeof setElementImageFill>[3],
      { tile: op.tile === true, ...(groupId ? { groupId } : {}) },
    )
    if (!used) {
      throw new GuidedError(`op "setImageFill": element "${id}" does not support a picture fill.`)
    }
    return { op, after: { mediaPath: used } }
  },
})
