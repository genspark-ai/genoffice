/**
 * Core op set, batch 1: the smallest handler cores moved down into the
 * registry (deleteElement / setFill / setStroke). The matching IPC handlers
 * are now thin shims — surface translation (px→EMU, option normalization) and
 * RenderSlide rebuilding stay in the shim; validation, mutation, and the
 * journal record live here. Later batches migrate the remaining handlers the
 * same way.
 */
import {
  deleteElement as engineDeleteElement,
  editGroupChildStroke,
  findGroupChild,
  setElementFill,
  strokePatchToModel,
  type GradientFillPatch,
  type StrokePatch,
  type TextElement,
} from '@genoffice/pptx-engine'
import {
  GuidedError,
  register,
  resolveElement,
  resolveGroup,
  resolveGroupChildId,
  resolveSlide,
  type OpRecord,
} from './registry'

// ── deleteElement ───────────────────────────────────────────────────────
register({
  name: 'deleteElement',
  validate(op, ctx) {
    resolveElement(ctx, op, { allowPart: true })
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op, { allowPart: true })
    if (!engineDeleteElement(ctx.opened, slide, el.id)) {
      throw new GuidedError(`op "deleteElement": element "${el.id}" could not be removed.`)
    }
    return { op, before: { type: el.type } }
  },
})

// ── setFill ─────────────────────────────────────────────────────────────
register({
  name: 'setFill',
  validate(op, ctx) {
    if (typeof op.fill !== 'string' && (typeof op.fill !== 'object' || op.fill === null)) {
      throw new GuidedError(
        'op "setFill" needs "fill": a solid color string or a gradient patch object.',
      )
    }
    if (op.group) {
      const { index, slide } = resolveSlide(ctx, op)
      resolveGroup(op, index, slide.elements)
      return
    }
    resolveElement(ctx, op, { types: ['text', 'shape'], allowPart: true })
  },
  apply(op, ctx): OpRecord {
    const fill = op.fill as string | GradientFillPatch
    if (op.group) {
      const { index, slide } = resolveSlide(ctx, op)
      const groupId = resolveGroup(op, index, slide.elements)
      const id = resolveGroupChildId(slide, groupId, String(op.target?.el ?? ''))
      if (!setElementFill(ctx.opened, slide, id, fill, { groupId })) {
        const found = findGroupChild(slide, groupId, id)
        const detail = found
          ? 'the child does not support fill'
          : `no child "${id}" in group "${groupId}"`
        throw new GuidedError(`op "setFill": ${detail}.`)
      }
      return { op, after: fill }
    }
    const { slide, el } = resolveElement(ctx, op, { types: ['text', 'shape'], allowPart: true })
    const before = (el as TextElement).fill
    if (!setElementFill(ctx.opened, slide, el.id, fill)) {
      throw new GuidedError(`op "setFill": element "${el.id}" does not support fill.`)
    }
    return { op, before, after: fill }
  },
})

// ── setStroke ───────────────────────────────────────────────────────────
register({
  name: 'setStroke',
  validate(op, ctx) {
    if (op.stroke !== null && (typeof op.stroke !== 'object' || Array.isArray(op.stroke))) {
      throw new GuidedError(
        'op "setStroke" needs "stroke": a stroke patch object, or null to remove the outline.',
      )
    }
    if (op.group) {
      const { index, slide } = resolveSlide(ctx, op)
      resolveGroup(op, index, slide.elements)
      return
    }
    // Pictures are strokable too (picture border)
    resolveElement(ctx, op, { types: ['text', 'shape', 'picture'], allowPart: true })
  },
  apply(op, ctx): OpRecord {
    const patch = op.stroke as StrokePatch | null
    if (op.group) {
      const { index, slide } = resolveSlide(ctx, op)
      const groupId = resolveGroup(op, index, slide.elements)
      const id = resolveGroupChildId(slide, groupId, String(op.target?.el ?? ''))
      if (!editGroupChildStroke(slide, groupId, id, patch)) {
        throw new GuidedError(
          `op "setStroke": no child "${id}" in group "${groupId}", or it does not support strokes.`,
        )
      }
      return { op, after: patch }
    }
    const { el } = resolveElement(ctx, op, { types: ['text', 'shape', 'picture'], allowPart: true })
    const te = el as TextElement
    const before = te.stroke
    te.stroke = patch ? strokePatchToModel(patch) : undefined
    te.dirtyStroke = true
    return { op, before, after: patch }
  },
})
