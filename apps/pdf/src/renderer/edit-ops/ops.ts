/**
 * Op definitions over the pending-edit buckets. Each op's apply reads only the
 * buckets it declares in `touches` (see registry.ts). Cross-bucket rules the UI
 * used to spell out at every call site live here once: deleting a page drops its
 * pending markups/drawings, rotating a page keeps image stamps upright, deleting
 * a saved note voids its pending content edit, removing a static-form-fill image
 * edit turns into a delete of the underlying image.
 */
import type {
  DrawingInput,
  FormValueInput,
  ImageEditInput,
  ImageLayer,
  MarkupInput,
  MetadataInput,
  StaticFormFillRecord,
  TextEditInput,
  TextInsertInput,
} from '../../shared/ipc'
import type { LocalMarkup } from '../annotations'
import type { LocalDrawing } from '../DrawLayer'
import type { LocalImageEdit } from '../ImageEditLayer'
import { imageRectKey } from '../ImageEditLayer'
import type { LocalTextEdit, LocalTextInsert } from '../text-edit-preview'
import type { SavedNoteAnnot } from '../note-threads'
import type { SavedMarkupAnnot, StampConfig } from '../edit-state'
import { GuidedError, register, type Op, type OpContext } from './registry'

type Rect = [number, number, number, number]

const str = (v: unknown, field: string): string => {
  if (typeof v !== 'string' || v.length === 0)
    throw new GuidedError(`"${field}" must be a non-empty string`)
  return v
}

const obj = <T>(v: unknown, field: string): T => {
  if (!v || typeof v !== 'object') throw new GuidedError(`"${field}" must be an object`)
  return v as T
}

const pageIndex = (v: unknown, ctx: OpContext, field = 'pageIndex'): number => {
  if (typeof v !== 'number' || !Number.isInteger(v))
    throw new GuidedError(`"${field}" must be a 0-based original page index`)
  if (v < 0 || v >= ctx.pageCount)
    throw new GuidedError(
      `${field} ${v} is out of range: the document has pages 0..${ctx.pageCount - 1}`,
    )
  if (ctx.deleted.has(v)) throw new GuidedError(`Page ${v} is deleted; undo the deletion first`)
  return v
}

const rect = (v: unknown, field: string): Rect => {
  if (!Array.isArray(v) || v.length !== 4 || v.some((n) => typeof n !== 'number'))
    throw new GuidedError(`"${field}" must be [x1,y1,x2,y2] in PDF user space`)
  return v as Rect
}

const id = (op: Op): string => str(op.id, 'id')

// ── annotate ──────────────────────────────────────────────────────────

register({
  name: 'addMarkup',
  touches: ['markups'],
  additive: true,
  validate(op, ctx) {
    const m = obj<MarkupInput>(op.markup, 'markup')
    pageIndex(m.pageIndex, ctx, 'markup.pageIndex')
    if (!['highlight', 'underline', 'strikeout'].includes(m.type))
      throw new GuidedError('markup.type must be highlight | underline | strikeout')
    if (!Array.isArray(m.quads) || m.quads.length === 0)
      throw new GuidedError('markup.quads must hold at least one 8-number quad')
  },
  apply(op, s) {
    const m = op.markup as MarkupInput
    const added: LocalMarkup = { id: id(op), ...m }
    return { markups: [...s.markups, added] }
  },
})

register({
  name: 'removeMarkup',
  touches: ['markups'],
  validate(op) {
    id(op)
  },
  apply(op, s) {
    return { markups: s.markups.filter((m) => m.id !== op.id) }
  },
})

register({
  name: 'deleteSavedAnnot',
  touches: ['annotDeletes', 'noteEdits'],
  additive: true,
  validate(op, ctx) {
    const a = obj<SavedMarkupAnnot | SavedNoteAnnot>(op.annot, 'annot')
    pageIndex(a.pageIndex, ctx, 'annot.pageIndex')
    if (typeof a.objNum !== 'number') throw new GuidedError('annot.objNum must be a number')
  },
  apply(op, s) {
    const annot = op.annot as SavedMarkupAnnot | SavedNoteAnnot
    return {
      annotDeletes: [...s.annotDeletes, { id: id(op), annot }],
      noteEdits: s.noteEdits.filter((e) => e.annot.objNum !== annot.objNum),
    }
  },
})

register({
  name: 'editSavedNote',
  touches: ['noteEdits'],
  additive: true,
  validate(op, ctx) {
    const a = obj<SavedNoteAnnot>(op.annot, 'annot')
    pageIndex(a.pageIndex, ctx, 'annot.pageIndex')
    if (typeof op.contents !== 'string') throw new GuidedError('"contents" must be a string')
  },
  apply(op, s) {
    const annot = op.annot as SavedNoteAnnot
    const contents = op.contents as string
    const rest = s.noteEdits.filter(
      (e) => e.annot.objNum !== annot.objNum || e.annot.pageIndex !== annot.pageIndex,
    )
    // Restoring the on-disk text drops the entry — unless a save for this note is
    // in flight (force): then the entry must stay to target whichever text lands
    const noop = contents === annot.contents && op.force !== true
    return { noteEdits: noop ? rest : [...rest, { id: id(op), annot, contents }] }
  },
})

register({
  name: 'addDrawing',
  touches: ['drawings'],
  additive: true,
  validate(op, ctx) {
    const d = obj<DrawingInput>(op.drawing, 'drawing')
    pageIndex(d.pageIndex, ctx, 'drawing.pageIndex')
    if (!['ink', 'rect', 'ellipse', 'line', 'arrow', 'image', 'note'].includes(d.kind))
      throw new GuidedError(
        'drawing.kind must be ink | rect | ellipse | line | arrow | image | note',
      )
    if (d.kind === 'note' && typeof d.contents !== 'string')
      throw new GuidedError('a note drawing needs string contents')
  },
  apply(op, s) {
    const drawing = op.drawing as DrawingInput
    const added: LocalDrawing = {
      id: id(op),
      // Pending notes are keyed by their record id for replies and thread lookup
      input: drawing.kind === 'note' ? { ...drawing, localId: id(op) } : drawing,
      ...(typeof op.formWidgetId === 'string' ? { formWidgetId: op.formWidgetId } : {}),
    }
    return { drawings: [...s.drawings, added] }
  },
})

register({
  name: 'removeDrawing',
  touches: ['drawings'],
  validate(op) {
    id(op)
  },
  apply(op, s) {
    return { drawings: s.drawings.filter((d) => d.id !== op.id) }
  },
})

register({
  name: 'setNoteContents',
  touches: ['drawings'],
  validate(op) {
    id(op)
    if (typeof op.contents !== 'string') throw new GuidedError('"contents" must be a string')
  },
  apply(op, s) {
    return {
      drawings: s.drawings.map((d) =>
        d.id === op.id && d.input.kind === 'note'
          ? { ...d, input: { ...d.input, contents: op.contents as string } }
          : d,
      ),
    }
  },
})

register({
  name: 'moveDrawing',
  touches: ['drawings'],
  validate(op) {
    id(op)
    if (typeof op.dx !== 'number' || typeof op.dy !== 'number')
      throw new GuidedError('"dx" and "dy" must be numbers (PDF user space)')
  },
  apply(op, s) {
    const dx = op.dx as number
    const dy = op.dy as number
    return {
      drawings: s.drawings.map((d) => {
        if (d.id !== op.id) return d
        const input = d.input
        switch (input.kind) {
          case 'ink':
            return {
              ...d,
              input: {
                ...input,
                paths: input.paths.map((p) => p.map((v, i) => (i % 2 === 0 ? v + dx : v + dy))),
              },
            }
          case 'rect':
          case 'ellipse':
          case 'image':
            return {
              ...d,
              input: {
                ...input,
                rect: [
                  input.rect[0] + dx,
                  input.rect[1] + dy,
                  input.rect[2] + dx,
                  input.rect[3] + dy,
                ] as Rect,
              },
            }
          case 'line':
          case 'arrow':
            return {
              ...d,
              input: {
                ...input,
                from: [input.from[0] + dx, input.from[1] + dy] as [number, number],
                to: [input.to[0] + dx, input.to[1] + dy] as [number, number],
              },
            }
          default:
            return d
        }
      }),
    }
  },
})

register({
  name: 'setDrawingRect',
  touches: ['drawings'],
  validate(op) {
    id(op)
    rect(op.rect, 'rect')
  },
  apply(op, s) {
    const r = op.rect as Rect
    return {
      drawings: s.drawings.map((d) =>
        d.id === op.id && d.input.kind === 'image' ? { ...d, input: { ...d.input, rect: r } } : d,
      ),
    }
  },
})

// ── text ──────────────────────────────────────────────────────────────

register({
  name: 'putTextEdit',
  touches: ['textEdits'],
  additive: true,
  validate(op, ctx) {
    const e = obj<TextEditInput>(op.input, 'input')
    pageIndex(e.pageIndex, ctx, 'input.pageIndex')
    rect(e.rect, 'input.rect')
    if (typeof e.oldText !== 'string' || typeof e.newText !== 'string')
      throw new GuidedError('input.oldText and input.newText must be strings')
  },
  apply(op, s) {
    const edit: LocalTextEdit = {
      id: id(op),
      input: op.input as TextEditInput,
      ...(op.cover ? { cover: op.cover as Rect } : {}),
      ...(op.moveBy ? { moveBy: op.moveBy as [number, number] } : {}),
      ...(op.baseInk ? { baseInk: op.baseInk as string } : {}),
      ...(op.baseFont ? { baseFont: op.baseFont as LocalTextEdit['baseFont'] } : {}),
    }
    const i = s.textEdits.findIndex((e) => e.id === edit.id)
    return {
      textEdits: i < 0 ? [...s.textEdits, edit] : s.textEdits.map((e, j) => (j === i ? edit : e)),
    }
  },
})

register({
  name: 'removeTextEdit',
  touches: ['textEdits'],
  validate(op) {
    id(op)
  },
  apply(op, s) {
    return { textEdits: s.textEdits.filter((e) => e.id !== op.id) }
  },
})

register({
  name: 'addTextInsert',
  touches: ['textInserts'],
  additive: true,
  validate(op, ctx) {
    const e = obj<TextInsertInput>(op.input, 'input')
    pageIndex(e.pageIndex, ctx, 'input.pageIndex')
    if (typeof e.text !== 'string') throw new GuidedError('input.text must be a string')
  },
  apply(op, s) {
    const added: LocalTextInsert = { id: id(op), input: op.input as TextInsertInput }
    return { textInserts: [...s.textInserts, added] }
  },
})

register({
  name: 'removeTextInsert',
  touches: ['textInserts'],
  validate(op) {
    id(op)
  },
  apply(op, s) {
    return { textInserts: s.textInserts.filter((e) => e.id !== op.id) }
  },
})

register({
  name: 'patchTextInsert',
  touches: ['textInserts'],
  validate(op) {
    id(op)
    obj(op.input, 'input')
  },
  apply(op, s) {
    const patch = op.input as Partial<TextInsertInput>
    return {
      textInserts: s.textInserts.map((e) =>
        e.id === op.id ? { ...e, input: { ...e.input, ...patch } } : e,
      ),
    }
  },
})

// ── image ─────────────────────────────────────────────────────────────

const claimKey = (input: ImageEditInput): string | null =>
  input.kind === 'insertImage' ? null : `${input.pageIndex}:${imageRectKey(input.oldRect)}`

register({
  name: 'addImageEdit',
  touches: ['imageEdits'],
  additive: true,
  validate(op, ctx) {
    const e = obj<ImageEditInput>(op.input, 'input')
    pageIndex(e.pageIndex, ctx, 'input.pageIndex')
    if (!['insertImage', 'transformImage', 'replaceImage', 'deleteImage'].includes(e.kind))
      throw new GuidedError(
        'input.kind must be insertImage | transformImage | replaceImage | deleteImage',
      )
    if (e.kind !== 'deleteImage') rect(e.rect, 'input.rect')
    if (e.kind !== 'insertImage') {
      rect(e.oldRect, 'input.oldRect')
      if (ctx.claimedImages.has(claimKey(e)!))
        throw new GuidedError(
          `The image at ${e.oldRect.join(',')} on page ${e.pageIndex} already has a pending edit; remove or update that edit instead`,
        )
    }
  },
  advance(op, ctx) {
    const key = claimKey(op.input as ImageEditInput)
    return key ? { claimedImages: new Set(ctx.claimedImages).add(key) } : {}
  },
  apply(op, s) {
    const added: LocalImageEdit = {
      id: id(op),
      input: op.input as ImageEditInput,
      ...(op.png !== undefined ? { png: op.png as string | null } : {}),
      ...(typeof op.origAbove === 'boolean' ? { origAbove: op.origAbove } : {}),
      ...(op.staticFill ? { staticFill: op.staticFill as StaticFormFillRecord } : {}),
      ...(typeof op.opacityBase === 'string' ? { opacityBase: op.opacityBase } : {}),
    }
    return { imageEdits: [...s.imageEdits, added] }
  },
})

register({
  name: 'setImageEditRect',
  touches: ['imageEdits'],
  validate(op) {
    id(op)
    rect(op.rect, 'rect')
  },
  apply(op, s) {
    const r = op.rect as Rect
    return {
      imageEdits: s.imageEdits.map((e) =>
        e.id === op.id && e.input.kind !== 'deleteImage'
          ? { ...e, input: { ...e.input, rect: r } }
          : e,
      ),
    }
  },
})

register({
  name: 'bakeImageEdit',
  touches: ['imageEdits'],
  validate(op) {
    id(op)
    str(op.image, 'image')
    rect(op.rect, 'rect')
  },
  apply(op, s) {
    const image = op.image as string
    const r = op.rect as Rect
    const opacityBase = typeof op.opacityBase === 'string' ? op.opacityBase : undefined
    return {
      imageEdits: s.imageEdits.map((e) => {
        if (e.id !== op.id) return e
        if (e.input.kind === 'insertImage')
          return { ...e, opacityBase, input: { ...e.input, image, rect: r } }
        if (e.input.kind !== 'transformImage' && e.input.kind !== 'replaceImage') return e
        // Transform ops morph into a replace with their quarter turns baked in
        return {
          ...e,
          opacityBase,
          input: {
            kind: 'replaceImage' as const,
            pageIndex: e.input.pageIndex,
            oldRect: e.input.oldRect,
            rect: r,
            image,
            ...(e.input.layer ? { layer: e.input.layer as ImageLayer } : {}),
          },
        }
      }),
    }
  },
})

register({
  name: 'patchImageEdit',
  touches: ['imageEdits'],
  validate(op) {
    id(op)
    if (op.input !== undefined) obj(op.input, 'input')
    if (
      op.opacityBase !== undefined &&
      op.opacityBase !== null &&
      typeof op.opacityBase !== 'string'
    )
      throw new GuidedError('"opacityBase" must be a string, or null to clear it')
  },
  apply(op, s) {
    const patch = (op.input ?? {}) as Partial<Extract<ImageEditInput, { kind: 'replaceImage' }>>
    return {
      imageEdits: s.imageEdits.map((e) => {
        if (e.id !== op.id || e.input.kind === 'deleteImage') return e
        const next: LocalImageEdit = { ...e, input: { ...e.input, ...patch } as ImageEditInput }
        if (op.opacityBase === null) delete next.opacityBase
        else if (typeof op.opacityBase === 'string') next.opacityBase = op.opacityBase
        return next
      }),
    }
  },
})

register({
  name: 'setStaticFillImage',
  touches: ['imageEdits'],
  validate(op) {
    id(op)
    str(op.image, 'image')
    obj(op.staticFill, 'staticFill')
  },
  apply(op, s) {
    const image = op.image as string
    const staticFill = op.staticFill as StaticFormFillRecord
    return {
      imageEdits: s.imageEdits.map((e) => {
        if (e.id !== op.id || e.input.kind === 'deleteImage') return e
        // A re-typed fill on a transform keeps the footprint but now carries pixels
        const input: ImageEditInput =
          e.input.kind === 'transformImage'
            ? {
                kind: 'replaceImage',
                pageIndex: e.input.pageIndex,
                oldRect: e.input.oldRect,
                rect: e.input.rect,
                image,
                layer: e.input.layer,
                quarterTurns: e.input.quarterTurns,
              }
            : { ...e.input, image }
        return { ...e, input, staticFill: { ...staticFill, rect: input.rect } }
      }),
    }
  },
})

register({
  name: 'removeImageEdit',
  touches: ['imageEdits'],
  validate(op) {
    id(op)
  },
  apply(op, s) {
    return {
      imageEdits: s.imageEdits.flatMap((e) => {
        if (e.id !== op.id) return [e]
        // A static form fill re-typed or moved: removing the edit removes the fill
        if (
          e.staticFill &&
          (e.input.kind === 'transformImage' || e.input.kind === 'replaceImage')
        ) {
          return [
            {
              ...e,
              input: {
                kind: 'deleteImage' as const,
                pageIndex: e.input.pageIndex,
                oldRect: e.input.oldRect,
              },
            },
          ]
        }
        return []
      }),
    }
  },
})

// ── form ──────────────────────────────────────────────────────────────

register({
  name: 'setFormValue',
  touches: ['formEdits'],
  validate(op) {
    const v = obj<FormValueInput>(op.value, 'value')
    str(v.name, 'value.name')
    if (!['text', 'checkbox', 'radio', 'choice'].includes(v.kind))
      throw new GuidedError('value.kind must be text | checkbox | radio | choice')
  },
  apply(op, s) {
    const v = op.value as FormValueInput
    return { formEdits: new Map(s.formEdits).set(v.name, v) }
  },
})

// ── page ──────────────────────────────────────────────────────────────

register({
  name: 'rotatePages',
  touches: ['rotations', 'drawings'],
  validate(op, ctx) {
    if (!Array.isArray(op.pages) || op.pages.length === 0)
      throw new GuidedError('"pages" must list at least one original page index')
    for (const p of op.pages) pageIndex(p, ctx, 'pages[]')
    if (op.dir !== 90 && op.dir !== -90 && op.dir !== 180)
      throw new GuidedError('"dir" must be 90, -90 or 180')
  },
  apply(op, s) {
    const pages = new Set(op.pages as number[])
    const dir = op.dir as 90 | -90 | 180
    const rotations = new Map(s.rotations)
    for (const p of pages) {
      const nv = ((rotations.get(p) ?? 0) + dir + 360) % 360
      if (nv === 0) rotations.delete(p)
      else rotations.set(p, nv)
    }
    if (dir === 180) return { rotations, drawings: s.drawings }
    // Image stamps draw upright, so a quarter turn swaps their displayed width and
    // height: swap the user-space rect about its center to keep the aspect ratio
    const drawings = s.drawings.map((d) => {
      if (d.input.kind !== 'image' || !pages.has(d.input.pageIndex)) return d
      const [x1, y1, x2, y2] = d.input.rect
      const cx = (x1 + x2) / 2
      const cy = (y1 + y2) / 2
      const hw = (x2 - x1) / 2
      const hh = (y2 - y1) / 2
      return { ...d, input: { ...d.input, rect: [cx - hh, cy - hw, cx + hh, cy + hw] as Rect } }
    })
    return { rotations, drawings }
  },
})

register({
  name: 'deletePage',
  touches: ['deleted', 'markups', 'drawings'],
  validate(op, ctx) {
    const p = pageIndex(op.pageIndex, ctx)
    if (ctx.pageCount - ctx.deleted.size <= 1)
      throw new GuidedError(`Cannot delete page ${p}: at least one page must remain`)
  },
  advance(op, ctx) {
    return { deleted: new Set(ctx.deleted).add(op.pageIndex as number) }
  },
  apply(op, s) {
    const p = op.pageIndex as number
    return {
      deleted: new Set(s.deleted).add(p),
      markups: s.markups.filter((m) => m.pageIndex !== p),
      drawings: s.drawings.filter((d) => d.input.pageIndex !== p),
    }
  },
})

register({
  name: 'setPageOrder',
  touches: ['order'],
  validate(op, ctx) {
    if (op.order === null) return
    const order = op.order
    if (!Array.isArray(order) || order.length !== ctx.pageCount)
      throw new GuidedError(
        `"order" must list every original page index exactly once (${ctx.pageCount} pages, deleted ones included) or be null`,
      )
    const seen = new Set<number>()
    for (const p of order) {
      if (typeof p !== 'number' || p < 0 || p >= ctx.pageCount || seen.has(p))
        throw new GuidedError(`"order" must be a permutation of 0..${ctx.pageCount - 1}`)
      seen.add(p)
    }
  },
  apply(op) {
    return { order: op.order as number[] | null }
  },
})

// ── document ──────────────────────────────────────────────────────────

register({
  name: 'setStamps',
  touches: ['stampCfg'],
  validate(op) {
    if (op.cfg !== null) obj<StampConfig>(op.cfg, 'cfg')
  },
  apply(op) {
    const cfg = op.cfg as StampConfig | null
    // A config with neither layer is the cleared state, not a pending edit
    return { stampCfg: cfg && (cfg.wm || cfg.hf) ? cfg : null }
  },
})

register({
  name: 'setMetadata',
  touches: ['metadata'],
  validate(op) {
    if (op.metadata !== null) obj<MetadataInput>(op.metadata, 'metadata')
  },
  apply(op) {
    return { metadata: op.metadata as MetadataInput | null }
  },
})
