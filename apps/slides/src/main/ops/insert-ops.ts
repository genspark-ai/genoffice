/**
 * Content-bearing add ops: every additive op carries its full content and
 * style in one op (the generation subset of the vocabulary). New element ids
 * are reported via the record's `created` — additive ops are how identity is
 * born, so the journal always knows what a transaction minted.
 */
import {
  addChart,
  addElement,
  addMedia,
  addModel3d,
  addPicture,
  addSmartArt,
  addTable,
  pasteElements,
  replacePictureBytes,
  type NewChartKind,
  type NewChartOptions,
  type NewElementOptions,
  type Paragraph,
} from '@genoffice/pptx-engine'
import {
  GuidedError,
  register,
  resolveElement,
  resolveSlide,
  type Op,
  type OpRecord,
} from './registry'

function reqRect(op: Op): { x: number; y: number; cx: number; cy: number } {
  const b = op.offset as { x?: unknown; y?: unknown; cx?: unknown; cy?: unknown } | undefined
  if (
    !b ||
    typeof b.x !== 'number' ||
    typeof b.y !== 'number' ||
    typeof b.cx !== 'number' ||
    typeof b.cy !== 'number'
  ) {
    throw new GuidedError(`op "${op.op}" needs "offset": an EMU rect {x, y, cx, cy}.`)
  }
  return { x: b.x, y: b.y, cx: b.cx, cy: b.cy }
}

function reqBytes(op: Op, field = 'bytes'): Uint8Array {
  const v = op[field]
  if (!(v instanceof Uint8Array) || v.length === 0) {
    throw new GuidedError(`op "${op.op}" needs "${field}": non-empty image/media bytes.`)
  }
  return v
}

// ── addElement (textbox / preset shape / line) ──────────────────────────
register({
  name: 'addElement',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    reqRect(op)
    if (typeof op.kind !== 'string' || !op.kind) {
      throw new GuidedError('op "addElement" needs "kind": "textbox" or a preset geometry name.')
    }
  },
  apply(op, ctx): OpRecord {
    const { slide } = resolveSlide(ctx, op)
    const el = addElement(slide, {
      kind: String(op.kind),
      offset: reqRect(op),
      ...(Array.isArray(op.paragraphs) && op.paragraphs.length
        ? { paragraphs: op.paragraphs as Paragraph[] }
        : {}),
      ...(typeof op.fill === 'string' ? { fillColor: op.fill } : {}),
      ...(op.stroke ? { stroke: op.stroke as NewElementOptions['stroke'] } : {}),
      ...(op.bodyPr ? { bodyPr: op.bodyPr as NewElementOptions['bodyPr'] } : {}),
    })
    return { op, created: [el.id] }
  },
})

// ── addPicture (also serves ink strokes: name/descr carry the ink payload) ──
register({
  name: 'addPicture',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    reqRect(op)
    reqBytes(op)
    if (typeof op.ext !== 'string' || !op.ext) {
      throw new GuidedError('op "addPicture" needs "ext": the image extension (png/jpg/…).')
    }
  },
  apply(op, ctx): OpRecord {
    const { slide } = resolveSlide(ctx, op)
    const el = addPicture(ctx.opened, slide, {
      bytes: reqBytes(op),
      ext: String(op.ext),
      offset: reqRect(op),
      ...(typeof op.name === 'string' && op.name ? { name: op.name } : {}),
      ...(typeof op.descr === 'string' && op.descr ? { descr: op.descr } : {}),
    })
    if (!el) {
      throw new GuidedError(
        `op "addPicture": unsupported image format ".${String(op.ext)}" (png/jpg/jpeg/gif/bmp/webp/tif supported).`,
      )
    }
    return { op, created: [el.id] }
  },
})

// ── replacePicture ──────────────────────────────────────────────────────
register({
  name: 'replacePicture',
  validate(op, ctx) {
    resolveElement(ctx, op)
    reqBytes(op)
    if (typeof op.ext !== 'string' || !op.ext) {
      throw new GuidedError('op "replacePicture" needs "ext": the image extension.')
    }
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op)
    const ok = replacePictureBytes(
      ctx.opened,
      slide,
      el.id,
      reqBytes(op),
      String(op.ext),
      op.keepSrcRect === true ? { keepSrcRect: true } : undefined,
    )
    if (!ok) {
      throw new GuidedError(
        `op "replacePicture": element "${el.id}" is not a picture or the format ".${String(op.ext)}" is unsupported.`,
      )
    }
    return { op }
  },
})

// ── addTable / addChart / addSmartArt ───────────────────────────────────
register({
  name: 'addTable',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    reqRect(op)
    if (typeof op.rows !== 'number' || typeof op.cols !== 'number' || op.rows < 1 || op.cols < 1) {
      throw new GuidedError('op "addTable" needs "rows" and "cols" (>= 1).')
    }
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    const r = addTable(ctx.opened, index, {
      rows: op.rows as number,
      cols: op.cols as number,
      offset: reqRect(op),
    })
    if (!r) throw new GuidedError('op "addTable": the table could not be inserted.')
    return { op, created: [r.elementId] }
  },
})

register({
  name: 'addChart',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    reqRect(op)
    if (typeof op.kind !== 'string') throw new GuidedError('op "addChart" needs "kind".')
    if (!Array.isArray(op.categories) || !Array.isArray(op.series)) {
      throw new GuidedError('op "addChart" needs "categories" and "series" arrays.')
    }
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    const r = addChart(ctx.opened, index, {
      kind: op.kind as NewChartKind,
      ...(op.barDir ? { barDir: op.barDir as 'bar' } : {}),
      ...(typeof op.title === 'string' && op.title ? { title: op.title } : {}),
      categories: op.categories as string[],
      series: op.series as NewChartOptions['series'],
      offset: reqRect(op),
    })
    if (!r) throw new GuidedError('op "addChart": the chart could not be inserted.')
    return { op, created: [r.elementId] }
  },
})

register({
  name: 'addSmartArt',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    reqRect(op)
    if (typeof op.layout !== 'string' || !Array.isArray(op.items)) {
      throw new GuidedError('op "addSmartArt" needs "layout" and "items".')
    }
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    const r = addSmartArt(ctx.opened, index, {
      layout: op.layout as Parameters<typeof addSmartArt>[2]['layout'],
      items: op.items as Parameters<typeof addSmartArt>[2]['items'],
      offset: reqRect(op),
    })
    if (!r) throw new GuidedError('op "addSmartArt": the diagram could not be inserted.')
    return { op, created: [r.elementId] }
  },
})

// ── addMedia / addModel3d ───────────────────────────────────────────────
register({
  name: 'addMedia',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    reqRect(op)
    reqBytes(op)
    if (op.kind !== 'video' && op.kind !== 'audio') {
      throw new GuidedError('op "addMedia" needs "kind": "video" or "audio".')
    }
    if (typeof op.ext !== 'string' || !op.ext) throw new GuidedError('op "addMedia" needs "ext".')
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    const added = addMedia(ctx.opened, index, {
      kind: op.kind as 'video' | 'audio',
      bytes: reqBytes(op),
      ext: String(op.ext),
      offset: reqRect(op),
      ...(typeof op.name === 'string' && op.name ? { name: op.name } : {}),
      ...(op.poster ? { poster: op.poster as { bytes: Uint8Array; ext: string } } : {}),
    })
    if (!added) throw new GuidedError('op "addMedia": the media could not be embedded.')
    return { op, created: [added.elementId] }
  },
})

register({
  name: 'addModel3d',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    reqRect(op)
    reqBytes(op)
    if (typeof op.ext !== 'string' || !op.ext) throw new GuidedError('op "addModel3d" needs "ext".')
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    const added = addModel3d(ctx.opened, index, {
      bytes: reqBytes(op),
      ext: String(op.ext),
      ...(op.poster ? { poster: op.poster as { bytes: Uint8Array; ext: string } } : {}),
      offset: reqRect(op),
      ...(typeof op.name === 'string' && op.name ? { name: op.name } : {}),
    })
    if (!added) throw new GuidedError('op "addModel3d": the model could not be embedded.')
    return { op, created: [added.elementId] }
  },
})

// ── pasteElements (paste / duplicate-in-place share the same core) ──────
register({
  name: 'pasteElements',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    if (!Array.isArray(op.items) || op.items.length === 0) {
      throw new GuidedError('op "pasteElements" needs "items": copied element data.')
    }
    if (typeof op.dx !== 'number' || typeof op.dy !== 'number') {
      throw new GuidedError('op "pasteElements" needs "dx"/"dy" (EMU offsets).')
    }
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    const r = pasteElements(ctx.opened, index, op.items as Parameters<typeof pasteElements>[2], {
      dx: op.dx as number,
      dy: op.dy as number,
    })
    if (!r) throw new GuidedError('op "pasteElements": the clipboard items could not be pasted.')
    return { op, created: r.elementIds }
  },
})
