/**
 * Slide- and deck-level ops: slide lifecycle (add/duplicate/paste/delete/
 * move/layout), backgrounds, transitions/animations/hidden, sections,
 * header-footer, notes/comments, find-replace, and the deck page size.
 * File dialogs, OS clipboard state, and RenderSlide rebuilding stay in the
 * IPC shims. applyTheme's document change (entry surgery) is an op; the
 * whole-deck reparse that refreshes resolved colors afterwards is a derived
 * view rebuild, not an edit, and stays in the shim.
 */
import {
  addPicture,
  addSection,
  addSlideComment,
  applyHeaderFooter,
  applyThemeToArchive,
  commitSaved,
  materializeSlide,
  remapDeckColors,
  reparseDeck,
  deleteSlide,
  deleteSlideComment,
  duplicateSlide,
  insertBlankSlide,
  insertSlideWithLayout,
  moveSection,
  moveSlide,
  pasteSlide,
  removeSection,
  renameSection,
  replaceAllInDeck,
  resetSlideBackground,
  resetSlideLayout,
  setSections,
  setSlideAdvanceTime,
  setSlideAnimations,
  setSlideBackground,
  setSlideBackgroundImage,
  setSlideBgGraphicsHidden,
  setSlideHidden,
  setSlideLayout,
  setSlideNotes,
  setSlideSize,
  setSlideTransition,
  elementSpid,
  type SectionInfo,
  type ThemeSpec,
  type SlideAnimation,
  type SlideTransitionKind,
  type TextElement,
} from '@genoffice/pptx-engine'
import { GuidedError, register, resolveSlide, type Op, type OpRecord } from './registry'

// ── slide lifecycle ─────────────────────────────────────────────────────

register({
  name: 'deleteSlide',
  validate(op, ctx) {
    resolveSlide(ctx, op)
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    if (!deleteSlide(ctx.opened, index)) {
      throw new GuidedError(`op "deleteSlide": slide ${index} could not be removed.`)
    }
    return { op, before: { index } }
  },
})

register({
  name: 'duplicateSlide',
  validate(op, ctx) {
    resolveSlide(ctx, op)
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    const slide = duplicateSlide(ctx.opened, index, { clearText: op.clearText === true })
    if (!slide)
      throw new GuidedError(`op "duplicateSlide": slide ${index} could not be duplicated.`)
    return { op, after: { index: index + 1 } }
  },
})

register({
  name: 'addBlankSlide',
  validate(op, ctx) {
    resolveSlide(ctx, op)
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    if (!insertBlankSlide(ctx.opened, index)) {
      throw new GuidedError(
        `op "addBlankSlide": a blank slide could not be inserted after ${index}.`,
      )
    }
    return { op, after: { index: index + 1 } }
  },
})

register({
  name: 'addSlideWithLayout',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    if (typeof op.layoutPath !== 'string' || !op.layoutPath) {
      throw new GuidedError(
        'op "addSlideWithLayout" needs "layoutPath" (a resolved layout part path).',
      )
    }
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    if (!insertSlideWithLayout(ctx.opened, index, String(op.layoutPath))) {
      throw new GuidedError(`op "addSlideWithLayout": layout "${op.layoutPath}" was not found.`)
    }
    return { op, after: { index: index + 1 } }
  },
})

// Paste a copied slide after afterIndex (raw legacy semantics — the engine
// clamps; picture mode flattens the copied slide's PNG onto the anchor slide).
// bundle/png come from the app clipboard, owned by the shim.
register({
  name: 'pasteSlide',
  validate(op) {
    if (typeof op.afterIndex !== 'number') {
      throw new GuidedError('op "pasteSlide" needs "afterIndex": the slide index to paste after.')
    }
    if (op.mode === 'picture') {
      if (typeof op.png !== 'string' || !op.png) {
        throw new GuidedError('op "pasteSlide" mode "picture" needs "png": base64 slide bitmap.')
      }
      return
    }
    if (typeof op.bundle !== 'object' || op.bundle === null) {
      throw new GuidedError('op "pasteSlide" needs "bundle": a copied slide bundle.')
    }
  },
  apply(op, ctx): OpRecord {
    const afterIndex = op.afterIndex as number
    if (op.mode === 'picture') {
      const { deck } = ctx.opened
      const anchorIndex = Math.min(Math.max(afterIndex, 0), deck.slides.length - 1)
      const slide = deck.slides[anchorIndex]
      if (!slide) throw new GuidedError('op "pasteSlide": no anchor slide for the bitmap.')
      const el = addPicture(ctx.opened, slide, {
        bytes: new Uint8Array(Buffer.from(String(op.png), 'base64')),
        ext: 'png',
        offset: { x: 0, y: 0, cx: deck.size.cx, cy: deck.size.cy },
      })
      if (!el) throw new GuidedError('op "pasteSlide": the slide bitmap could not be placed.')
      return { op, created: [el.id], after: { index: anchorIndex } }
    }
    const slide = pasteSlide(
      ctx.opened,
      afterIndex,
      op.bundle as Parameters<typeof pasteSlide>[2],
      {
        keepSourceFormatting: op.mode === 'source',
      },
    )
    if (!slide) throw new GuidedError('op "pasteSlide": the copied slide could not be pasted.')
    return { op, after: { index: ctx.opened.deck.slides.indexOf(slide) } }
  },
})

register({
  name: 'moveSlide',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    if (typeof op.to !== 'number')
      throw new GuidedError('op "moveSlide" needs "to": a slide index.')
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    if (!moveSlide(ctx.opened, index, op.to as number)) {
      throw new GuidedError(`op "moveSlide": cannot move slide ${index} to ${op.to}.`)
    }
    return { op, before: { index }, after: { index: op.to } }
  },
})

register({
  name: 'setSlideLayout',
  validate(op, ctx) {
    resolveSlide(ctx, op)
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    const ok =
      typeof op.layoutPath === 'string' && op.layoutPath
        ? setSlideLayout(ctx.opened, index, op.layoutPath)
        : resetSlideLayout(ctx.opened, index)
    if (!ok) {
      throw new GuidedError(
        `op "setSlideLayout": layout ${op.layoutPath ? `"${op.layoutPath}"` : '(reset)'} could not be applied to slide ${index}.`,
      )
    }
    return { op, after: op.layoutPath ?? null }
  },
})

register({
  name: 'setSlideSize',
  validate(op) {
    if (typeof op.cx !== 'number' || typeof op.cy !== 'number') {
      throw new GuidedError('op "setSlideSize" needs "cx"/"cy" (EMU).')
    }
  },
  apply(op, ctx): OpRecord {
    const before = { ...ctx.opened.deck.size }
    if (!setSlideSize(ctx.opened, op.cx as number, op.cy as number)) {
      throw new GuidedError('op "setSlideSize": the deck size could not be changed.')
    }
    return { op, before, after: { cx: op.cx, cy: op.cy } }
  },
})

// ── background ──────────────────────────────────────────────────────────

/**
 * Full-page "backdrop" rectangles: design templates often use a text-free solid
 * rectangle covering the whole page as background; changing only the page
 * background would be hidden behind them — so repaint such rectangles along
 * with the background, or make them transparent when the new background is a
 * picture.
 */
function repaintFullBleedBackdrops(
  ctx: { opened: { deck: { size: { cx: number; cy: number } } } },
  slide: { elements: unknown[] },
  fill:
    | { type: 'solid'; color: string }
    | {
        type: 'gradient'
        stops: Array<{ pos: number; color: string }>
        angle?: number
        path?: 'circle'
      }
    | { type: 'none' },
): void {
  const size = ctx.opened.deck.size
  for (const raw of slide.elements) {
    const el = raw as TextElement
    if (el.type !== 'shape' && el.type !== 'text') continue
    const fillType = el.fill?.type
    if (fillType !== 'solid' && fillType !== 'gradient') continue
    if (el.text?.paragraphs.some((p) => p.runs.some((r) => r.text.trim()))) continue
    const { x, y, cx, cy } = el.transform.offset
    const coversX = x <= size.cx * 0.05 && x + cx >= size.cx * 0.95
    const coversY = y <= size.cy * 0.05 && y + cy >= size.cy * 0.95
    if (!coversX || !coversY) continue
    el.fill = fill
    el.dirtyFill = true
  }
}

// One slide per op; multi-slide "apply to all" fans out to N ops in one
// transaction. Image backgrounds report the landed media part in
// after.mediaPath so later ops in the same batch reuse it (one part, N rels).
register({
  name: 'setBackground',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    const kind = op.kind
    if (kind === 'solid') {
      if (typeof op.color !== 'string')
        throw new GuidedError('op "setBackground" solid needs "color".')
    } else if (kind === 'gradient') {
      if (typeof op.from !== 'string' || typeof op.to !== 'string') {
        throw new GuidedError('op "setBackground" gradient needs "from"/"to" colors.')
      }
    } else if (kind === 'image') {
      const source = op.source as { bytes?: unknown; mediaPath?: unknown } | undefined
      if (!source || (!source.bytes && !source.mediaPath)) {
        throw new GuidedError(
          'op "setBackground" image needs "source": {bytes, ext} or {mediaPath}.',
        )
      }
    } else if (kind !== 'reset' && kind !== 'graphics') {
      throw new GuidedError('op "setBackground" needs "kind": solid/gradient/image/reset/graphics.')
    }
  },
  apply(op, ctx): OpRecord {
    const { slide } = resolveSlide(ctx, op)
    if (op.kind === 'solid') {
      setSlideBackground(ctx.opened, slide, String(op.color))
      repaintFullBleedBackdrops(ctx, slide, { type: 'solid', color: String(op.color) })
      return { op, after: { kind: 'solid' } }
    }
    if (op.kind === 'gradient') {
      const stops = [
        { pos: 0, color: String(op.from) },
        { pos: 1, color: String(op.to) },
      ]
      const angle = Math.round(((op.angleDeg as number | undefined) ?? 0) * 60000)
      setSlideBackground(ctx.opened, slide, {
        stops,
        ...(op.radial ? { radial: true } : { angle }),
      })
      repaintFullBleedBackdrops(ctx, slide, {
        type: 'gradient',
        stops,
        ...(op.radial ? { path: 'circle' as const } : { angle }),
      })
      return { op, after: { kind: 'gradient' } }
    }
    if (op.kind === 'image') {
      const used = setSlideBackgroundImage(
        ctx.opened,
        slide,
        op.source as Parameters<typeof setSlideBackgroundImage>[2],
        op.tile === true,
      )
      if (!used) throw new GuidedError('op "setBackground": the image could not be landed.')
      repaintFullBleedBackdrops(ctx, slide, { type: 'none' })
      return { op, after: { kind: 'image', mediaPath: used } }
    }
    if (op.kind === 'reset') {
      resetSlideBackground(ctx.opened, slide)
      return { op, after: { kind: 'reset' } }
    }
    setSlideBgGraphicsHidden(ctx.opened, slide, op.hidden === true)
    return { op, after: { kind: 'graphics', hidden: op.hidden === true } }
  },
})

// ── per-slide toggles ───────────────────────────────────────────────────

register({
  name: 'setHidden',
  validate(op, ctx) {
    resolveSlide(ctx, op)
  },
  apply(op, ctx): OpRecord {
    const { slide } = resolveSlide(ctx, op)
    setSlideHidden(slide, op.hidden === true)
    return { op, after: op.hidden === true }
  },
})

register({
  name: 'setTransition',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    if (typeof op.kind !== 'string') throw new GuidedError('op "setTransition" needs "kind".')
  },
  apply(op, ctx): OpRecord {
    const { slide } = resolveSlide(ctx, op)
    setSlideTransition(slide, op.kind as SlideTransitionKind)
    return { op, after: op.kind }
  },
})

register({
  name: 'setAdvanceTime',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    if (op.ms !== null && typeof op.ms !== 'number') {
      throw new GuidedError('op "setAdvanceTime" needs "ms": milliseconds or null to clear.')
    }
  },
  apply(op, ctx): OpRecord {
    const { slide } = resolveSlide(ctx, op)
    setSlideAdvanceTime(slide, op.ms as number | null)
    return { op, after: op.ms }
  },
})

// sourceId → spid mapping happens here (model concern): animations address
// shapes by spid in <p:timing>, but callers speak element ids.
register({
  name: 'setAnimations',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    if (!Array.isArray(op.items)) throw new GuidedError('op "setAnimations" needs "items".')
  },
  apply(op, ctx): OpRecord {
    const { slide } = resolveSlide(ctx, op)
    const anims: SlideAnimation[] = []
    for (const raw of op.items as Array<Record<string, unknown>>) {
      const el = slide.elements.find((x) => x.id === raw.sourceId)
      const spid = el ? elementSpid(el) : null
      if (spid == null) continue
      anims.push({
        spid,
        effect: raw.effect as SlideAnimation['effect'],
        trigger: raw.trigger as SlideAnimation['trigger'],
        durationMs: Math.max(0, Math.round(raw.durationMs as number)),
        delayMs: Math.max(0, Math.round(raw.delayMs as number)),
        ...(raw.motionPath != null
          ? { motionPath: raw.motionPath as SlideAnimation['motionPath'] }
          : {}),
        ...(raw.paragraph != null ? { paragraph: raw.paragraph as number } : {}),
      })
    }
    setSlideAnimations(slide, anims)
    return { op, after: { count: anims.length } }
  },
})

// ── deck-level ──────────────────────────────────────────────────────────

register({
  name: 'findReplace',
  validate(op) {
    if (typeof op.find !== 'string' || !op.find) {
      throw new GuidedError('op "findReplace" needs "find": a non-empty search string.')
    }
    if (typeof op.replace !== 'string') throw new GuidedError('op "findReplace" needs "replace".')
  },
  apply(op, ctx): OpRecord {
    const { count } = replaceAllInDeck(ctx.opened.deck, String(op.find), String(op.replace), {
      matchCase: op.matchCase as boolean | undefined,
      firstOnly: op.firstOnly as boolean | undefined,
      slideIndex: op.slideIndex as number | undefined,
      elementId: op.elementId as string | undefined,
    })
    if (!count) {
      throw new GuidedError(`op "findReplace": "${op.find}" was not found anywhere in the deck.`)
    }
    return { op, after: { count } }
  },
})

register({
  name: 'applyHeaderFooter',
  validate(op) {
    if (typeof op.settings !== 'object' || op.settings === null) {
      throw new GuidedError('op "applyHeaderFooter" needs "settings".')
    }
  },
  apply(op, ctx): OpRecord {
    if (!applyHeaderFooter(ctx.opened, op.settings as Parameters<typeof applyHeaderFooter>[1])) {
      throw new GuidedError('op "applyHeaderFooter": nothing changed (same settings as before).')
    }
    return { op, after: op.settings }
  },
})

// ── sections ────────────────────────────────────────────────────────────

const sectionOp = (
  name: string,
  applyFn: (op: Op, ctx: { opened: Parameters<typeof setSections>[0] }) => unknown,
  needs?: (op: Op) => void,
): void =>
  register({
    name,
    validate(op) {
      needs?.(op)
    },
    apply(op, ctx): OpRecord {
      const r = applyFn(op, ctx)
      if (!r) throw new GuidedError(`op "${name}": the section change was rejected.`)
      return { op, after: r }
    },
  })

sectionOp(
  'setSections',
  (op, ctx) => {
    setSections(ctx.opened, op.sections as SectionInfo[])
    return true
  },
  (op) => {
    if (!Array.isArray(op.sections)) throw new GuidedError('op "setSections" needs "sections".')
  },
)
sectionOp('addSection', (op, ctx) =>
  addSection(ctx.opened, op.atSlideIndex as number, op.name as string),
)
sectionOp('renameSection', (op, ctx) =>
  renameSection(ctx.opened, op.id as string, op.name as string),
)
sectionOp('removeSection', (op, ctx) =>
  removeSection(ctx.opened, op.id as string, { keepSlides: true }),
)
sectionOp('moveSection', (op, ctx) =>
  moveSection(ctx.opened, op.id as string, op.dir as 'up' | 'down'),
)

// ── notes / comments ────────────────────────────────────────────────────

register({
  name: 'setNotes',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    if (typeof op.text !== 'string') throw new GuidedError('op "setNotes" needs "text".')
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    if (!setSlideNotes(ctx.opened, index, String(op.text))) {
      throw new GuidedError(`op "setNotes": notes for slide ${index} could not be written.`)
    }
    return { op, after: op.text }
  },
})

register({
  name: 'addComment',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    if (typeof op.text !== 'string' || !op.text)
      throw new GuidedError('op "addComment" needs "text".')
    if (typeof op.author !== 'string' || !op.author)
      throw new GuidedError('op "addComment" needs "author".')
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    const added = addSlideComment(ctx.opened, index, {
      author: String(op.author),
      text: String(op.text),
    })
    if (!added)
      throw new GuidedError(`op "addComment": the comment could not be added to slide ${index}.`)
    return { op, after: added }
  },
})

register({
  name: 'deleteComment',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    if (typeof op.authorId !== 'number' || typeof op.idx !== 'number') {
      throw new GuidedError('op "deleteComment" needs "authorId" and "idx".')
    }
  },
  apply(op, ctx): OpRecord {
    const { index } = resolveSlide(ctx, op)
    if (
      !deleteSlideComment(ctx.opened, index, {
        authorId: op.authorId as number,
        idx: op.idx as number,
      })
    ) {
      throw new GuidedError(
        `op "deleteComment": no comment (author ${op.authorId}, #${op.idx}) on slide ${index}.`,
      )
    }
    return { op }
  },
})

// ── applyTheme ──────────────────────────────────────────────────────────
// Theme swap as entry surgery: bake unsaved edits first (dirty slices saved
// later would overwrite the surgery), patch theme parts + remap explicit
// colors, then give background-less slides the theme base color so dark
// themes don't leave white pages. The whole-deck reparse that refreshes
// resolved colors afterwards is a derived view rebuild and stays in the shim.
register({
  name: 'applyTheme',
  validate(op) {
    if (typeof op.name !== 'string' || !op.name) {
      throw new GuidedError('op "applyTheme" needs "name": the theme name.')
    }
    const colors = op.colors
    if (typeof colors !== 'object' || colors === null || Array.isArray(colors)) {
      throw new GuidedError(
        'op "applyTheme" needs "colors": a scheme-slot → "#RRGGBB" record (dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink).',
      )
    }
    for (const [k, v] of Object.entries(colors as Record<string, unknown>)) {
      if (typeof v !== 'string' || !/^#?[0-9a-fA-F]{6}$/.test(v)) {
        throw new GuidedError(`op "applyTheme": colors.${k} must be a "#RRGGBB" string.`)
      }
    }
  },
  apply(op, ctx): OpRecord {
    const spec: ThemeSpec = {
      name: op.name as string,
      colors: op.colors as Record<string, string>,
      ...(typeof op.majorFont === 'string' && op.majorFont ? { majorFont: op.majorFont } : {}),
      ...(typeof op.minorFont === 'string' && op.minorFont ? { minorFont: op.minorFont } : {}),
    }
    commitSaved(ctx.opened)
    const patched = applyThemeToArchive(ctx.opened, spec)
    const remapped = remapDeckColors(ctx.opened, spec)
    let backgrounds = 0
    const lt1 = spec.colors.lt1
    if ((patched > 0 || remapped > 0) && lt1) {
      // The remap is entry surgery only — refresh the in-memory model from the
      // remapped bytes first, or materializing a background below would rewrite
      // exactly those slides from the stale model and undo the remap
      ctx.opened.deck.slides = reparseDeck(ctx.opened).deck.slides
      const slides = ctx.opened.deck.slides
      for (let i = 0; i < slides.length; i++) {
        const s = slides[i]!
        if (s.background) continue
        setSlideBackground(ctx.opened, s, `#${lt1.replace(/^#/, '')}`)
        materializeSlide(ctx.opened, i)
        backgrounds++
      }
    }
    return { op, after: { patched, remapped, backgrounds } }
  },
})
