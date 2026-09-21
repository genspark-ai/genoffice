/**
 * Text-family ops (batch 2): setText / setFont / setParagraphFormat, extracted
 * from their IPC handlers. Session/render concerns (autofit resize, RenderSlide
 * rebuilding, undo bookkeeping) stay in the shims; model mutation, link-rel
 * upkeep, superseded-resource cleanup, and level rematerialization live here.
 */
import {
  DEFAULT_FONT_SIZE_PT,
  FONT_SIZE_PT_MIN,
  FONT_SIZE_PT_MAX,
  applyFontSizeStep,
  type FontSizeStep,
} from '../font-size'
import {
  addImageMediaAndRel,
  cleanupSupersededSlideResources,
  DEFAULT_BODY_INSETS,
  ensureRunLinkRels,
  findGroupChild,
  groupChildSlices,
  isConnectorXml,
  materializeSlide,
  patchGroupChildText,
  patchSlideXml,
  setElementFont,
  setElementParagraphFormat,
  setGroupChildFont,
  setGroupChildParagraphFormat,
  type OpenedPptx,
  type Slide,
  type ElementFontPatch,
  type GroupElement,
  type Paragraph,
  type ParagraphFormatPatch,
  type TextBody,
  type TextElement,
} from '@genoffice/pptx-engine'
import type { EditParagraph } from '../types'
import {
  applyEditParagraphs,
  collectParagraphFormatPatches,
  levelsChanged,
  type EditParagraphFormatPatch,
} from '../edit-text'
import {
  GuidedError,
  register,
  requireFinite,
  requireHexColor,
  requireLinkTarget,
  resolveElement,
  resolveGroup,
  resolveGroupChildId,
  resolveSlide,
  type OpRecord,
} from './registry'

function plainText(paragraphs: Paragraph[] | undefined): string {
  if (!paragraphs) return ''
  return paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n')
}

// ── First text on a shape that has no <p:txBody> ────────────────────────
// The engine injects the element into the bytes on write-back; the model side just
// has to describe the body PowerPoint would have created.

/**
 * Only an autoshape gets PowerPoint's centered authoring defaults. A placeholder
 * inherits anchor and alignment from the layout/master — in a survey of 421
 * PowerPoint-authored decks ~90% of placeholders carry neither attribute, so baking
 * them in would override the template. A text box stays top-left.
 */
function isAutoShape(el: TextElement): boolean {
  return el.type === 'shape' && !el.placeholder && !el.txBox
}

/** Insets mirror the parser's <a:bodyPr> defaults, so the live render and the
 *  reopened file agree; the anchor is the vertical half of an autoshape's centering. */
function createTextBody(autoShape: boolean): TextBody {
  return {
    paragraphs: [],
    insets: { ...DEFAULT_BODY_INSETS },
    ...(autoShape ? { anchor: 'middle' as const } : {}),
  }
}

/** Connectors cannot hold text: CT_Connector has no txBody child, so anything typed
 *  into one would be dropped on save (or reopen as a repaired file). */
function refuseConnector(bytes: string, what: string): void {
  if (isConnectorXml(bytes)) {
    throw new GuidedError(`op "setText": ${what} is a connector, which cannot hold text.`)
  }
}

/** Group children carry an empty byte anchor (the group saves as one blob), so a
 *  child's bytes come from the group's slice table. */
function groupChildBytes(grp: GroupElement, child: TextElement): string {
  return groupChildSlices(grp.anchor.originalXml).find((s) => s.nvId === child.nvId)?.xml ?? ''
}

/** Horizontal half of an autoshape's centering; an alignment the caller asked for wins. */
function centerParagraphs(paragraphs: Paragraph[]): void {
  for (const p of paragraphs) if (!p.align) p.align = 'center'
}

// ── setText ─────────────────────────────────────────────────────────────
// Run-level rich-text rebuild (srcPara/srcRun back-tracing preserves unedited
// formatting). Level changes are rematerialized here (they change inherited
// defaults); the record's after.levelDirty tells the shim to skip autofit.
/**
 * Picture-bullet source → media part + slide rel, once per distinct image within a
 * transaction; the engine patch carries only the landed rel.
 */
function landBulletImage(
  opened: OpenedPptx,
  slide: Slide,
  patch: EditParagraphFormatPatch,
  landed: Map<string, { rid: string; mediaPath: string }>,
  opName: string,
): ParagraphFormatPatch {
  const { bulletImage, ...rest } = patch
  if (!bulletImage || rest.bullet !== 'blip') return rest
  const key = `${bulletImage.ext}|${bulletImage.base64}`
  let added = landed.get(key)
  if (!added) {
    const r = addImageMediaAndRel(
      opened,
      slide,
      Buffer.from(bulletImage.base64, 'base64'),
      bulletImage.ext,
    )
    if (!r) {
      throw new GuidedError(
        `op "${opName}": unsupported picture bullet format ".${bulletImage.ext}".`,
      )
    }
    added = r
    landed.set(key, added)
  }
  return { ...rest, bulletBlip: { mediaRef: added.mediaPath, blipEmbedId: added.rid } }
}

register({
  name: 'setText',
  validate(op, ctx) {
    if (!Array.isArray(op.paragraphs)) {
      throw new GuidedError('op "setText" needs "paragraphs": an EditParagraph array.')
    }
    for (const p of op.paragraphs as Array<{ runs?: Array<{ link?: unknown }> }>) {
      for (const r of p?.runs ?? []) {
        if (r && r.link !== undefined) requireLinkTarget('setText', r.link, 'runs[].link')
      }
    }
    if (op.group) {
      const { index, slide } = resolveSlide(ctx, op)
      resolveGroup(op, index, slide.elements)
      return
    }
    resolveElement(ctx, op, { types: ['text', 'shape'], allowPart: true })
  },
  apply(op, ctx): OpRecord {
    const paragraphs = op.paragraphs as EditParagraph[]
    const { index, slide } = resolveSlide(ctx, op, { allowPart: true })
    if (op.group) {
      const groupId = resolveGroup(op, index, slide.elements)
      const id = resolveGroupChildId(slide, groupId, String(op.target?.el ?? ''))
      const found = findGroupChild(slide, groupId, id)
      const child = found?.child
      if (!child || (child.type !== 'text' && child.type !== 'shape')) {
        throw new GuidedError(`op "setText": no text child "${id}" in group "${groupId}".`)
      }
      const textChild = child as TextElement
      const freshBody = !textChild.text
      let autoShape = false
      if (freshBody) {
        refuseConnector(groupChildBytes(found!.grp, textChild), `group child "${id}"`)
        autoShape = isAutoShape(textChild)
        textChild.text = createTextBody(autoShape)
      }
      const childBody = textChild.text!
      const previousXml = patchSlideXml(slide)
      const before = plainText(childBody.paragraphs)
      childBody.paragraphs = applyEditParagraphs(childBody.paragraphs, paragraphs)
      if (autoShape) centerParagraphs(childBody.paragraphs)
      ensureRunLinkRels(ctx.opened, index, childBody.paragraphs)
      if (!patchGroupChildText(slide, groupId, textChild)) {
        // Executor snapshot rollback restores the already-mutated model
        throw new GuidedError(
          `op "setText": the child slice for "${id}" could not be located inside group "${groupId}".`,
        )
      }
      const landed = new Map<string, { rid: string; mediaPath: string }>()
      for (const { index: pi, patch } of collectParagraphFormatPatches(paragraphs)) {
        const fmt = landBulletImage(ctx.opened, slide, patch, landed, 'setText')
        setGroupChildParagraphFormat(slide, groupId, id, fmt, [pi])
      }
      cleanupSupersededSlideResources(ctx.opened, slide, previousXml, patchSlideXml(slide))
      return { op, before, after: { levelDirty: false } }
    }
    const { el } = resolveElement(ctx, op, { types: ['text', 'shape'], allowPart: true })
    const te = el as TextElement
    const freshBody = !te.text
    let autoShape = false
    if (freshBody) {
      refuseConnector(te.anchor.originalXml, `element "${te.id}"`)
      autoShape = isAutoShape(te)
      te.text = createTextBody(autoShape)
    }
    const body = te.text!
    const previousXml = patchSlideXml(slide)
    const before = plainText(body.paragraphs)
    const levelDirty = levelsChanged(body.paragraphs, paragraphs)
    body.paragraphs = applyEditParagraphs(body.paragraphs, paragraphs)
    if (autoShape) centerParagraphs(body.paragraphs)
    // Link rels and level rematerialization are deck-slide concerns (rels live on
    // the slide part; levels resolve against the chrome being edited) — skip on parts
    if (index >= 0) ensureRunLinkRels(ctx.opened, index, body.paragraphs)
    te.dirty = true
    const landed = new Map<string, { rid: string; mediaPath: string }>()
    for (const { index: pi, patch } of collectParagraphFormatPatches(paragraphs)) {
      const fmt = landBulletImage(ctx.opened, slide, patch, landed, 'setText')
      setElementParagraphFormat(slide, te.id, fmt, [pi])
    }
    cleanupSupersededSlideResources(ctx.opened, slide, previousXml, patchSlideXml(slide))
    if (levelDirty && index >= 0) {
      // Level changes affect inheritance (font size/bullet/indent take master defaults by lvl); bake into bytes then reparse
      te.dirtyPPr = { ...te.dirtyPPr, level: true, indents: true }
      materializeSlide(ctx.opened, index)
    }
    return { op, before, after: { levelDirty } }
  },
})

// ── setFont ─────────────────────────────────────────────────────────────
// Wholesale font/format change on one element (text/shape/table); the engine
// decides acceptance so semantics can't drift from the legacy handler. Multi-
// element selections are N ops in one per_op transaction (shim).
register({
  name: 'setFont',
  validate(op, ctx) {
    if (typeof op.font !== 'object' || op.font === null) {
      throw new GuidedError('op "setFont" needs "font": an ElementFontPatch object.')
    }
    const font = op.font as { color?: unknown; fontSizePt?: unknown; fontSizeStep?: unknown }
    if (font.color !== undefined) requireHexColor(font.color, 'setFont', 'font.color')
    if (font.fontSizePt !== undefined) {
      requireFinite(font.fontSizePt, 'setFont', 'font.fontSizePt')
      if (font.fontSizePt < FONT_SIZE_PT_MIN || font.fontSizePt > FONT_SIZE_PT_MAX) {
        throw new GuidedError(
          `op "setFont": "font.fontSizePt" must be ${FONT_SIZE_PT_MIN}..${FONT_SIZE_PT_MAX} (points).`,
        )
      }
    }
    if (font.fontSizeStep !== undefined) {
      const step = font.fontSizeStep as Partial<FontSizeStep> | null
      if (
        font.fontSizePt !== undefined ||
        !step ||
        (step.dir !== 1 && step.dir !== -1) ||
        (step.mode !== 'ladder' && step.mode !== 'point')
      ) {
        throw new GuidedError(
          'op "setFont": "font.fontSizeStep" must be { dir: 1 | -1, mode: "ladder" | "point" } and cannot be combined with "font.fontSizePt".',
        )
      }
    }
    if (op.group) {
      const { index, slide } = resolveSlide(ctx, op)
      resolveGroup(op, index, slide.elements)
      return
    }
    resolveElement(ctx, op)
  },
  apply(op, ctx): OpRecord {
    const { fontSizeStep, ...font } = op.font as ElementFontPatch & { fontSizeStep?: FontSizeStep }
    const sizeMap = fontSizeStep
      ? (pt: number | undefined) => applyFontSizeStep(pt ?? DEFAULT_FONT_SIZE_PT, fontSizeStep)
      : undefined
    const { index, slide } = resolveSlide(ctx, op)
    let id = String(op.target?.el ?? '')
    let ok: boolean
    if (op.group) {
      const groupId = resolveGroup(op, index, slide.elements)
      id = resolveGroupChildId(slide, groupId, id)
      ok = setGroupChildFont(slide, groupId, id, font, sizeMap)
    } else {
      id = resolveElement(ctx, op).el.id
      ok = setElementFont(slide, id, font, sizeMap)
    }
    if (!ok) {
      throw new GuidedError(`op "setFont": element "${id}" has no editable text to format.`)
    }
    return { op, after: op.font }
  },
})

// ── setParagraphFormat ──────────────────────────────────────────────────
register({
  name: 'setParagraphFormat',
  validate(op, ctx) {
    if (typeof op.format !== 'object' || op.format === null) {
      throw new GuidedError(
        'op "setParagraphFormat" needs "format": a ParagraphFormatPatch object.',
      )
    }
    const fmt = op.format as Record<string, unknown>
    // schema ranges: ST_TextSpacingPercent / ST_TextSpacingPoint / ST_TextBulletSizePercent / ST_TextMargin
    const ranges: Record<string, [number, number]> = {
      lineSpacingPct: [0, 13200],
      spaceBeforePt: [0, 1584],
      spaceAfterPt: [0, 1584],
      bulletSizePct: [25, 400],
      bulletHangEmu: [0, 51206400],
    }
    for (const [field, [min, max]] of Object.entries(ranges)) {
      if (fmt[field] !== undefined) {
        requireFinite(fmt[field], 'setParagraphFormat', `format.${field}`)
        const v = fmt[field] as number
        if (v < min || v > max) {
          throw new GuidedError(
            `op "setParagraphFormat": "format.${field}" must be ${min}..${max}.`,
          )
        }
      }
    }
    if (fmt.bulletColor !== undefined) {
      requireHexColor(fmt.bulletColor, 'setParagraphFormat', 'format.bulletColor')
    }
    for (const field of ['bulletChar', 'bulletFont', 'numType']) {
      if (fmt[field] !== undefined && typeof fmt[field] !== 'string') {
        throw new GuidedError(`op "setParagraphFormat": "format.${field}" must be a string.`)
      }
    }
    if (fmt.startAt !== undefined) {
      requireFinite(fmt.startAt, 'setParagraphFormat', 'format.startAt')
      if ((fmt.startAt as number) < 1 || !Number.isInteger(fmt.startAt)) {
        throw new GuidedError('op "setParagraphFormat": "format.startAt" must be an integer >= 1.')
      }
    }
    if (fmt.bullet === 'blip') {
      const img = fmt.bulletImage as { base64?: unknown; ext?: unknown } | undefined
      if (typeof img?.base64 !== 'string' || typeof img.ext !== 'string') {
        throw new GuidedError(
          'op "setParagraphFormat": bullet "blip" needs "format.bulletImage": { base64, ext }.',
        )
      }
    }
    if (fmt.rtl !== undefined && typeof fmt.rtl !== 'boolean') {
      throw new GuidedError('op "setParagraphFormat": "format.rtl" must be a boolean.')
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
    const format = landBulletImage(
      ctx.opened,
      slide,
      op.format as EditParagraphFormatPatch,
      new Map(),
      'setParagraphFormat',
    )
    let id = String(op.target?.el ?? '')
    let ok: boolean
    if (op.group) {
      const groupId = resolveGroup(op, index, slide.elements)
      id = resolveGroupChildId(slide, groupId, id)
      ok = setGroupChildParagraphFormat(slide, groupId, id, format)
    } else {
      id = resolveElement(ctx, op).el.id
      ok = setElementParagraphFormat(slide, id, format)
    }
    if (!ok) {
      throw new GuidedError(`op "setParagraphFormat": element "${id}" has no paragraphs to format.`)
    }
    return { op, after: format }
  },
})
