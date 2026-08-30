/**
 * Text-family ops (batch 2): setText / setFont / setParagraphFormat, extracted
 * from their IPC handlers. Session/render concerns (autofit resize, RenderSlide
 * rebuilding, undo bookkeeping) stay in the shims; model mutation, link-rel
 * upkeep, superseded-resource cleanup, and level rematerialization live here.
 */
import {
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
  type ElementFontPatch,
  type GroupElement,
  type Paragraph,
  type ParagraphFormatPatch,
  type TextBody,
  type TextElement,
} from '@genoffice/pptx-engine'
import type { EditParagraph } from '../../shared/ipc'
import { applyEditParagraphs, collectParagraphFormatPatches, levelsChanged } from '../edit-text'
import {
  GuidedError,
  register,
  requireFinite,
  requireHexColor,
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
register({
  name: 'setText',
  validate(op, ctx) {
    if (!Array.isArray(op.paragraphs)) {
      throw new GuidedError('op "setText" needs "paragraphs": an EditParagraph array.')
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
      for (const { index: pi, patch } of collectParagraphFormatPatches(paragraphs)) {
        setGroupChildParagraphFormat(slide, groupId, id, patch, [pi])
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
    for (const { index: pi, patch } of collectParagraphFormatPatches(paragraphs)) {
      setElementParagraphFormat(slide, te.id, patch, [pi])
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
    const font = op.font as { color?: unknown; fontSizePt?: unknown }
    if (font.color !== undefined) requireHexColor(font.color, 'setFont', 'font.color')
    if (font.fontSizePt !== undefined) {
      requireFinite(font.fontSizePt, 'setFont', 'font.fontSizePt')
      if (font.fontSizePt <= 0) {
        throw new GuidedError('op "setFont": "font.fontSizePt" must be > 0 (points).')
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
    const font = op.font as ElementFontPatch
    const { index, slide } = resolveSlide(ctx, op)
    let id = String(op.target?.el ?? '')
    let ok: boolean
    if (op.group) {
      const groupId = resolveGroup(op, index, slide.elements)
      id = resolveGroupChildId(slide, groupId, id)
      ok = setGroupChildFont(slide, groupId, id, font)
    } else {
      id = resolveElement(ctx, op).el.id
      ok = setElementFont(slide, id, font)
    }
    if (!ok) {
      throw new GuidedError(`op "setFont": element "${id}" has no editable text to format.`)
    }
    return { op, after: font }
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
    for (const field of [
      'lineSpacingPct',
      'spaceBeforePt',
      'spaceAfterPt',
      'bulletSizePct',
      'bulletHangEmu',
    ]) {
      if (fmt[field] !== undefined) {
        requireFinite(fmt[field], 'setParagraphFormat', `format.${field}`)
        if ((fmt[field] as number) < 0) {
          throw new GuidedError(`op "setParagraphFormat": "format.${field}" must be >= 0.`)
        }
      }
    }
    if (fmt.bulletColor !== undefined) {
      requireHexColor(fmt.bulletColor, 'setParagraphFormat', 'format.bulletColor')
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
    const format = op.format as ParagraphFormatPatch
    const { index, slide } = resolveSlide(ctx, op)
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
