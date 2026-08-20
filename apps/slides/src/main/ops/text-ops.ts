/**
 * Text-family ops (batch 2): setText / setFont / setParagraphFormat, extracted
 * from their IPC handlers. Session/render concerns (autofit resize, RenderSlide
 * rebuilding, undo bookkeeping) stay in the shims; model mutation, link-rel
 * upkeep, superseded-resource cleanup, and level rematerialization live here.
 */
import {
  cleanupSupersededSlideResources,
  ensureRunLinkRels,
  findGroupChild,
  materializeSlide,
  patchGroupChildText,
  patchSlideXml,
  setElementFont,
  setElementParagraphFormat,
  setGroupChildFont,
  setGroupChildParagraphFormat,
  type ElementFontPatch,
  type Paragraph,
  type ParagraphFormatPatch,
  type TextElement,
} from '@genoffice/pptx-engine'
import type { EditParagraph } from '../../shared/ipc'
import { applyEditParagraphs, collectParagraphFormatPatches, levelsChanged } from '../edit-text'
import {
  GuidedError,
  register,
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
      if (!textChild.text) {
        throw new GuidedError(`op "setText": group child "${id}" has no editable text body.`)
      }
      const previousXml = patchSlideXml(slide)
      const before = plainText(textChild.text.paragraphs)
      textChild.text.paragraphs = applyEditParagraphs(textChild.text.paragraphs, paragraphs)
      ensureRunLinkRels(ctx.opened, index, textChild.text.paragraphs)
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
    if (!te.text)
      throw new GuidedError(`op "setText": element "${te.id}" has no editable text body.`)
    const previousXml = patchSlideXml(slide)
    const before = plainText(te.text.paragraphs)
    const levelDirty = levelsChanged(te.text.paragraphs, paragraphs)
    te.text.paragraphs = applyEditParagraphs(te.text.paragraphs, paragraphs)
    // Link rels and level rematerialization are deck-slide concerns (rels live on
    // the slide part; levels resolve against the chrome being edited) — skip on parts
    if (index >= 0) ensureRunLinkRels(ctx.opened, index, te.text.paragraphs)
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
