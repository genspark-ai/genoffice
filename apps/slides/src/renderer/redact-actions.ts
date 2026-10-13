/**
 * Withholding content from the model, for the two things a reader can withhold.
 *
 * The two gestures do not share a mechanism, and the difference is not incidental:
 *
 * - A **picture, video or audio clip** has no runs to mark, so it goes through the
 *   op surface (`setRedaction`). That path is already journaled and undoable, so it
 *   is used as-is rather than rebuilt here. It also returns the whole deck, so the
 *   result replaces `slides` the same way any other op transaction does.
 * - **Text** is marked in the editor DOM and left to the overlay's ordinary commit
 *   (EditParagraph[] → slides:edit-text). No op, no own history: pressing away from
 *   the text box is what has always saved an edit, and the mark rides along with it.
 *
 * In both gestures the words or the bytes stay exactly where they were. A mark is a
 * label over the reader's own content, not a replacement for it — losing their data
 * to hide it from a model would be the worse failure.
 *
 * Neither path writes a status string on success: the mark is its own feedback (an
 * underlined run, a marked shape on the canvas), and the locales must not grow a key
 * for it. A failure reports the op layer's own guided message.
 */
import type { ActionCtx } from './action-context'
import type { PictureRenderNode } from '@genoffice/pptx-render'
import { mediaLabelOf, sanitizeLabel } from './ai/redact'
import { applySelectionRedaction, saveEditSelection, selectionRedaction } from './TextEditOverlay'

/** The picture/video/audio shape at `sourceId` on the current slide, if it is one. */
function pictureAt(ctx: ActionCtx, sourceId: string): PictureRenderNode | null {
  const node = ctx.findNodeCtx(sourceId)?.node
  return node?.type === 'picture' ? node : null
}

/**
 * The context-menu item for a shape: a picture, a video or an audio clip.
 *
 * A shape that is already withheld offers to stop withholding it, on the same item:
 * one label has to serve both, since the two differ only in what they do. Clearing is
 * immediate and reversible (the op is journaled and undoable) — the alternative would
 * be a second string for a locale set that must not grow.
 */
export async function toggleElementRedaction(ctx: ActionCtx, sourceId: string): Promise<void> {
  const pic = pictureAt(ctx, sourceId)
  if (!pic) return
  if (mediaLabelOf({ redact: pic.redact }) !== null) {
    await applyElementRedaction(ctx, sourceId, null)
    return
  }
  ctx.setRedactDialog({ sourceId, run: false, seed: pic.name ?? '' })
}

/**
 * The context-menu item for a text selection. Already withheld means "stop
 * withholding"; otherwise the dialog asks what the span stands for.
 */
export function openTextRedaction(ctx: ActionCtx): void {
  if (!ctx.editing && !ctx.editingCell) return
  if (selectionRedaction() !== null) {
    applySelectionRedaction(null)
    return
  }
  // the dialog takes focus; the range is saved and restored when it closes
  saveEditSelection()
  ctx.setRedactDialog({ sourceId: null, run: true, seed: selectedText() })
}

/** Confirm (or clear) from the dialog, then hand the result to whichever path is open. */
export async function applyRedaction(ctx: ActionCtx, label: string | null): Promise<void> {
  const target = ctx.redactDialog
  if (!target) return
  const clean = label === null ? null : sanitizeLabel(label)
  if (target.run) {
    // Applied to the DOM, not committed here: the overlay commits the whole edit the
    // next time focus leaves it, exactly as it does for any other change.
    applySelectionRedaction(clean)
    ctx.setRedactDialog(null)
    return
  }
  ctx.setRedactDialog(null)
  if (!target.sourceId) return
  await applyElementRedaction(ctx, target.sourceId, clean)
}

/** Withhold a picture/video/audio shape through the op surface, or clear it with null. */
async function applyElementRedaction(
  ctx: ActionCtx,
  sourceId: string,
  label: string | null,
): Promise<void> {
  const result = await window.slidesApi.applyTxn({
    ops: [{ op: 'setRedaction', target: { slide: ctx.current, el: sourceId }, label }],
  })
  if (!result?.applied) {
    // the op layer's own guided wording; a refusal it explains is the useful message
    const reason = result?.failures?.[0]?.error
    if (reason) ctx.setStatus(reason)
    return
  }
  // the transaction may touch any slide, so the rebuilt deck replaces the old one
  if (result.slides) {
    ctx.applyDeck(result.slides, Math.max(0, Math.min(ctx.current, result.slides.length - 1)))
  }
}

/** The text the editor selection covers, as the dialog's starting label. */
function selectedText(): string {
  const sel = window.getSelection()
  if (!sel?.rangeCount) return ''
  return sel.getRangeAt(0).toString().trim().slice(0, 24)
}
