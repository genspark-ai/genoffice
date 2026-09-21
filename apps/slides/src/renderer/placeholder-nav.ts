/**
 * PowerPoint for Windows: Ctrl+Enter while editing commits the text and moves the
 * caret into the next placeholder in spTree order; past the last one it inserts a
 * slide with the same layout and enters that slide's first placeholder.
 */
import type { RenderNode } from '@genoffice/pptx-render'
import type { ActionCtx } from './action-context'
import type { EditParagraph } from '../shared/ipc'
import { FIT_WIDTH } from './app-constants'
import { isEditableText } from './konva-adapter'
import { paragraphsBlank } from './textbox-insert'

/** Header/footer placeholders are not on PowerPoint's Ctrl+Enter cycle */
const SKIPPED = new Set(['dt', 'ftr', 'sldNum'])

const isPlaceholder = (n: RenderNode): boolean =>
  !n.decoration && isEditableText(n) && !!n.placeholder && !SKIPPED.has(n.placeholder)

export function firstPlaceholderId(nodes: RenderNode[]): string | null {
  return nodes.find(isPlaceholder)?.sourceId ?? null
}

/** Next top-level placeholder after currentId (any node kind); a currentId not in nodes scans from the start. */
export function nextPlaceholderId(nodes: RenderNode[], currentId: string): string | null {
  const from = nodes.findIndex((n) => n.sourceId === currentId) + 1
  return nodes.slice(from).find(isPlaceholder)?.sourceId ?? null
}

export async function nextPlaceholder(
  ctx: ActionCtx,
  paragraphs: EditParagraph[] | null,
): Promise<void> {
  const { editing, slide, current } = ctx
  if (!editing || !slide) return
  if (editing.discardIfEmpty && (!paragraphs || paragraphsBlank(paragraphs))) {
    const updated = await window.slidesApi.deleteElement({
      slideIndex: current,
      sourceId: editing.sourceId,
    })
    if (updated) ctx.applySlide(current, updated)
  } else if (paragraphs) {
    const updated = await window.slidesApi.editText({
      slideIndex: current,
      sourceId: editing.sourceId,
      paragraphs,
      ...(editing.groupId ? { groupId: editing.groupId } : {}),
    })
    if (updated) ctx.applySlide(current, updated)
  }
  const next = nextPlaceholderId(slide.nodes, editing.groupId ?? editing.sourceId)
  if (next) {
    ctx.setSelectedIds([next])
    ctx.setEditing({ sourceId: next })
    return
  }
  const r = await window.slidesApi.addBlankSlide({ sourceIndex: current, fitWidthPx: FIT_WIDTH })
  if (!r) {
    ctx.setEditing(null)
    ctx.setSelectedIds([editing.sourceId])
    return
  }
  const first = firstPlaceholderId(r.slides[r.index]?.nodes ?? [])
  ctx.setSlides(r.slides)
  ctx.setCurrent(r.index)
  ctx.setSelectedSlides([r.index])
  ctx.setSelectedIds(first ? [first] : [])
  ctx.setEditing(first ? { sourceId: first } : null)
  ctx.setDirty(true)
}
