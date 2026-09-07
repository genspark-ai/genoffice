/**
 * Slides element-manipulation channels — add/edit/delete/transform/copy/paste
 * of slides, elements, charts, tables, text, ink, media, etc. All handlers
 * remain placeholders returning `{ ok: true }` shapes; Phase 1.2/1.3 will
 * wire them to real engines.
 */
import { registerHandle } from '../common/index.js'

export function registerSlidesElementHandlers(): void {
  // ----- slide-level mutations ---------------------------------------------
  registerHandle('slides:add-blank-slide', () => ({ ok: true, slideId: `slide-${Date.now()}` }))
  registerHandle('slides:add-slide', () => ({ ok: true, slideId: `slide-${Date.now()}` }))
  registerHandle('slides:add-slide-with-layout', () => ({ ok: true, slideId: `slide-${Date.now()}` }))
  registerHandle('slides:delete-slide', () => ({ ok: true }))
  registerHandle('slides:copy-slide', () => ({ ok: true, slideId: `slide-${Date.now()}` }))
  registerHandle('slides:paste-slide', () => ({ ok: true }))
  registerHandle('slides:repaste-slide', () => ({ ok: true }))
  registerHandle('slides:move-slide', () => ({ ok: true }))
  registerHandle('slides:is-dirty', () => false)

  // ----- element add / edit / delete ---------------------------------------
  registerHandle('slides:add-chart', () => ({ ok: true, chartId: `chart-${Date.now()}` }))
  registerHandle('slides:add-image-bytes', () => ({ ok: true, imageId: `image-${Date.now()}` }))
  registerHandle('slides:add-table', () => ({ ok: true, tableId: `table-${Date.now()}` }))
  registerHandle('slides:add-text', () => ({ ok: true, elementId: `text-${Date.now()}` }))
  registerHandle('slides:add-element', () => ({ ok: true, elementId: `element-${Date.now()}` }))
  registerHandle('slides:add-media-bytes', () => ({ ok: true }))
  registerHandle('slides:add-ink', () => ({ ok: true }))
  registerHandle('slides:add-smartart', () => ({ ok: true }))

  registerHandle('slides:edit-text', () => ({ ok: true }))
  registerHandle('slides:edit-background', () => ({ ok: true }))
  registerHandle('slides:edit-chart', () => ({ ok: true }))
  registerHandle('slides:edit-connector-endpoints', () => ({ ok: true }))
  registerHandle('slides:edit-fill', () => ({ ok: true }))
  registerHandle('slides:edit-image-fill', () => ({ ok: true }))
  registerHandle('slides:edit-picture-opacity', () => ({ ok: true }))
  registerHandle('slides:edit-picture-src-rect', () => ({ ok: true }))
  registerHandle('slides:edit-stroke', () => ({ ok: true }))
  registerHandle('slides:edit-table-cell', () => ({ ok: true }))
  registerHandle('slides:edit-table-style', () => ({ ok: true }))
  registerHandle('slides:edit-transform', () => ({ ok: true }))
  registerHandle('slides:batch-edit-transform', () => ({ ok: true }))
  registerHandle('slides:delete-element', () => ({ ok: true }))
  registerHandle('slides:duplicate-elements', () => ({ ok: true }))
  registerHandle('slides:flip-elements', () => ({ ok: true }))
  registerHandle('slides:reorder-element', () => ({ ok: true }))
  registerHandle('slides:insert-image', () => ({ ok: true }))
  registerHandle('slides:set-element-font', () => ({ ok: true }))
  registerHandle('slides:set-element-paragraph-format', () => ({ ok: true }))
  registerHandle('slides:set-table-cell-anchor', () => ({ ok: true }))
  registerHandle('slides:set-table-col-width', () => ({ ok: true }))
  registerHandle('slides:set-table-row-height', () => ({ ok: true }))
  registerHandle('slides:table-merge', () => ({ ok: true }))
  registerHandle('slides:replace-picture-bytes', () => ({ ok: true }))
  registerHandle('slides:apply-edit-script', () => ({ ok: true }))
  registerHandle('slides:apply-header-footer', () => ({ ok: true }))
  registerHandle('slides:apply-theme', () => ({ ok: true }))
  registerHandle('slides:apply-txn', () => ({ ok: true }))

  // ----- element set / link / transition -----------------------------------
  registerHandle('slides:set-advance-times', () => ({ ok: true }))
  registerHandle('slides:set-animations', () => ({ ok: true }))
  registerHandle('slides:set-hidden', () => ({ ok: true }))
  registerHandle('slides:set-link', () => ({ ok: true }))
  registerHandle('slides:set-notes', () => ({ ok: true }))
  registerHandle('slides:set-sections', () => ({ ok: true }))
  registerHandle('slides:set-slide-layout', () => ({ ok: true }))
  registerHandle('slides:set-slide-size', () => ({ ok: true }))
  registerHandle('slides:set-transition', () => ({ ok: true }))
  registerHandle('slides:get-transition', () => ({ type: 'none', duration: 0 }))

  // ----- clipboard ops ------------------------------------------------------
  registerHandle('slides:copy-elements', () => ({ ok: true }))
  registerHandle('slides:paste-elements', () => ({ ok: true }))
  registerHandle('slides:group-elements', () => ({ ok: true, groupId: `group-${Date.now()}` }))
  registerHandle('slides:ungroup-element', () => ({ ok: true }))

  // ----- undo/redo ---------------------------------------------------------
  registerHandle('slides:undo', () => ({ ok: true }))
  registerHandle('slides:redo', () => ({ ok: true }))
  registerHandle('slides:history-batch-begin', () => ({ ok: true }))
  registerHandle('slides:history-batch-end', () => ({ ok: true }))
  registerHandle('slides:ai-snapshot-restore', () => ({ ok: true }))

  // ----- find / replace ----------------------------------------------------
  registerHandle('slides:find-replace', () => ({ ok: true, count: 0 }))

  // ----- presenter / display ----------------------------------------------
  registerHandle('slides:presenter-start', () => ({ ok: true }))
  registerHandle('slides:presenter-end', () => ({ ok: true }))
  registerHandle('slides:presenter-swap', () => ({ ok: true }))
  registerHandle('slides:audience-ready', () => ({ ok: true }))
  registerHandle('slides:show-fullscreen', () => ({ ok: true }))

  // ----- sections ----------------------------------------------------------
  registerHandle('slides:add-section', () => ({ ok: true, sectionId: `section-${Date.now()}` }))
  registerHandle('slides:rename-section', () => ({ ok: true }))
  registerHandle('slides:remove-section', () => ({ ok: true }))
  registerHandle('slides:move-section', () => ({ ok: true }))

  // ----- comments ----------------------------------------------------------
  registerHandle('slides:add-comment', () => ({ ok: true, commentId: `comment-${Date.now()}` }))
  registerHandle('slides:delete-comment', () => ({ ok: true }))
}
