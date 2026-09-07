/**
 * Slides master-view channels — master-open/enter/close, edit-fill/stroke/text
 * transform on master shapes. Placeholders matching the legacy single-file
 * implementation.
 */
import { registerHandle } from '../common/index.js'

export function registerSlidesMasterHandlers(): void {
  registerHandle('slides:master-open', () => ({ ok: true }))
  registerHandle('slides:master-enter', () => ({ ok: true }))
  registerHandle('slides:master-close', () => ({ ok: true }))
  registerHandle('slides:master-delete-element', () => ({ ok: true }))
  registerHandle('slides:master-edit-fill', () => ({ ok: true }))
  registerHandle('slides:master-edit-stroke', () => ({ ok: true }))
  registerHandle('slides:master-edit-text', () => ({ ok: true }))
  registerHandle('slides:master-edit-transform', () => ({ ok: true }))
}
