/**
 * Slides read-only query channels — get-render-slides, get-animations,
 * get-chart-data, get-comments, get-header-footer, get-layouts, etc.
 * Most of these return empty arrays / null placeholders.
 */
import { registerHandle } from '../common/index.js'

export function registerSlidesStateHandlers(): void {
  registerHandle('slides:get-render-slides', () => [])
  registerHandle('slides:get-animations', () => [])
  registerHandle('slides:get-chart-data', () => ({}))
  registerHandle('slides:get-comments', () => [])
  registerHandle('slides:get-header-footer', () => ({ enabled: false }))
  registerHandle('slides:get-layouts', () => [])
  registerHandle('slides:get-link', () => null)
  registerHandle('slides:get-notes', () => '')
  registerHandle('slides:get-sections', () => [])
  registerHandle('slides:get-shape-keys', () => [])
  registerHandle('slides:get-slide-links', () => [])
  registerHandle('slides:get-slide-size', () => ({ width: 960, height: 540 }))
  registerHandle('slides:get-run-links', () => [])
  registerHandle('slides:has-slide-clipboard', () => false)
  registerHandle('slides:font-catalog', () => [])
  registerHandle('slides:font-missing', () => [])
  registerHandle('slides:chart-color-schemes', () => [])
  registerHandle('slides:clipboard-external', () => ({}))
  registerHandle('slides:clipboard-probe', () => ({}))
  registerHandle('slides:media-data', () => ({}))
  registerHandle('slides:native-clipboard', () => ({}))
  registerHandle('slides:table-structure', () => ({}))
  registerHandle('slides:private-font-data', () => ({}))
  registerHandle('slides:private-font-faces', () => [])
  registerHandle('slides:cloud-gen-status', () => ({ status: 'idle' }))
}
