/**
 * Delete / Backspace on a multi-cell selection must clear every selected
 * cell (Excel / Google Sheets). Univer 0.25 binds Backspace to
 * "delete-and-start-editing" (SetCellEditVisibleOperation) which only
 * affects the active cell, and on some hosts that same binding races
 * Delete. Intercept at window-capture — before Univer's shortcut
 * dispatcher — and route to sheet.command.clear-selection-content.
 *
 * Univer parks grid focus on a hidden contenteditable host, so a naive
 * "skip every contentEditable" check would never intercept the grid.
 * Native inputs (including Univer find/replace and rule panels) always
 * skip; remaining contenteditables skip unless they sit in the sheet
 * canvas rather than Univer chrome or app chrome.
 */

export const CLEAR_SELECTION_CONTENT_COMMAND = 'sheet.command.clear-selection-content'

const NATIVE_FIELD_SELECTOR = 'input, textarea, select'
const CONTENT_EDITABLE_SELECTOR = '[contenteditable="true"]'
const SHEET_CONTAINER_SELECTOR = '#univer-container'
export const SKIP_HOST_SELECTOR = [
  '[data-u-comp="formula-bar"]',
  '[data-u-comp="input"]',
  '[data-u-comp="textarea"]',
  '[data-u-comp="panel"]',
  '[data-u-comp="panel-field"]',
  '[data-u-comp="cell-popup"]',
  '[data-u-comp="defined-name"]',
  '[data-u-comp="defined-name-container"]',
  '[data-u-comp="select"]',
  '[data-u-comp="multiple-select"]',
  '[data-u-comp="sheets-dropdown-list"]',
  '[data-u-comp="gallery"]',
  '.shape-editable',
  '.chart-editor',
  '.dialog-backdrop',
  '[role="dialog"]',
].join(', ')

export function isClearSelectionHotkey(
  event: Pick<ClearSelectionKeyEvent, 'key' | 'metaKey' | 'ctrlKey' | 'altKey' | 'shiftKey'>,
): boolean {
  if (event.key !== 'Delete' && event.key !== 'Backspace') return false
  return !event.metaKey && !event.ctrlKey && !event.altKey && !event.shiftKey
}

export interface ClearSelectionKeyEvent {
  readonly key: string
  readonly metaKey: boolean
  readonly ctrlKey: boolean
  readonly altKey: boolean
  readonly shiftKey: boolean
  readonly defaultPrevented: boolean
  readonly isComposing: boolean
  readonly target: { closest(selector: string): unknown } | EventTarget | null
}

export function shouldInterceptClearSelection(
  event: ClearSelectionKeyEvent,
  isCellEditing: boolean,
): boolean {
  if (!isClearSelectionHotkey(event)) return false
  if (event.defaultPrevented) return false
  if (event.isComposing) return false
  if (isCellEditing) return false
  const target = event.target
  if (!hasClosest(target)) return true
  if (target.closest(SKIP_HOST_SELECTOR)) return false
  // Find/replace, data-validation, CF, and app chrome all use native fields.
  if (target.closest(NATIVE_FIELD_SELECTOR)) return false
  // App chrome (AI composer) is contenteditable outside the sheet container.
  if (target.closest(CONTENT_EDITABLE_SELECTOR) && !target.closest(SHEET_CONTAINER_SELECTOR)) {
    return false
  }
  return true
}

function hasClosest(value: ClearSelectionKeyEvent['target']): value is Element {
  return value != null && typeof (value as Element).closest === 'function'
}
