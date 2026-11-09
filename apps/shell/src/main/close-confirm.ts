/**
 * Window-close confirmation for multi-document sessions (genoffice#2490).
 *
 * The shell is a single window with one tab per document, so the OS close
 * button closes every tab at once. Closing is only guarded today for unsaved
 * changes, which made an easy misclick lose several clean documents at once.
 * When more than one document tab is open the window close is confirmed with
 * a "Close N documents?" dialog; the user can opt out with
 * "Don't ask again and close", persisted here in app-settings.json.
 */
export const CONFIRM_CLOSE_ALL_KEY = 'confirmCloseAllDocuments'

/** multi-document means a second document tab shares the window with the first */
export const MIN_DOCUMENTS_TO_CONFIRM_CLOSE = 2

/** confirm by default; an explicit `false` in app-settings.json opts out */
export function closeConfirmEnabled(settings: Record<string, unknown>): boolean {
  return settings[CONFIRM_CLOSE_ALL_KEY] !== false
}

/** whether the window close needs the multi-document confirmation */
export function shouldConfirmCloseAll(
  openDocumentCount: number,
  settings: Record<string, unknown>,
): boolean {
  return openDocumentCount >= MIN_DOCUMENTS_TO_CONFIRM_CLOSE && closeConfirmEnabled(settings)
}
