/**
 * Refresh on Open: like pressing Zotero ▸ Refresh right after a document with
 * Zotero citations opens, so edits made to the items in Zotero (a corrected
 * title, a merged duplicate) reach the citations and the bibliography. It is a
 * per-machine setting in the Zotero tab, on until turned off.
 */
const KEY = 'genoffice.docs.zoteroAutoRefresh'

export function loadZoteroAutoRefresh(): boolean {
  try {
    return localStorage.getItem(KEY) !== 'off'
  } catch {
    return true
  }
}

export function saveZoteroAutoRefresh(on: boolean): void {
  try {
    localStorage.setItem(KEY, on ? 'on' : 'off')
  } catch {
    // storage unavailable: the setting lasts for this window only
  }
}

export interface AutoRefreshInput {
  enabled: boolean
  /** Zotero is installed on this machine */
  zoteroInstalled: boolean
  editable: boolean
  /** the document has Zotero preferences (ZOTERO_PREF) */
  hasDocumentData: boolean
  hasFields: boolean
  /** Zotero fields in notes: a Refresh would drop their works from the bibliography */
  noteFields: boolean
}

/** whether an opened document is refreshed; anything missing skips it without a word */
export function shouldAutoRefresh(input: AutoRefreshInput): boolean {
  return (
    input.enabled &&
    input.zoteroInstalled &&
    input.editable &&
    input.hasDocumentData &&
    input.hasFields &&
    !input.noteFields
  )
}
