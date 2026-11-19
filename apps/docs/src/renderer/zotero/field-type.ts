/**
 * The citation storage this renderer actually writes is Word fields —
 * `ADDIN ZOTERO_ITEM` instructions between fldChars — but the session talks
 * to Zotero over the LibreOffice integration socket, so Zotero's document
 * preferences name whatever the LibreOffice integration would have written:
 * `ReferenceMark`, or `Bookmark` when the document preferences dialog picked
 * it. Word's plugin accepts `Field` and `Bookmark` but not `ReferenceMark`,
 * and WPS accepts only `Field`, so a file saved with the session's value
 * reopens elsewhere as a citation no plugin can edit.
 *
 * The fix is a two-sided spelling: what reaches disk always says `Field`,
 * because fields are what the file holds, and what the session reads back is
 * mapped to `ReferenceMark` again, because the session is the LibreOffice
 * bridge and its own mismatch check would reject the on-disk value.
 */

/** The storage the renderer writes, and the only value Word and WPS agree on. */
const DISK_FIELD_TYPE = 'Field'

/** What the LibreOffice bridge session expects to read back. */
const SESSION_FIELD_TYPE = 'ReferenceMark'

/** One `<pref name="fieldType">` in either spelling Zotero has used. */
const FIELD_TYPE_PREF_RE =
  /(<pref\s+name="fieldType"\s*>\s*<value>)([^<]*)(<\/value>)|(<pref\s+name="fieldType"\s+value=")([^"]*)(")/g

/**
 * Rewrites every fieldType preference through `map`. Each direction has to be
 * conservative in its own way: the disk form flattens whatever the session
 * reported, while the session form only translates a value that says `Field`,
 * leaving the session's own spellings (and so a round trip) untouched.
 */
function remap(data: string, map: (value: string) => string): string {
  return data.replace(
    FIELD_TYPE_PREF_RE,
    (match, open, element, close, attrOpen, attr, attrClose) => {
      if (attrOpen !== undefined) return `${attrOpen}${map(attr)}${attrClose}`
      return `${open}${map(element)}${close}`
    },
  )
}

/** The preferences as they belong in the saved file: whatever the session
 * reported is rewritten to the storage the file actually carries. */
export function fieldTypeForDisk(data: string): string {
  return remap(data, (value) => (value === DISK_FIELD_TYPE ? value : DISK_FIELD_TYPE))
}

/** The preferences as the LibreOffice bridge session expects them: the
 * on-disk `Field` is mapped back before Zotero compares it against the fields
 * this renderer reports. `ReferenceMark` and `Bookmark` are already
 * session-shaped and pass through. */
export function fieldTypeForSession(data: string): string {
  return remap(data, (value) => (value === DISK_FIELD_TYPE ? SESSION_FIELD_TYPE : value))
}
