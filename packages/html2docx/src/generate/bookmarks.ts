// @ts-nocheck — generation layer ported verbatim from untyped JS; it is typed
// file by file without logic changes, and until then strict consumers
// (apps/html, apps/shell) must not fail on it.
import { Bookmark } from 'docx'

function bookmarkName(id) {
  const value = String(id || '')
  let hash = 2166136261
  for (const char of value) {
    hash ^= char.codePointAt(0)
    hash = Math.imul(hash, 16777619)
  }
  const safe = value.replace(/[^A-Za-z0-9_]/g, '_').slice(0, 24) || 'anchor'
  return `h2d_${safe}_${(hash >>> 0).toString(36)}`.slice(0, 40)
}

function withBookmarks(children, bookmarks = []) {
  let wrapped = children
  for (const id of [...bookmarks].reverse()) {
    wrapped = [new Bookmark({ id: bookmarkName(id), children: wrapped })]
  }
  return wrapped
}

export { bookmarkName, withBookmarks }
