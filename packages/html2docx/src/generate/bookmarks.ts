// @ts-nocheck — generation layer ported verbatim from untyped JS; it is typed
// file by file without logic changes, and until then strict consumers
// (apps/html, apps/shell) must not fail on it.
import { BookmarkEnd, BookmarkStart, bookmarkUniqueNumericIdGen } from 'docx'

// docx's document-wide counter, so ids never collide with its own Bookmarks.
const nextBookmarkLinkId = bookmarkUniqueNumericIdGen()

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
  if (!bookmarks?.length) return children
  // docx only unwraps a top-level Bookmark; nesting them drops the runs.
  const starts = []
  const ends = []
  for (const id of bookmarks) {
    const linkId = nextBookmarkLinkId()
    starts.push(new BookmarkStart(bookmarkName(id), linkId))
    ends.unshift(new BookmarkEnd(linkId))
  }
  return [...starts, ...children, ...ends]
}

function mergeBookmarks(existing, incoming) {
  return [...new Set([...(existing || []), ...incoming])]
}

function carryInto(entry, bookmarks) {
  if (!entry || typeof entry !== 'object') return null
  if (Array.isArray(entry.children)) {
    const children = carryBookmarks(entry.children, bookmarks)
    return children ? { ...entry, children } : null
  }
  if (Array.isArray(entry.runs)) {
    return { ...entry, bookmarks: mergeBookmarks(entry.bookmarks, bookmarks) }
  }
  return null
}

// w:bookmarkStart is paragraph-scoped: move a container's bookmark onto its
// first paragraph/heading/list item. null means nothing was anchorable.
function carryBookmarks(nodes, bookmarks) {
  if (!Array.isArray(nodes)) return null
  for (let index = 0; index < nodes.length; index++) {
    const node = nodes[index]
    if (!node || typeof node !== 'object') continue
    let carried = null
    if (node.type === 'para' || node.type === 'heading') {
      carried = { ...node, bookmarks: mergeBookmarks(node.bookmarks, bookmarks) }
    } else if (node.type === 'list' && node.items?.length) {
      const items = [...node.items]
      items[0] = { ...items[0], bookmarks: mergeBookmarks(items[0].bookmarks, bookmarks) }
      carried = { ...node, items }
    } else if (node.rows?.[0]?.cells?.[0]) {
      const { bookmarks: _, ...table } = carryTableBookmarks({
        ...node,
        bookmarks: mergeBookmarks(node.bookmarks, bookmarks),
      })
      carried = table
    } else {
      for (const key of ['children', 'cells']) {
        const kids = node[key]
        if (!Array.isArray(kids)) continue
        for (let k = 0; k < kids.length; k++) {
          const replacement = carryInto(kids[k], bookmarks)
          if (!replacement) continue
          const nextKids = [...kids]
          nextKids[k] = replacement
          carried = { ...node, [key]: nextKids }
          break
        }
        if (carried) break
      }
    }
    if (!carried) continue
    const out = [...nodes]
    out[index] = carried
    return out
  }
  return null
}

function carryTableBookmarks(node) {
  const first = node.rows?.[0]?.cells?.[0]
  if (!node.bookmarks?.length || !first) return node
  const [head, ...rest] = node.rows
  const [lead, ...tail] = head.cells
  return {
    ...node,
    rows: [
      {
        ...head,
        cells: [{ ...lead, bookmarks: mergeBookmarks(lead.bookmarks, node.bookmarks) }, ...tail],
      },
      ...rest,
    ],
  }
}

export { bookmarkName, carryBookmarks, carryTableBookmarks, withBookmarks }
