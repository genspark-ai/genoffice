import { useMemo, useState } from 'react'
import type { ReactElement } from 'react'
import type { NoteThreadItem } from './note-threads'
import type { TFunc } from './i18n/locale'

/** One thread root in the all-document list; pageNo is the display position (1-based) */
export interface CommentEntry {
  origIdx: number
  pageNo: number
  root: NoteThreadItem
}

export function countReplies(root: NoteThreadItem): number {
  let n = 0
  const walk = (items: NoteThreadItem[]) => {
    for (const it of items) {
      n++
      walk(it.replies)
    }
  }
  walk(root.replies)
  return n
}

/**
 * Filter + order the entries for the comments panel: case-insensitive substring
 * over the thread's own texts (root and replies) and authors, page order first
 * and reply-thread time as the tiebreaker (null times sort last within a page).
 */
export function filterCommentEntries(entries: CommentEntry[], query: string): CommentEntry[] {
  const q = query.trim().toLowerCase()
  const matches = q
    ? entries.filter((e) => {
        const texts: string[] = [e.root.contents, e.root.author]
        const walk = (items: NoteThreadItem[]) => {
          for (const it of items) {
            texts.push(it.contents, it.author)
            walk(it.replies)
          }
        }
        walk(e.root.replies)
        return texts.some((s) => s.toLowerCase().includes(q))
      })
    : entries
  return [...matches].sort((a, b) => {
    if (a.pageNo !== b.pageNo) return a.pageNo - b.pageNo
    const at = a.root.timeMs ?? Number.MAX_SAFE_INTEGER
    const bt = b.root.timeMs ?? Number.MAX_SAFE_INTEGER
    return at - bt
  })
}

function Row({
  entry,
  replies,
  onOpen,
}: {
  entry: CommentEntry
  replies: number
  onOpen: (e: CommentEntry) => void
}): ReactElement {
  const root = entry.root
  const title = [root.author, root.timeMs ? new Date(root.timeMs).toLocaleString() : '']
    .filter(Boolean)
    .join(' · ')
  return (
    <button
      type="button"
      className="pdf-comment-row"
      data-tip={root.contents}
      title={title || undefined}
      onClick={() => onOpen(entry)}
    >
      <span className="pdf-comment-page">P{entry.pageNo}</span>
      <span className="pdf-comment-body">
        <span className="pdf-comment-author">{root.author || '\u2014'}</span>
        <span className="pdf-comment-text">{root.contents}</span>
      </span>
      {replies > 0 && <span className="pdf-comment-replies">↩ {replies}</span>}
    </button>
  )
}

/**
 * All-document comments list (Acrobat's Comments panel): thread roots grouped by
 * page, a text filter over contents/authors, click jumps to the page and opens
 * the thread popover at its pin.
 */
export function CommentsPanel({
  t,
  entries,
  loading,
  onOpen,
}: {
  t: TFunc
  entries: CommentEntry[]
  loading: boolean
  onOpen: (e: CommentEntry) => void
}): ReactElement {
  const [query, setQuery] = useState('')
  const shown = useMemo(() => filterCommentEntries(entries, query), [entries, query])
  return (
    <div className="pdf-comments" role="list" aria-label={t('commentsPanel')}>
      <input
        className="pdf-comments-filter"
        type="search"
        placeholder={t('commentsFilter')}
        value={query}
        onChange={(e) => setQuery(e.target.value)}
      />
      {loading ? (
        <div className="pdf-outline-empty" role="status">
          {t('commentsLoading')}
        </div>
      ) : shown.length === 0 ? (
        <div className="pdf-outline-empty" role="status">
          {entries.length === 0 ? t('commentsEmpty') : t('searchNoResults')}
        </div>
      ) : (
        shown.map((e) => (
          <Row key={e.root.key} entry={e} replies={countReplies(e.root)} onOpen={onOpen} />
        ))
      )}
    </div>
  )
}
