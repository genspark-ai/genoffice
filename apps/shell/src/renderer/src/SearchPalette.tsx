import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import type { ProjectHomeApi, RecentEntry, TimelineEntryItem } from '../../shared/home-api'
import { FileBadge } from './FileBadge'
import { useI18n } from './locale'
import type { StringKey } from './locale'
import { useToast } from './toast'
import iconFolder from './assets/file-folder.svg'
import './palette.css'

// ── Search palette (rail search button, Ctrl+P, Ctrl+K) ─────
// One floating panel with two modes: Files (recent files, filtered as you
// type) and Conversations (AI chats recorded per file, newest first).

export type PaletteMode = 'files' | 'conversations'

/** rows shown before "Show N more…" */
const COLLAPSED_ROWS = 6
/** extra rows fetched per "Show N more…" click */
const PAGE = 50

interface ConversationItem {
  key: string
  title: string
  filePath: string
  projectName: string
  tsMs: number
}

/** 'now' · '5m' · '3h' · '2d' · '3mo' · '1y' — compact age like the reference */
export function formatAge(tsMs: number, nowMs: number = Date.now()): string {
  const s = Math.max(0, Math.floor((nowMs - tsMs) / 1000))
  if (s < 60) return 'now'
  const m = Math.floor(s / 60)
  if (m < 60) return `${m}m`
  const h = Math.floor(m / 60)
  if (h < 24) return `${h}h`
  const d = Math.floor(h / 24)
  if (d < 60) return `${d}d`
  const mo = Math.floor(d / 30)
  if (mo < 12) return `${mo}mo`
  return `${Math.floor(d / 365)}y`
}

function baseName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path
}

function dirName(path: string): string {
  const parts = path.split(/[\\/]/)
  parts.pop()
  return parts.slice(-2).join('/')
}

/** group per-message timeline rows into one item per chat; title = the first thing the user asked */
export function groupConversations(
  rows: readonly TimelineEntryItem[],
  projectName: string,
): ConversationItem[] {
  const byChat = new Map<string, TimelineEntryItem[]>()
  for (const row of rows) {
    const key = `${row.filePath}\u0000${row.chatId}`
    const list = byChat.get(key)
    if (list) list.push(row)
    else byChat.set(key, [row])
  }
  const items: ConversationItem[] = []
  for (const [key, list] of byChat) {
    const ordered = [...list].sort((a, b) => a.seq - b.seq)
    const first = ordered.find((r) => r.role === 'user' && r.preview.trim()) ?? ordered[0]!
    const last = ordered.reduce((acc, r) => (Date.parse(r.ts) > Date.parse(acc.ts) ? r : acc))
    const tsMs = Date.parse(last.ts)
    items.push({
      key,
      title: first.preview.trim().replace(/\s+/g, ' ') || first.fileName,
      filePath: first.filePath,
      projectName,
      tsMs: Number.isFinite(tsMs) ? tsMs : 0,
    })
  }
  return items
}

async function loadConversations(): Promise<ConversationItem[]> {
  const api = (window as unknown as { aiOfficeProject?: ProjectHomeApi }).aiOfficeProject
  if (!api) return []
  const projects = await api.listProjects()
  const lists = await Promise.all(
    projects.map(async (p: { id: string; name: string }) => {
      try {
        return groupConversations(await api.getTimeline(p.id, 300), p.name)
      } catch {
        return []
      }
    }),
  )
  return lists.flat()
}

interface PaletteProps {
  mode: PaletteMode
  onModeChange: (mode: PaletteMode) => void
  onClose: () => void
}

export function SearchPalette({ mode, onModeChange, onClose }: PaletteProps) {
  const { t } = useI18n()
  const toast = useToast()
  const inputRef = useRef<HTMLInputElement>(null)
  const listRef = useRef<HTMLDivElement>(null)
  const [query, setQuery] = useState('')
  const [debounced, setDebounced] = useState('')
  const [menuOpen, setMenuOpen] = useState(false)
  const [active, setActive] = useState(0)
  const [expanded, setExpanded] = useState(false)
  const [sortOldest, setSortOldest] = useState(false)

  const [files, setFiles] = useState<RecentEntry[]>([])
  const [fileTotal, setFileTotal] = useState(0)
  const [filesLoaded, setFilesLoaded] = useState(false)
  const [conversations, setConversations] = useState<ConversationItem[] | null>(null)
  const seq = useRef(0)

  // typing debounce → one request per pause, stale replies dropped via seq
  useEffect(() => {
    const id = window.setTimeout(() => setDebounced(query.trim()), 140)
    return () => window.clearTimeout(id)
  }, [query])

  // switching modes keeps the box but resets list state
  useEffect(() => {
    setActive(0)
    setExpanded(false)
    setMenuOpen(false)
    inputRef.current?.focus()
  }, [mode])

  // files: fetch the first page (or the expanded amount) for the current text
  const limit = expanded ? COLLAPSED_ROWS + PAGE : COLLAPSED_ROWS
  useEffect(() => {
    if (mode !== 'files') return
    const mine = ++seq.current
    void window.aiOffice
      .recents({ query: debounced || undefined, limit, offset: 0 })
      .then((page) => {
        if (mine !== seq.current) return
        setFiles(page.entries)
        setFileTotal(page.total)
        setFilesLoaded(true)
        setActive((a) => Math.min(a, Math.max(0, page.entries.length - 1)))
      })
      .catch(() => {
        if (mine !== seq.current) return
        setFiles([])
        setFileTotal(0)
        setFilesLoaded(true)
      })
  }, [mode, debounced, limit])

  // conversations: load once on first entry, filter locally
  useEffect(() => {
    if (mode !== 'conversations' || conversations !== null) return
    let alive = true
    void loadConversations()
      .then((items) => alive && setConversations(items))
      .catch(() => alive && setConversations([]))
    return () => {
      alive = false
    }
  }, [mode, conversations])

  const visibleConversations = useMemo(() => {
    if (!conversations) return []
    const q = debounced.toLowerCase()
    const filtered = q
      ? conversations.filter(
          (c) => c.title.toLowerCase().includes(q) || c.projectName.toLowerCase().includes(q),
        )
      : conversations
    return [...filtered].sort((a, b) => (sortOldest ? a.tsMs - b.tsMs : b.tsMs - a.tsMs))
  }, [conversations, debounced, sortOldest])

  const shownConversations = expanded
    ? visibleConversations
    : visibleConversations.slice(0, COLLAPSED_ROWS + 6)
  const hiddenConversations = visibleConversations.length - shownConversations.length
  const hiddenFiles = Math.max(0, fileTotal - files.length)

  const rowCount = mode === 'files' ? files.length : shownConversations.length

  const openFile = useCallback(
    async (path: string) => {
      onClose()
      try {
        await window.aiOffice.openPath(path)
      } catch {
        toast.show({ message: t('palOpenError') })
      }
    },
    [onClose, t, toast],
  )

  const activate = (index: number): void => {
    if (mode === 'files') {
      const entry = files[index]
      if (entry) void openFile(entry.path)
    } else {
      const item = shownConversations[index]
      if (item) void openFile(item.filePath)
    }
  }

  // keep the highlighted row scrolled into view
  useEffect(() => {
    listRef.current
      ?.querySelector<HTMLElement>(`[data-index="${active}"]`)
      ?.scrollIntoView({ block: 'nearest' })
  }, [active, rowCount])

  const onKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === 'Escape') {
      e.preventDefault()
      e.stopPropagation()
      if (menuOpen) setMenuOpen(false)
      else onClose()
    } else if (e.key === 'ArrowDown') {
      e.preventDefault()
      if (rowCount > 0) setActive((a) => (a + 1) % rowCount)
    } else if (e.key === 'ArrowUp') {
      e.preventDefault()
      if (rowCount > 0) setActive((a) => (a - 1 + rowCount) % rowCount)
    } else if (e.key === 'Home' && !query) {
      e.preventDefault()
      setActive(0)
    } else if (e.key === 'End' && !query) {
      e.preventDefault()
      setActive(Math.max(0, rowCount - 1))
    } else if (e.key === 'Enter') {
      e.preventDefault()
      activate(active)
    }
  }

  const modeLabel = mode === 'files' ? t('palFiles') : t('palConversations')
  const placeholder = mode === 'files' ? t('palSearchFiles') : t('palSearchConversations')
  const listId = 'pal-list'

  const node = (
    <div
      className="pal-overlay"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose()
      }}
    >
      <div
        className="pal-panel"
        role="dialog"
        aria-modal="true"
        aria-label={t('palDialog')}
        onKeyDown={onKeyDown}
      >
        <div className="pal-head">
          <div className="pal-mode-wrap">
            <button
              type="button"
              className="pal-mode"
              aria-haspopup="listbox"
              aria-expanded={menuOpen}
              aria-label={t('palMode')}
              onClick={() => setMenuOpen((v) => !v)}
            >
              {modeLabel}
              <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden="true">
                <path
                  d="M2 3.6 5 6.6l3-3"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="1.4"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            </button>
            {menuOpen && (
              <div className="pal-mode-menu" role="listbox" aria-label={t('palMode')}>
                {(
                  [
                    ['files', 'palFiles'],
                    ['conversations', 'palConversations'],
                  ] as const satisfies readonly (readonly [PaletteMode, StringKey])[]
                ).map(([value, key]) => (
                  <button
                    key={value}
                    type="button"
                    role="option"
                    aria-selected={mode === value}
                    className={`pal-mode-item${mode === value ? ' on' : ''}`}
                    onClick={() => {
                      onModeChange(value)
                      setMenuOpen(false)
                    }}
                  >
                    {t(key)}
                    {mode === value && (
                      <svg width="12" height="12" viewBox="0 0 12 12" aria-hidden="true">
                        <path
                          d="m2.5 6.4 2.3 2.3 4.7-5"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="1.5"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        />
                      </svg>
                    )}
                  </button>
                ))}
              </div>
            )}
          </div>
          <input
            ref={inputRef}
            className="pal-input"
            value={query}
            placeholder={placeholder}
            spellCheck={false}
            autoFocus
            role="combobox"
            aria-expanded="true"
            aria-controls={listId}
            aria-activedescendant={rowCount > 0 ? `pal-row-${active}` : undefined}
            aria-label={placeholder}
            onChange={(e) => {
              setQuery(e.target.value)
              setActive(0)
            }}
          />
          {mode === 'conversations' && (
            <button
              type="button"
              className="pal-sort"
              aria-pressed={sortOldest}
              title={sortOldest ? t('palSortOldest') : t('palSortNewest')}
              aria-label={sortOldest ? t('palSortOldest') : t('palSortNewest')}
              onClick={() => setSortOldest((v) => !v)}
            >
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                <path
                  d="M2.5 4h11M4.5 8h7M6.5 12h3"
                  stroke="currentColor"
                  strokeWidth="1.3"
                  strokeLinecap="round"
                />
              </svg>
            </button>
          )}
        </div>

        <div className="pal-list" id={listId} role="listbox" ref={listRef} aria-label={modeLabel}>
          {mode === 'files' && (
            <>
              {files.length > 0 && (
                <div className="pal-section">{debounced ? t('palResults') : t('palRecentFiles')}</div>
              )}
              {files.map((f, i) => (
                <div
                  key={f.path}
                  id={`pal-row-${i}`}
                  data-index={i}
                  role="option"
                  aria-selected={i === active}
                  className={`pal-row${i === active ? ' active' : ''}`}
                  onMouseMove={() => i !== active && setActive(i)}
                  onClick={() => activate(i)}
                >
                  <FileBadge ext={f.ext} size={16} />
                  <span className="pal-title">{f.name}</span>
                  {debounced && <span className="pal-meta">{dirName(f.path)}</span>}
                </div>
              ))}
              {hiddenFiles > 0 && (
                <button type="button" className="pal-more" onClick={() => setExpanded(true)}>
                  {t('palShowMore', { n: hiddenFiles })}
                </button>
              )}
              {filesLoaded && files.length === 0 && (
                <div className="pal-empty">{t('palNoFiles')}</div>
              )}
            </>
          )}
          {mode === 'conversations' && (
            <>
              {shownConversations.map((c, i) => (
                <div
                  key={c.key}
                  id={`pal-row-${i}`}
                  data-index={i}
                  role="option"
                  aria-selected={i === active}
                  className={`pal-row${i === active ? ' active' : ''}`}
                  onMouseMove={() => i !== active && setActive(i)}
                  onClick={() => activate(i)}
                  title={baseName(c.filePath)}
                >
                  <span className="pal-title">{c.title}</span>
                  <span className="pal-project">
                    <img src={iconFolder} width={13} height={13} alt="" aria-hidden="true" />
                    {c.projectName}
                  </span>
                  <span className="pal-age">{formatAge(c.tsMs)}</span>
                </div>
              ))}
              {hiddenConversations > 0 && (
                <button type="button" className="pal-more" onClick={() => setExpanded(true)}>
                  {t('palShowMore', { n: hiddenConversations })}
                </button>
              )}
              {conversations === null && <div className="pal-empty">{t('palLoading')}</div>}
              {conversations !== null && visibleConversations.length === 0 && (
                <div className="pal-empty">{t('palNoConversations')}</div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  )

  return typeof document !== 'undefined' ? createPortal(node, document.body) : node
}
