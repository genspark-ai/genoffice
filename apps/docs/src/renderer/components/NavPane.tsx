import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type DragEvent,
  type PointerEvent as ReactPointerEvent,
} from 'react'
import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { TextSelection, type Transaction } from '@tiptap/pm/state'
import { collectHeadings, type HeadingRef, type HeadingStyles } from '../editor/headings'
import { searchPluginKey } from '../editor/extensions'
import {
  HIGHLIGHT_CAP,
  RESULTS_LIST_CAP,
  SCAN_THRESHOLD,
  compileFind,
  findInDoc,
  findInDocAsync,
  type FindMatch,
  type FindOptions,
} from '../editor/find'
import {
  MAX_HEADING_LEVEL,
  buildHeadingTree,
  deleteHeadingSubtree,
  headingAtPos,
  headingKey,
  hitCounts,
  insertHeadingSibling,
  moveHeadingSubtree,
  resultContext,
  selectHeadingSubtree,
  setHeadingLevel,
  visibleRows,
  type NavNode,
} from '../editor/nav-outline'
import { useI18n } from '../i18n/locale'
import { NavPageThumbs } from './NavPageThumbs'
import { IconClose, IconSearch } from './icons'

export type NavTab = 'headings' | 'pages' | 'results'

/** per-document-session pane state (Word keeps collapse state until the document closes) */
interface NavSession {
  tab: NavTab
  collapsed: Set<string>
  maxLevel: number
  query: string
}

const sessions = new WeakMap<Editor, NavSession>()
const sessionOf = (editor: Editor): NavSession => {
  let s = sessions.get(editor)
  if (!s) {
    s = { tab: 'headings', collapsed: new Set(), maxLevel: MAX_HEADING_LEVEL, query: '' }
    sessions.set(editor, s)
  }
  return s
}

const WIDTH_KEY = 'aidocs.navWidth'
const MIN_WIDTH = 180
const MAX_WIDTH = 560
const DEFAULT_WIDTH = 232
const SCAN_DEBOUNCE_MS = 150
const SEARCH_OPTIONS: FindOptions = { matchCase: false, wholeWord: false, mode: 'literal' }

const readWidth = (): number => {
  const n = Number(localStorage.getItem(WIDTH_KEY))
  return n >= MIN_WIDTH && n <= MAX_WIDTH ? n : DEFAULT_WIDTH
}

interface MenuState {
  x: number
  y: number
  index: number
}

export function NavPane({
  editor,
  doc,
  zoom,
  pageInfo,
  onClose,
}: {
  editor: Editor
  doc: PmNode
  /** canvas zoom factor (1 = 100%) */
  zoom: number
  pageInfo: { current: number; total: number }
  onClose: () => void
}) {
  const { t } = useI18n()
  const session = sessionOf(editor)
  const canEdit = editor.isEditable
  const styles = editor.storage.listNumbering?.styles as HeadingStyles | undefined
  const headings = useMemo(() => collectHeadings(doc, styles, true), [doc, styles])
  const tree = useMemo(() => buildHeadingTree(headings), [headings])
  const [tab, setTabState] = useState<NavTab>(session.tab)
  const [collapsedVersion, setCollapsedVersion] = useState(0)
  const [maxLevel, setMaxLevelState] = useState(session.maxLevel)
  const [query, setQuery] = useState(session.query)
  const [matches, setMatches] = useState<FindMatch[]>([])
  const [index, setIndex] = useState(0)
  const [scanning, setScanning] = useState(false)
  const [width, setWidth] = useState(readWidth)
  const [caret, setCaret] = useState(() => editor.state.selection.from)
  const [menu, setMenu] = useState<MenuState | null>(null)
  const [drag, setDrag] = useState<{ index: number; where: 'before' | 'after' } | null>(null)
  const dragFromRef = useRef<number | null>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  const treeRef = useRef<HTMLDivElement>(null)
  const indexRef = useRef(0)
  const matchesRef = useRef<FindMatch[]>([])
  const scanTokenRef = useRef<{ aborted: boolean } | null>(null)
  const timerRef = useRef<number | null>(null)
  const pendingKeepRef = useRef<'reset' | 'current'>('reset')
  // the search decorations are shared with the Find panel: only touch them while this pane has a query
  const ownsHighlightRef = useRef(false)
  const preSearchTabRef = useRef<NavTab>(session.tab === 'results' ? 'headings' : session.tab)

  const setTab = useCallback(
    (next: NavTab) => {
      session.tab = next
      setTabState(next)
    },
    [session],
  )
  const setMaxLevel = (n: number) => {
    session.maxLevel = n
    setMaxLevelState(n)
  }
  const toggleCollapsed = (ref: HeadingRef) => {
    const key = headingKey(ref)
    if (session.collapsed.has(key)) session.collapsed.delete(key)
    else session.collapsed.add(key)
    setCollapsedVersion((v) => v + 1)
  }
  const setAllCollapsed = (collapsed: boolean) => {
    session.collapsed.clear()
    if (collapsed) for (const h of headings) session.collapsed.add(headingKey(h))
    setCollapsedVersion((v) => v + 1)
  }

  const rows = useMemo(
    () => visibleRows(tree, session.collapsed, maxLevel),
    // collapsedVersion invalidates the memo when the (mutable) set changes
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [tree, maxLevel, collapsedVersion],
  )

  // ---- caret tracking: the heading whose section holds the caret is highlighted ----
  useEffect(() => {
    let raf = 0
    const onSel = () => {
      cancelAnimationFrame(raf)
      raf = requestAnimationFrame(() => setCaret(editor.state.selection.from))
    }
    editor.on('selectionUpdate', onSel)
    editor.on('update', onSel)
    return () => {
      cancelAnimationFrame(raf)
      editor.off('selectionUpdate', onSel)
      editor.off('update', onSel)
    }
  }, [editor])

  const currentIndex = useMemo(() => {
    let idx = headingAtPos(headings, caret)
    if (idx < 0) return -1
    const shown = new Set(rows.map((r) => r.node.index))
    const parentOf = new Map<number, number>()
    const walk = (nodes: NavNode[]) =>
      nodes.forEach((n) => {
        parentOf.set(n.index, n.parent)
        walk(n.children)
      })
    walk(tree)
    // a caret inside a collapsed or level-hidden branch lights its nearest visible ancestor
    while (idx >= 0 && !shown.has(idx)) idx = parentOf.get(idx) ?? -1
    return idx
  }, [headings, caret, rows, tree])

  useEffect(() => {
    if (currentIndex < 0) return
    treeRef.current
      ?.querySelector<HTMLElement>(`[data-index="${currentIndex}"]`)
      ?.scrollIntoView({ block: 'nearest' })
  }, [currentIndex])

  // ---- search (same engine as the Find panel; literal, case-insensitive) ----
  const highlight = useCallback(
    (ranges: FindMatch[], activeIndex: number) => {
      if (ranges.length === 0 && !ownsHighlightRef.current) return
      ownsHighlightRef.current = ranges.length > 0
      const shown =
        ranges.length > HIGHLIGHT_CAP && ranges[activeIndex]
          ? { ranges: [ranges[activeIndex]], activeIndex: 0 }
          : { ranges, activeIndex }
      editor.view.dispatch(editor.state.tr.setMeta(searchPluginKey, shown))
    },
    [editor],
  )

  const commit = useCallback(
    (ranges: FindMatch[], keepIndex: number) => {
      const active = ranges.length === 0 ? 0 : Math.min(keepIndex, ranges.length - 1)
      setMatches(ranges)
      matchesRef.current = ranges
      setIndex(active)
      indexRef.current = active
      highlight(ranges, active)
      return ranges
    },
    [highlight],
  )

  const refresh = useCallback(
    (q: string, keepIndex = 0): FindMatch[] | null => {
      if (scanTokenRef.current) scanTokenRef.current.aborted = true
      scanTokenRef.current = null
      setScanning(false)
      matchesRef.current = []
      const compiled = q ? compileFind(q, SEARCH_OPTIONS) : null
      if (!compiled || !compiled.ok) return commit([], 0)
      const current = editor.state.doc
      if (current.nodeSize <= SCAN_THRESHOLD.nodeSize)
        return commit(findInDoc(current, compiled.query, SEARCH_OPTIONS), keepIndex)
      const token = { aborted: false }
      scanTokenRef.current = token
      setScanning(true)
      setMatches([])
      void findInDocAsync(current, compiled.query, SEARCH_OPTIONS, token).then((found) => {
        if (!found || scanTokenRef.current !== token) return
        scanTokenRef.current = null
        setScanning(false)
        commit(found, keepIndex)
      })
      return null
    },
    [editor, commit],
  )
  const refreshRef = useRef(refresh)
  refreshRef.current = refresh

  const scheduleRefresh = useCallback((q: string, keep: 'reset' | 'current') => {
    if (timerRef.current !== null) window.clearTimeout(timerRef.current)
    pendingKeepRef.current = keep
    timerRef.current = window.setTimeout(() => {
      timerRef.current = null
      refreshRef.current(q, keep === 'current' ? indexRef.current : 0)
    }, SCAN_DEBOUNCE_MS)
  }, [])

  const flushPending = useCallback(
    (q: string) => {
      if (timerRef.current === null) return null
      window.clearTimeout(timerRef.current)
      timerRef.current = null
      const keep = pendingKeepRef.current
      const ranges = refresh(q, keep === 'current' ? indexRef.current : 0)
      return ranges ? { ranges, queryChanged: keep === 'reset' } : null
    },
    [refresh],
  )

  const queryRef = useRef(query)
  queryRef.current = query
  useEffect(() => {
    const onUpdate = () => {
      if (queryRef.current) scheduleRefresh(queryRef.current, 'current')
    }
    editor.on('update', onUpdate)
    return () => {
      editor.off('update', onUpdate)
    }
  }, [editor, scheduleRefresh])

  // a query restored from the session runs on mount; highlights leave with the pane
  useEffect(() => {
    if (session.query) scheduleRefresh(session.query, 'reset')
    return () => {
      if (timerRef.current !== null) window.clearTimeout(timerRef.current)
      if (scanTokenRef.current) scanTokenRef.current.aborted = true
      if (!editor.isDestroyed && ownsHighlightRef.current)
        editor.view.dispatch(
          editor.state.tr.setMeta(searchPluginKey, { ranges: [], activeIndex: 0 }),
        )
    }
    // mount only
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const onQueryChange = (q: string) => {
    if (!query && q) {
      preSearchTabRef.current = tab === 'results' ? 'headings' : tab
      setTab('results')
    } else if (query && !q) setTab(preSearchTabRef.current)
    session.query = q
    setQuery(q)
    scheduleRefresh(q, 'reset')
  }

  const goToMatch = useCallback(
    (ranges: FindMatch[], next: number) => {
      const m = ranges[next]
      if (!m) return
      setIndex(next)
      indexRef.current = next
      highlight(ranges, next)
      const { state, view } = editor
      view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, m.from, m.to)))
      const { node } = view.domAtPos(m.from)
      const el = node instanceof HTMLElement ? node : node.parentElement
      el?.scrollIntoView({ block: 'center' })
    },
    [editor, highlight],
  )

  const step = (dir: 1 | -1) => {
    const fresh = flushPending(query)
    if (scanTokenRef.current) return
    const ranges = fresh ? fresh.ranges : matchesRef.current
    if (ranges.length === 0) return
    // a flushed scan for a new query already sits on its first hit; a rescan after a
    // document edit keeps the position and Enter still moves from there
    const next = fresh?.queryChanged
      ? indexRef.current
      : (indexRef.current + dir + ranges.length) % ranges.length
    goToMatch(ranges, next)
  }

  const counts = useMemo(
    () => (query ? hitCounts(headings, matches, doc.content.size) : null),
    [query, headings, matches, doc],
  )
  const hitPositions = useMemo(
    () => (query ? matches.slice(0, HIGHLIGHT_CAP).map((m) => m.from) : null),
    [query, matches],
  )

  // ---- heading actions ----
  const jumpToHeading = (pos: number) => {
    const { state, view } = editor
    const sel = TextSelection.near(state.doc.resolve(pos + 1), 1)
    view.dispatch(state.tr.setSelection(sel))
    const dom = view.nodeDOM(pos)
    if (dom instanceof HTMLElement) dom.scrollIntoView({ block: 'start' })
    view.focus()
  }

  const run = (fn: (tr: Transaction) => boolean) => {
    const { state, view } = editor
    const tr = state.tr
    if (fn(tr)) {
      view.dispatch(tr.scrollIntoView())
      view.focus()
    }
  }

  const changeLevel = (i: number, delta: number) =>
    run((tr) =>
      setHeadingLevel(tr, editor.schema, headings[i].pos, headings[i].level + delta, styles),
    )

  const onDragStart = (e: DragEvent, i: number) => {
    dragFromRef.current = i
    e.dataTransfer.effectAllowed = 'move'
    e.dataTransfer.setData('text/plain', headings[i].text)
  }
  const onDragOver = (e: DragEvent, i: number) => {
    const from = dragFromRef.current
    if (from === null || !canEdit) return
    const size = doc.content.size
    const srcEnd =
      headings.slice(from + 1).find((h) => h.level <= headings[from].level)?.pos ?? size
    // no drop inside the dragged subtree
    if (headings[i].pos >= headings[from].pos && headings[i].pos < srcEnd) {
      setDrag(null)
      return
    }
    e.preventDefault()
    e.dataTransfer.dropEffect = 'move'
    const rect = e.currentTarget.getBoundingClientRect()
    const where = e.clientY - rect.top < rect.height / 2 ? 'before' : 'after'
    setDrag((cur) => (cur && cur.index === i && cur.where === where ? cur : { index: i, where }))
  }
  const onDrop = (e: DragEvent, i: number) => {
    e.preventDefault()
    const from = dragFromRef.current
    const where = drag?.index === i ? drag.where : 'before'
    dragFromRef.current = null
    setDrag(null)
    if (from === null) return
    run((tr) => moveHeadingSubtree(tr, headings, from, i, where))
  }
  const endDrag = () => {
    dragFromRef.current = null
    setDrag(null)
  }

  // ---- resize ----
  const onResizeStart = (e: ReactPointerEvent<HTMLDivElement>) => {
    const startX = e.clientX
    const startW = width
    const el = e.currentTarget
    el.setPointerCapture(e.pointerId)
    let next = startW
    const move = (ev: PointerEvent) => {
      next = Math.min(MAX_WIDTH, Math.max(MIN_WIDTH, startW + ev.clientX - startX))
      setWidth(next)
    }
    const up = () => {
      el.removeEventListener('pointermove', move)
      el.removeEventListener('pointerup', up)
      localStorage.setItem(WIDTH_KEY, String(next))
    }
    el.addEventListener('pointermove', move)
    el.addEventListener('pointerup', up)
  }

  const [layoutVersion, setLayoutVersion] = useState(0)
  useEffect(() => setLayoutVersion((v) => v + 1), [doc, pageInfo.total])

  const tabs: Array<{ id: NavTab; label: string }> = [
    { id: 'headings', label: t('appNavTabHeadings') },
    { id: 'pages', label: t('appNavTabPages') },
    { id: 'results', label: t('appNavTabResults') },
  ]

  return (
    <aside className="nav-pane" style={{ width }}>
      <div className="nav-pane-head">
        <span className="nav-pane-title">{t('appNavTitle')}</span>
        <button
          className="nav-pane-close"
          aria-label={t('appClose')}
          data-tip={t('appClose')}
          onClick={onClose}
        >
          <IconClose size={14} />
        </button>
      </div>
      <div className="nav-search">
        <IconSearch size={13} />
        <input
          ref={inputRef}
          className="nav-search-input"
          type="search"
          value={query}
          placeholder={t('appNavSearch')}
          spellCheck={false}
          onChange={(e) => onQueryChange(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault()
              step(e.shiftKey ? -1 : 1)
            } else if (e.key === 'Escape' && query) {
              e.preventDefault()
              onQueryChange('')
            }
          }}
        />
        {query && (
          <span className="nav-search-count">
            {scanning ? '…' : matches.length === 0 ? '0' : `${index + 1}/${matches.length}`}
          </span>
        )}
      </div>
      <div className="nav-tabs" role="tablist">
        {tabs.map((tb) => (
          <button
            key={tb.id}
            role="tab"
            type="button"
            aria-selected={tab === tb.id}
            className={`nav-tab${tab === tb.id ? ' on' : ''}`}
            onClick={() => setTab(tb.id)}
          >
            {tb.label}
          </button>
        ))}
      </div>
      <div className="nav-body">
        {tab === 'headings' && (
          <div
            className="nav-tree"
            role="tree"
            ref={treeRef}
            onDragLeave={(e) => {
              if (!e.currentTarget.contains(e.relatedTarget as Node | null)) setDrag(null)
            }}
          >
            {rows.map((row) => {
              const i = row.node.index
              const count = counts ? counts.subtree[i] : 0
              const cls = [
                'nav-row',
                i === currentIndex ? 'on' : '',
                counts && count === 0 ? 'dim' : '',
                drag?.index === i ? `drop-${drag.where}` : '',
              ]
                .filter(Boolean)
                .join(' ')
              return (
                <div
                  key={`${headingKey(row.node.ref)}-${i}`}
                  className={cls}
                  role="treeitem"
                  aria-level={row.node.ref.level}
                  aria-expanded={row.hasChildren ? !row.collapsed : undefined}
                  aria-selected={i === currentIndex}
                  data-index={i}
                  data-tip={row.node.ref.text}
                  tabIndex={0}
                  draggable={canEdit}
                  style={{ paddingLeft: 6 + row.depth * 14 }}
                  onClick={() => jumpToHeading(row.node.ref.pos)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') jumpToHeading(row.node.ref.pos)
                  }}
                  onContextMenu={(e) => {
                    e.preventDefault()
                    setMenu({ x: e.clientX, y: e.clientY, index: i })
                  }}
                  onDragStart={(e) => onDragStart(e, i)}
                  onDragOver={(e) => onDragOver(e, i)}
                  onDrop={(e) => onDrop(e, i)}
                  onDragEnd={endDrag}
                >
                  <button
                    type="button"
                    className={`nav-twisty${row.hasChildren ? '' : ' leaf'}${row.collapsed ? ' closed' : ''}`}
                    tabIndex={-1}
                    aria-label={row.collapsed ? t('appNavExpand') : t('appNavCollapse')}
                    onClick={(e) => {
                      e.stopPropagation()
                      if (row.hasChildren) toggleCollapsed(row.node.ref)
                    }}
                  >
                    {row.hasChildren && (
                      <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden>
                        <path d="M1.5 1l4 3-4 3z" fill="currentColor" />
                      </svg>
                    )}
                  </button>
                  <span className="nav-row-text">{row.node.ref.text || '\u00a0'}</span>
                  {counts && count > 0 && (
                    <span className="nav-badge" title={t('appFindResults', { count })}>
                      {count}
                    </span>
                  )}
                </div>
              )
            })}
            {rows.length === 0 && <div className="nav-empty">{t('appNavNoHeadings')}</div>}
          </div>
        )}
        {tab === 'pages' && (
          <NavPageThumbs
            editor={editor}
            zoom={zoom}
            width={width}
            layoutVersion={layoutVersion}
            currentPage={pageInfo.current}
            hitPositions={hitPositions}
          />
        )}
        {tab === 'results' && (
          <div className="nav-results" role="listbox">
            {!query && <div className="nav-empty">{t('appNavSearchHint')}</div>}
            {query && !scanning && matches.length === 0 && (
              <div className="nav-empty">{t('appNavNoResults')}</div>
            )}
            {matches.slice(0, RESULTS_LIST_CAP).map((m, i) => {
              const ctx = resultContext(m)
              return (
                <button
                  type="button"
                  key={`${m.from}-${i}`}
                  role="option"
                  aria-selected={i === index}
                  className={`nav-result${i === index ? ' on' : ''}`}
                  onClick={() => goToMatch(matches, i)}
                >
                  <span className="nav-result-ctx">{ctx.before}</span>
                  <mark>{ctx.hit}</mark>
                  <span className="nav-result-ctx">{ctx.after}</span>
                </button>
              )
            })}
            {matches.length > RESULTS_LIST_CAP && (
              <div className="nav-results-note">
                {t('appFindResultsCapped', { shown: RESULTS_LIST_CAP, total: matches.length })}
              </div>
            )}
          </div>
        )}
      </div>
      <div className="nav-pane-resizer" onPointerDown={onResizeStart} />
      {menu && (
        <NavMenu
          menu={menu}
          heading={headings[menu.index]}
          canEdit={canEdit}
          maxLevel={maxLevel}
          onClose={() => setMenu(null)}
          onPromote={() => changeLevel(menu.index, -1)}
          onDemote={() => changeLevel(menu.index, 1)}
          onNewBefore={() =>
            run((tr) => insertHeadingSibling(tr, editor.schema, headings, menu.index, 'before'))
          }
          onNewAfter={() =>
            run((tr) => insertHeadingSibling(tr, editor.schema, headings, menu.index, 'after'))
          }
          onDelete={() =>
            run((tr) => deleteHeadingSubtree(tr, editor.schema, headings, menu.index))
          }
          onSelect={() =>
            run((tr) => {
              selectHeadingSubtree(tr, headings, menu.index)
              return true
            })
          }
          onExpandAll={() => setAllCollapsed(false)}
          onCollapseAll={() => setAllCollapsed(true)}
          onShowLevels={setMaxLevel}
        />
      )}
    </aside>
  )
}

function NavMenu({
  menu,
  heading,
  canEdit,
  maxLevel,
  onClose,
  onPromote,
  onDemote,
  onNewBefore,
  onNewAfter,
  onDelete,
  onSelect,
  onExpandAll,
  onCollapseAll,
  onShowLevels,
}: {
  menu: MenuState
  heading: HeadingRef | undefined
  canEdit: boolean
  maxLevel: number
  onClose: () => void
  onPromote: () => void
  onDemote: () => void
  onNewBefore: () => void
  onNewAfter: () => void
  onDelete: () => void
  onSelect: () => void
  onExpandAll: () => void
  onCollapseAll: () => void
  onShowLevels: (n: number) => void
}) {
  const { t } = useI18n()
  const ref = useRef<HTMLDivElement>(null)
  const [pos, setPos] = useState({ left: menu.x, top: menu.y })
  const [levels, setLevels] = useState(false)

  useEffect(() => {
    const el = ref.current
    if (!el) return
    const rect = el.getBoundingClientRect()
    let left = menu.x
    let top = menu.y
    if (left + rect.width > window.innerWidth - 8) left = window.innerWidth - rect.width - 8
    if (top + rect.height > window.innerHeight - 8) top = Math.max(8, menu.y - rect.height)
    setPos({ left, top })
  }, [menu])

  useEffect(() => {
    const close = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose()
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('mousedown', close)
    window.addEventListener('keydown', onKey)
    window.addEventListener('blur', onClose)
    return () => {
      window.removeEventListener('mousedown', close)
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('blur', onClose)
    }
  }, [onClose])

  if (!heading) return null
  const item = (label: string, onClick: () => void, disabled = false) => (
    <button
      type="button"
      className="ctx-item"
      disabled={disabled}
      onMouseEnter={() => setLevels(false)}
      onClick={() => {
        onClose()
        onClick()
      }}
    >
      <span className="ctx-label">{label}</span>
    </button>
  )
  return (
    <div
      ref={ref}
      className="ctx-menu nav-menu"
      style={{ left: pos.left, top: pos.top }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {canEdit && (
        <>
          {item(t('appNavPromote'), onPromote, heading.level <= 1)}
          {item(t('appNavDemote'), onDemote, heading.level >= MAX_HEADING_LEVEL)}
          <div className="ctx-sep" />
          {item(t('appNavNewBefore'), onNewBefore)}
          {item(t('appNavNewAfter'), onNewAfter)}
          {item(t('appNavDelete'), onDelete)}
          <div className="ctx-sep" />
        </>
      )}
      {item(t('appNavSelectContent'), onSelect)}
      <div className="ctx-sep" />
      {item(t('appNavExpandAll'), onExpandAll)}
      {item(t('appNavCollapseAll'), onCollapseAll)}
      <div className="ctx-item-wrap" onMouseLeave={() => setLevels(false)}>
        <button type="button" className="ctx-item" onMouseEnter={() => setLevels(true)}>
          <span className="ctx-label">{t('appNavShowLevels')}</span>
          <span className="ctx-arrow">›</span>
        </button>
        {levels && (
          <div className="ctx-submenu">
            {Array.from({ length: MAX_HEADING_LEVEL }, (_, k) => k + 1).map((n) => (
              <button
                key={n}
                type="button"
                className={`ctx-item${n === maxLevel ? ' ctx-item-strong' : ''}`}
                onClick={() => {
                  onClose()
                  onShowLevels(n)
                }}
              >
                <span className="ctx-label">{t('appNavShowLevel', { n })}</span>
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  )
}
