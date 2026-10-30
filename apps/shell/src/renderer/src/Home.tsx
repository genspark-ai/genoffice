import { useCallback, useEffect, useLayoutEffect, useReducer, useRef, useState } from 'react'
import type { DragEvent as ReactDragEvent, ReactElement } from 'react'
import logoLockup from './assets/genoffice-logo.svg'
import { KbAskPanel } from './KbAskPanel'
import { FileBadge } from './file-badge'
import { placeFolderMenu } from './folder-menu-position'
import type {
  AccountStatus,
  FolderRoot,
  HomeApi,
  MoveConflictPolicy,
  RecentEntry,
  FileSearchPage,
  FileSearchRerank,
} from '../../shared/home-api'
import type { IntegrationsApi } from '../../shared/integrations-api'
import { useDismissablePopover } from '@genoffice/ui'
import { fileCountLabel, visiblePageCount } from './counts'
import { useI18n } from './locale'
import type { StringKey } from './locale'
import type { SettingsTarget } from './SettingsModal'
import { TemplateGallery } from './TemplateGallery'
import { TEMPLATE_GALLERY_STRINGS } from './template-gallery-strings'
import { CommandPalette } from './CommandPalette'
import { normalizeAskDeepLink, type AskQueryOrigin, type PendingAskQuery } from './ask-intent'
import { DesktopView } from './DesktopView'
import { DESKTOP_STRINGS } from './desktop-strings'
import { TodoWidget } from './TodoWidget'
import { bindRecentsChanged, onFilesChanged } from './file-events'
import { aiSetupNeeded } from './ai-setup-state'
import { AccountEntry } from './AccountEntry'
import { CloudProjectsView } from './CloudProjectsView'
import { DropToOpenOverlay } from './DropToOpenOverlay'
import { FolderPicker } from './FolderPicker'
import { ConflictPrompt, GroupPicker } from './home-dialogs'
import { FileRow, SubfolderRow, type FileRowActions } from './HomeFileRows'
import { FileSearchBox, FileSearchHeading, FileSearchResults } from './HomeSearchPanel'
import { useFolderListings } from './useFolderListings'
import {
  Chevron,
  DRAG_EXPAND_DELAY_MS,
  DRAG_PATHS_MIME,
  EmptyHint,
  FILTERS,
  FolderIcon,
  MoreDots,
  OPEN_LOCAL_EXTENSIONS,
  SortCheck,
  baseName,
  crumbsOf,
  dirOf,
  fileName,
  fixedViewportWidth,
  isUnder,
  matchesFilter,
  readTreeState,
  rootOf,
  writeTreeState,
} from './home-shared'
import { homeViewReducer, initialHomeViewState, readLaunchView, writeLaunchView } from './home-view'

import type { ProjectApi } from '@genoffice/project-store'

declare global {
  interface Window {
    aiOffice: HomeApi
    aiOfficeIntegrations?: IntegrationsApi
    /** Optional: a build without the project bridge keeps the panel in memory. */
    projectApi?: ProjectApi
  }
}

/** page size of the home list; scrolling to the bottom auto-loads the next page */
const PAGE_SIZE = 50

/** greeting sublines on the home page: one is picked at random on entry */
const GREET_ASK_KEYS = [
  'greetAsk1',
  'greetAsk2',
  'greetAsk3',
  'greetAsk4',
  'greetAsk5',
  'greetAsk6',
] as const satisfies readonly StringKey[]

// ── Main component ──────────────────────────────────────

export function Home() {
  const i18n = useI18n()
  const { t, lang } = i18n
  // ── Paged list state (rows loaded for the current view + filter) ──
  const [entries, setEntries] = useState<RecentEntry[]>([])
  /** total count under the current view + filter (not just the loaded rows) */
  const [listTotal, setListTotal] = useState(0)
  /** sidebar Recent / Starred counts under the active type filter */
  const [navCounts, setNavCounts] = useState({ recent: 0, starred: 0 })
  const [loadingMore, setLoadingMore] = useState(false)
  // ── view-mode state machine ──
  // One reducer owns the six mutually exclusive surfaces (see home-view.ts);
  // the old five hand-written flags each switch function had to clear are the
  // reducer's single transition now. The `view`/`desktopMode`/… locals below
  // are the derived names the rest of this component was written against.
  const [viewState, dispatchView] = useReducer(homeViewReducer, undefined, () =>
    initialHomeViewState(readLaunchView()),
  )
  const { surface, list: view } = viewState
  const desktopMode = surface === 'desktop'
  // Genspark web projects take over the content area (like a selected folder)
  const cloudMode = surface === 'cloud'
  // the knowledge-base ask panel takes over the content area the same way
  const askMode = surface === 'ask'
  // non-null only while the folder surface is up, exactly like the old flag set
  const selectedFolder = surface === 'folder' ? viewState.folder : null
  const [filter, setFilter] = useState('all')
  // ── the desktop's lifted wallpaper ──
  // Home mounts the backdrop host at its root while the desktop view is
  // active; DesktopView portals its wallpaper layers in, so the whole window
  // (sidebar included) sits on the wallpaper and the sidebar wears its glass
  // pane. Leaving the desktop unmounts both — the list views see none of it.
  const [backdropHost, setBackdropHost] = useState<HTMLDivElement | null>(null)
  // ── Starred groups (named buckets the Starred view can filter on) ──
  // groups that currently have at least one starred file (drives the pills)
  const [starredGroups, setStarredGroups] = useState<string[]>([])
  // the group pill filtering the Starred list; null shows every group
  const [groupFilter, setGroupFilter] = useState<string | null>(null)
  // starred rows being moved to a group (GroupPicker modal)
  const [groupPicker, setGroupPicker] = useState<string[] | null>(null)
  // ── File search (names + indexed content); active while the box has text ──
  const [searchQuery, setSearchQuery] = useState('')
  const [searchPage, setSearchPage] = useState<FileSearchPage | null>(null)
  const [rerank, setRerank] = useState<{ key: string; result: FileSearchRerank } | null>(null)
  // bumped when the Jev settings change so the current results are judged again (or the order dropped)
  const [rerankSettingsTick, setRerankSettingsTick] = useState(0)
  const [settingsRequest, setSettingsRequest] = useState<SettingsTarget | null>(null)
  const [templatesOpen, setTemplatesOpen] = useState(false)
  const [paletteOpen, setPaletteOpen] = useState(false)
  useEffect(() => window.aiOffice.onPaletteOpen?.(() => setPaletteOpen(true)), [])
  // a tab was pinned to the desktop: main asks Home to show the surface the
  // widget lives on (the pin action lands the shell on this view)
  useEffect(
    () =>
      window.aiOffice.onShowDesktop?.(() => {
        dispatchView({ type: 'show-desktop' })
        writeLaunchView('desktop')
        setSelected(new Set())
        setRowMenu(null)
      }),
    [],
  )
  // a question handed to the ask panel from outside: the palette's Ask AI
  // fallback row or a genoffice:// deep link. The seq bumps per submission so
  // the same question twice is still two handoffs; the origin decides whether
  // the panel may submit it (palette: the user's own text) or must park it in
  // the composer for an explicit Enter (deep link: web-supplied text).
  const [pendingAskQuery, setPendingAskQuery] = useState<PendingAskQuery | null>(null)
  useEffect(
    () =>
      window.aiOffice.onOpenSettings?.((target) =>
        setSettingsRequest({ section: target.section as SettingsTarget['section'] }),
      ),
    [],
  )
  const searchInputRef = useRef<HTMLInputElement>(null)
  // IME composition: wait for the committed text instead of searching each keystroke
  const composingRef = useRef(false)
  const searchSeq = useRef(0)
  const searchActive = searchQuery.trim().length > 0
  // modified-column sort (WPS-style header popover), shared by the global and folder tables
  const [fileSort, setFileSort] = useState<'recent' | 'oldest'>('recent')
  const [fileSortMenuOpen, setFileSortMenuOpen] = useState(false)
  const fileSortRef = useRef<HTMLDivElement>(null)
  const [rowMenu, setRowMenu] = useState<string | null>(null)
  // actions cell (… button + menu) of the row whose menu is open — the dismissal guard root
  const rowMenuWrapRef = useRef<HTMLSpanElement>(null)
  const [selected, setSelected] = useState<ReadonlySet<string>>(new Set())
  const [renaming, setRenaming] = useState<{ path: string; value: string } | null>(null)
  const [confirmDelete, setConfirmDelete] = useState<string[] | null>(null)
  // unavailable recent entry (missing flag) the user clicked — offer list removal
  const [confirmMissing, setConfirmMissing] = useState<RecentEntry | null>(null)
  // name in the greeting; omitted when logged out
  const [accountName, setAccountName] = useState('')
  // Genspark Projects is web-account data, so its nav entry only shows when logged in
  const [loggedIn, setLoggedIn] = useState(false)
  // single source of account state: AccountEntry reports every change (initial
  // load, login, logout), keeping the greeting name and the nav entry in sync
  const handleAccountStatus = useCallback((s: AccountStatus | null) => {
    const on = s?.loggedIn ?? false
    setLoggedIn(on)
    // web-account data leaves with the account; every other surface stays
    if (!on) dispatchView({ type: 'exit-cloud' })
    const name = on ? (s?.email ?? '').split('@')[0] : ''
    setAccountName(name ? name[0].toUpperCase() + name.slice(1) : '')
  }, [])
  // AI setup entry: shows while no provider is configured and gsk is signed
  // out; the ai:settings-changed broadcast re-checks it after any save
  const [aiSetup, setAiSetup] = useState<'needed' | 'ready'>('ready')
  const refreshAiSetup = useCallback(() => {
    let alive = true
    void window.aiOffice
      .aiGetSettings()
      .then((settings) =>
        window.aiOffice
          .aiGskStatus()
          .then((gsk) => {
            if (alive) setAiSetup(aiSetupNeeded(settings, gsk.loggedIn) ? 'needed' : 'ready')
          })
          .catch(() => undefined),
      )
      .catch(() => undefined)
    return () => {
      alive = false
    }
  }, [])
  useEffect(() => {
    const off = refreshAiSetup()
    const offSettings = window.aiOffice.onAiSettingsChanged?.(() => refreshAiSetup())
    return () => {
      off?.()
      offSettings?.()
    }
  }, [refreshAiSetup])
  const showAiSetupEntry = aiSetup === 'needed'
  const [greetAskKey] = useState(
    () => GREET_ASK_KEYS[Math.floor(Math.random() * GREET_ASK_KEYS.length)]!,
  )

  // ── Folder tree state ──
  // the default save folder first, then the folders the user added; `root` is the default one
  const [roots, setRoots] = useState<FolderRoot[]>([])
  const root = roots[0] ?? null
  const canMove = roots.some((r) => r.usable)
  const editableAt = (path: string) => rootOf(path, roots)?.usable !== false
  const [panelDrop, setPanelDrop] = useState(false)
  const [treeState] = useState(readTreeState)
  const [expanded, setExpanded] = useState<ReadonlySet<string>>(() => new Set(treeState.expanded))
  // a root without a saved layout of its own is opened once when first seen
  const seededRoot = useRef<string | null>(null)
  const {
    listings,
    load: loadFolder,
    invalidate: invalidateFolders,
    reset: resetFolders,
    tracked: trackedFolder,
  } = useFolderListings()
  // open folder menu: path + fixed-position anchor (viewport coords), so the
  // popup can escape the scrollable tree without the tree losing overflow-y
  const [folderMenu, setFolderMenu] = useState<{
    path: string
    where: 'tree' | 'table'
    x: number
    y: number
    above: number
    align: 'start' | 'end'
  } | null>(null)
  const menuOpenAt = (where: 'tree' | 'table', path: string) =>
    folderMenu?.where === where && folderMenu.path === path
  const folderMenuWrapRef = useRef<HTMLDivElement>(null)
  const folderMenuPopRef = useRef<HTMLDivElement>(null)
  const [folderRenaming, setFolderRenaming] = useState<{
    path: string
    where: 'tree' | 'table'
    value: string
  } | null>(null)
  // inline "new folder" input in the tree, under this parent
  const [creating, setCreating] = useState<{ parent: string } | null>(null)
  const [newFolderName, setNewFolderName] = useState('')
  const [confirmDeleteFolder, setConfirmDeleteFolder] = useState<string | null>(null)
  // move-to-folder picker for these paths; then the conflict prompt for the ones that collided
  const [movePicker, setMovePicker] = useState<string[] | null>(null)
  const [conflict, setConflict] = useState<{ paths: string[]; targetDir: string } | null>(null)
  // folder row currently hovered by a drag (sidebar or table)
  const [dropTarget, setDropTarget] = useState<string | null>(null)
  const dragExpandTimer = useRef<number | null>(null)

  const loadRoot = useCallback(() => {
    void window.aiOffice.folderRoots().then((next) => {
      setRoots((prev) => {
        if (prev[0] && prev[0].path !== next[0]?.path) {
          // the default save folder changed in settings: the old tree is meaningless
          resetFolders()
          dispatchView({ type: 'set-folder', dir: null })
          setExpanded(new Set(next.map((r) => r.path)))
        } else if (prev.length > 0) {
          // a folder just added opens so its contents show right away
          const known = new Set(prev.map((r) => r.path))
          const added = next.filter((r) => !known.has(r.path)).map((r) => r.path)
          if (added.length > 0) setExpanded((e) => new Set([...e, ...added]))
        }
        return next
      })
    })
  }, [resetFolders])

  useEffect(loadRoot, [loadRoot])

  useEffect(() => {
    if (root?.usable) writeTreeState({ expanded: [...expanded], root: root.path })
  }, [expanded, root])

  useEffect(() => {
    if (!root?.usable) return
    if (seededRoot.current !== root.path) {
      seededRoot.current = root.path
      if (treeState.root !== root.path) {
        // a layout saved for another root is not ours: start from the root rows
        setExpanded(new Set(roots.map((r) => r.path)))
        return
      }
    }
    for (const dir of expanded) loadFolder(dir)
  }, [root, roots, expanded, loadFolder, treeState])

  useEffect(() => {
    if (selectedFolder) loadFolder(selectedFolder)
  }, [selectedFolder, loadFolder])

  // a remembered selection that no longer exists (deleted in Finder) falls back to the root
  useEffect(() => {
    if (!selectedFolder || roots.length === 0) return
    const owner = rootOf(selectedFolder, roots)
    if (!owner?.readable) {
      dispatchView({ type: 'set-folder', dir: null })
      return
    }
    if (listings.get(selectedFolder)?.missing) {
      dispatchView({
        type: 'set-folder',
        dir: selectedFolder === owner.path ? null : dirOf(selectedFolder),
      })
    }
  }, [roots, selectedFolder, listings])

  useEffect(() => {
    return window.aiOffice.onFolderChanged((dirs) => {
      invalidateFolders(dirs.filter(trackedFolder))
    })
  }, [invalidateFolders, trackedFolder])

  // ── Paged loading ──
  // stale responses are dropped via a request sequence number (when views/filters switch quickly)
  const requestSeq = useRef(0)
  const entriesLen = useRef(0)
  entriesLen.current = entries.length

  /** reload the list; keepCount keeps the loaded row count (refresh), otherwise back to page one */
  const reload = (keepCount: boolean) => {
    const seq = ++requestSeq.current
    const ext = filter === 'all' ? undefined : filter
    // the group pills only scope the Starred view's own list; the sidebar
    // counters and the other view stay group-agnostic
    const group = view === 'starred' && groupFilter ? groupFilter : undefined
    const limit = keepCount ? Math.max(entriesLen.current, PAGE_SIZE) : PAGE_SIZE
    const primary = view === 'recent' ? window.aiOffice.recents : window.aiOffice.starred
    const secondary = view === 'recent' ? window.aiOffice.starred : window.aiOffice.recents
    void primary({ offset: 0, limit, ext, group }).then((page) => {
      if (seq !== requestSeq.current) return
      setEntries(page.entries)
      setListTotal(page.total)
      setNavCounts((prev) =>
        view === 'recent'
          ? { ...prev, recent: visiblePageCount(page) }
          : { ...prev, starred: visiblePageCount(page) },
      )
    })
    // The other view fetches only its count under the same active filter.
    void secondary({ offset: 0, limit: 0, ext }).then((page) => {
      if (seq !== requestSeq.current) return
      setNavCounts((prev) =>
        view === 'recent'
          ? { ...prev, starred: visiblePageCount(page) }
          : { ...prev, recent: visiblePageCount(page) },
      )
    })
    // the group pills live on the renderer, so their list rides along with
    // every reload; a vanished group (last file unstarred/ungrouped) clears
    // the filter instead of showing a silently empty list
    void window.aiOffice.starredGroups().then((groups) => {
      if (seq !== requestSeq.current) return
      setStarredGroups(groups)
      setGroupFilter((cur) => (cur && !groups.includes(cur) ? null : cur))
    })
  }
  const reloadRef = useRef(reload)
  reloadRef.current = reload

  /** everything on screen re-pulls: the paged list, the root and every loaded folder */
  const refresh = () => {
    reloadRef.current(true)
    loadRoot()
    invalidateFolders([...listings.keys()])
  }
  const refreshRef = useRef(refresh)
  refreshRef.current = refresh

  useEffect(() => {
    reloadRef.current(false)
  }, [view, filter, groupFilter])

  const q = searchQuery.trim()
  useEffect(() => {
    if (!q) {
      searchSeq.current++
      setSearchPage(null)
      return
    }
    let timer = 0
    const run = () => {
      if (composingRef.current) {
        timer = window.setTimeout(run, 150)
        return
      }
      const seq = ++searchSeq.current
      const ext = filter === 'all' ? undefined : filter
      void window.aiOffice.searchFiles({ q, ext, limit: 100 }).then((page) => {
        if (seq !== searchSeq.current) return
        setSearchPage(page)
        // results grow while the background index catches up
        if (page.index.pending > 0 || page.index.scanning) timer = window.setTimeout(run, 1500)
      })
    }
    timer = window.setTimeout(run, 150)
    return () => window.clearTimeout(timer)
  }, [q, filter])

  // Jev judges the top local hits once they settle; the main process answers null when reranking is off.
  // The key changes only with the query, filter or candidate set, so index polls do not restart the timer.
  const rerankKey =
    q && searchPage && searchPage.hits.length >= 2
      ? [q, filter, ...searchPage.hits.slice(0, 20).map((h) => h.path)].join('\n')
      : ''
  useEffect(() => {
    if (!rerankKey) return
    const key = rerankKey
    const paths = key.split('\n').slice(2)
    const timer = window.setTimeout(() => {
      void window.aiOffice.rerankSearch({ q, paths }).then((result) => {
        if (result) setRerank({ key, result })
        else setRerank((cur) => (cur && cur.key === key ? null : cur))
      })
    }, 600)
    return () => window.clearTimeout(timer)
  }, [q, filter, rerankKey, rerankSettingsTick])

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (!(event.metaKey || event.ctrlKey) || event.altKey || event.shiftKey) return
      if (event.key !== 'f' && event.key !== 'p') return
      const input = searchInputRef.current
      if (!input) return
      event.preventDefault()
      input.focus()
      input.select()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [])

  useEffect(() => {
    const onFocus = () => refreshRef.current()
    window.addEventListener('focus', onFocus)
    return () => window.removeEventListener('focus', onFocus)
  }, [])

  // a rename from the tab strip happens while this renderer already has focus
  useEffect(() => onFilesChanged(() => refreshRef.current()), [])

  // a save landing in any editor tab pushes recents-changed from the main
  // process; binding it here re-pulls the list immediately instead of waiting
  // for the next window focus (a save inside a tab never moves focus)
  useEffect(() => bindRecentsChanged(), [])

  const hasMore = entries.length < listTotal

  // unified dismissal: outside press, window blur, chrome press (tab strip / window drag)
  useDismissablePopover(fileSortMenuOpen, () => setFileSortMenuOpen(false), {
    inside: () => [fileSortRef.current],
  })

  const loadMore = () => {
    if (loadingMore || !hasMore) return
    setLoadingMore(true)
    const seq = requestSeq.current
    const ext = filter === 'all' ? undefined : filter
    const api = view === 'recent' ? window.aiOffice.recents : window.aiOffice.starred
    void api({ offset: entriesLen.current, limit: PAGE_SIZE, ext }).then((page) => {
      setLoadingMore(false)
      if (seq !== requestSeq.current) return
      setEntries((prev) => [...prev, ...page.entries])
      setListTotal(page.total)
    })
  }
  const loadMoreRef = useRef(loadMore)
  loadMoreRef.current = loadMore

  // oldest-first over a partially loaded list would miss the tail pages —
  // keep pulling until the list is complete (backend caps recents at 100)
  useEffect(() => {
    if (fileSort === 'oldest' && hasMore) loadMoreRef.current()
  }, [fileSort, hasMore, entries.length])

  // Load the next page once the bottom sentinel enters the viewport (240px early);
  // depending on entries.length rebuilds the observer after each page — observe fires an immediate
  // callback, so while the sentinel stays in view we keep loading until full or exhausted
  const sentinelRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const el = sentinelRef.current
    if (!el) return
    const observer = new IntersectionObserver(
      (records) => {
        if (records.some((r) => r.isIntersecting)) loadMoreRef.current()
      },
      { rootMargin: '240px' },
    )
    observer.observe(el)
    return () => observer.disconnect()
  }, [hasMore, entries.length])

  // unified dismissal: outside press, window blur, chrome press (tab strip / window drag)
  useDismissablePopover(rowMenu !== null, () => setRowMenu(null), {
    inside: () => [rowMenuWrapRef.current],
  })
  useDismissablePopover(folderMenu !== null, () => setFolderMenu(null), {
    inside: () => [folderMenuWrapRef.current],
  })
  // the fixed-position folder menu would detach from its row while the tree scrolls
  useEffect(() => {
    if (!folderMenu) return
    const close = () => setFolderMenu(null)
    window.addEventListener('scroll', close, true)
    return () => window.removeEventListener('scroll', close, true)
  }, [folderMenu])
  // open to the right of the anchor, flipped above when short of room below, clamped to the viewport
  useLayoutEffect(() => {
    const el = folderMenuPopRef.current
    if (!folderMenu || !el) return
    const { left, top } = placeFolderMenu(
      folderMenu,
      { width: el.offsetWidth, height: el.offsetHeight },
      { width: fixedViewportWidth(), height: document.documentElement.clientHeight },
    )
    el.style.left = `${left}px`
    el.style.top = `${top}px`
  }, [folderMenu])

  // Escape closes the row menu and the confirm dialogs
  useEffect(() => {
    if (
      rowMenu === null &&
      folderMenu === null &&
      confirmDelete === null &&
      confirmMissing === null &&
      confirmDeleteFolder === null
    )
      return
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setRowMenu(null)
        setFolderMenu(null)
        setConfirmDelete(null)
        setConfirmMissing(null)
        setConfirmDeleteFolder(null)
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [rowMenu, folderMenu, confirmDelete, confirmMissing, confirmDeleteFolder])

  // WPS-style sortable "modified" column header, shared by both file tables
  const renderModifiedHeader = () => (
    <div className="cloud-col-sort" ref={fileSortRef}>
      <button
        className="cloud-col-sort-btn"
        aria-haspopup="menu"
        aria-expanded={fileSortMenuOpen}
        onClick={() => setFileSortMenuOpen((o) => !o)}
      >
        {t('colModified')}
        <svg
          width="12"
          height="12"
          viewBox="0 0 16 16"
          fill="none"
          aria-hidden="true"
          style={fileSort === 'oldest' ? { transform: 'rotate(180deg)' } : undefined}
        >
          <path
            d="M8 3v10M4.5 9.5L8 13l3.5-3.5"
            stroke="currentColor"
            strokeWidth="1.4"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
      </button>
      {fileSortMenuOpen && (
        <div className="cloud-sort-menu" role="menu">
          {(['recent', 'oldest'] as const).map((key) => (
            <button
              key={key}
              className={fileSort === key ? 'active' : ''}
              role="menuitemradio"
              aria-checked={fileSort === key}
              onClick={() => {
                setFileSort(key)
                setFileSortMenuOpen(false)
              }}
            >
              <SortCheck visible={fileSort === key} />
              {t(key === 'recent' ? 'cloudSortRecent' : 'cloudSortOldest')}
            </button>
          ))}
        </div>
      )}
    </div>
  )

  // ── Plain view (no folder selected): filtering runs in the main process; entries is the visible list ──
  const selectedPaths = entries.filter((e) => selected.has(e.path)).map((e) => e.path)
  const allSelected = entries.length > 0 && selectedPaths.length === entries.length

  // folder view: files of the selected folder under the type filter; shares the same `selected` set
  const folderListing = selectedFolder ? listings.get(selectedFolder) : undefined
  const folderFiles = (folderListing?.files ?? []).filter((e) => matchesFilter(e, filter))
  const folderSubfolders = folderListing?.folders ?? []
  const folderSelectedPaths = folderFiles.filter((e) => selected.has(e.path)).map((e) => e.path)
  const folderAllSelected =
    folderFiles.length > 0 && folderSelectedPaths.length === folderFiles.length

  // ── surface switching: one reducer transition each, plus the row-UI resets
  // the hand-written switchers also performed (selection / menus are shared
  // row state, not view-mode state, so they stay here with the wiring) ──

  const changeView = (next: 'recent' | 'starred') => {
    dispatchView({ type: 'show-list', list: next })
    writeLaunchView(next)
    setSelected(new Set())
    setRowMenu(null)
    setGroupFilter(null)
  }

  /** the desktop nav item: the icon desktop takes over the content area */
  const showDesktop = () => {
    dispatchView({ type: 'show-desktop' })
    writeLaunchView('desktop')
    setSelected(new Set())
    setRowMenu(null)
  }

  const changeFilter = (key: string) => {
    setFilter(key)
    setSelected(new Set())
    setRowMenu(null)
  }

  const changeGroupFilter = (group: string | null) => {
    setGroupFilter(group)
    setSelected(new Set())
    setRowMenu(null)
  }

  const selectFolder = useCallback((dir: string) => {
    dispatchView({ type: 'select-folder', dir })
    setSelected(new Set())
    setRowMenu(null)
    setFolderMenu(null)
  }, [])

  /** the desktop's AI-ask tile: same switch the sidebar's Ask AI nav performs */
  const openAskView = useCallback(() => {
    dispatchView({ type: 'show-ask' })
    writeLaunchView('ask')
    setSelected(new Set())
    setRowMenu(null)
  }, [])
  /** the ask-intent handoff's shared core: bump the seq and switch to the ask view */
  const handAskQuery = useCallback(
    (query: string, origin: AskQueryOrigin) => {
      const text = query.trim()
      if (!text) return
      setPendingAskQuery((prev) => ({ text, seq: (prev?.seq ?? 0) + 1, origin }))
      openAskView()
    },
    [openAskView],
  )
  /** the intent bus's fallback route: switch to the ask view and hand it the question */
  const submitAskQuery = useCallback(
    (query: string) => handAskQuery(query, 'palette'),
    [handAskQuery],
  )
  // genoffice://ask?q=… and friends: the main process forwards deep links over
  // the home:deep-link channel. Softened because a preload from before the
  // contract shipped simply has no onDeepLink — the palette path above works
  // regardless. The query is web-supplied (any page can location.href a link),
  // so it is handed over with the deep-link origin: it parks in the ask
  // composer as a draft and only an explicit Enter sends it.
  useEffect(() => {
    const api = window.aiOffice as Partial<HomeApi>
    return api.onDeepLink?.((payload) => {
      const link = normalizeAskDeepLink(payload)
      if (!link) return
      if (link.query.trim()) handAskQuery(link.query, 'deep-link')
      else openAskView()
    })
  }, [handAskQuery, openAskView])
  const openTemplates = useCallback(() => setTemplatesOpen(true), [])
  const openSettings = useCallback((target: SettingsTarget) => setSettingsRequest(target), [])
  const openFolderView = useCallback((root: FolderRoot) => selectFolder(root.path), [selectFolder])

  const toggleExpanded = (dir: string) => {
    setExpanded((prev) => {
      const next = new Set(prev)
      if (next.has(dir)) next.delete(dir)
      else next.add(dir)
      return next
    })
  }

  const expandTo = (dir: string) => {
    const owner = rootOf(dir, roots)
    if (!owner) return
    setExpanded((prev) => {
      const next = new Set(prev)
      for (const crumb of crumbsOf(owner, dir)) next.add(crumb.path)
      return next
    })
  }

  const toggleSelect = (path: string, on: boolean) => {
    setSelected((prev) => {
      const next = new Set(prev)
      if (on) next.add(path)
      else next.delete(path)
      return next
    })
  }

  const toggleSelectAll = () => {
    setSelected(allSelected ? new Set() : new Set(entries.map((e) => e.path)))
  }

  const toggleSelectAllFolder = () => {
    setSelected(folderAllSelected ? new Set() : new Set(folderFiles.map((e) => e.path)))
  }

  const toggleStar = (path: string) => {
    void window.aiOffice.toggleStar(path).then(refresh)
  }

  const removeRecent = (paths: string[]) => {
    setRowMenu(null)
    setSelected(new Set())
    void window.aiOffice.removeRecent(paths).then(refresh)
  }

  /** the Starred view's bulk action: drop the stars, keep the recents entries */
  const unstarSelected = (paths: string[]) => {
    setRowMenu(null)
    setSelected(new Set())
    void window.aiOffice.unstarPaths(paths).then(refresh)
  }

  /** put rows into a group (null = out of their group); the pills list and
   *  the filtered list come back with the refresh */
  const assignGroup = (paths: string[], group: string | null) => {
    setGroupPicker(null)
    setRowMenu(null)
    void window.aiOffice.setStarredGroup(paths, group).then(refresh)
  }

  const deleteFiles = (paths: string[]) => {
    setRowMenu(null)
    setConfirmDelete(paths)
  }

  const confirmDeleteNow = () => {
    const paths = confirmDelete ?? []
    setConfirmDelete(null)
    setSelected(new Set())
    void window.aiOffice.deleteFiles(paths).then(refresh)
  }

  const duplicateFile = (path: string) => {
    setRowMenu(null)
    void window.aiOffice.duplicateFile(path).then(refresh)
  }

  const startRename = (entry: RecentEntry) => {
    setRowMenu(null)
    setRenaming({ path: entry.path, value: baseName(entry) })
  }

  const commitRename = (entry: RecentEntry) => {
    const value = renaming?.value.trim() ?? ''
    setRenaming(null)
    if (!value || value === baseName(entry)) return
    const newName = entry.ext ? `${value}.${entry.ext}` : value
    void window.aiOffice.renameFile(entry.path, newName).then((result) => {
      if (!result.ok) window.alert(result.error ?? t('renameFailed'))
      refresh()
    })
  }

  // ── Folder actions ──

  const startCreateFolder = (parent: string) => {
    setFolderMenu(null)
    expandTo(parent)
    setCreating({ parent })
    setNewFolderName('')
  }

  // inline inputs commit on Enter or blur and cancel on Escape; the blur an
  // unmount may fire reads the edit from refs mirrored during render, so a
  // finished or cancelled edit is a no-op (see FolderPicker.commitCreate)
  const creatingFolderRef = useRef<{ parent: string; name: string } | null>(null)
  creatingFolderRef.current = creating ? { parent: creating.parent, name: newFolderName } : null
  const folderRenamingRef = useRef(folderRenaming)
  folderRenamingRef.current = folderRenaming

  const commitCreateFolder = async () => {
    const pending = creatingFolderRef.current
    creatingFolderRef.current = null
    const name = pending?.name.trim()
    setCreating(null)
    setNewFolderName('')
    if (!pending || !name) return
    const result = await window.aiOffice.createFolder(pending.parent, name)
    if (!result.ok) {
      window.alert(result.error ?? t('renameFailed'))
      return
    }
    invalidateFolders([pending.parent])
  }

  const startRenameFolder = (entry: { path: string; name: string }, where: 'tree' | 'table') => {
    setFolderMenu(null)
    setRowMenu(null)
    setFolderRenaming({ path: entry.path, where, value: entry.name })
  }

  const commitRenameFolder = async () => {
    const pending = folderRenamingRef.current
    folderRenamingRef.current = null
    setFolderRenaming(null)
    if (!pending) return
    const value = pending.value.trim()
    if (!value || value === fileName(pending.path)) return
    const result = await window.aiOffice.renameFolder(pending.path, value)
    if (!result.ok) {
      window.alert(result.error ?? t('renameFailed'))
      return
    }
    if (result.path) {
      const renamed = result.path
      // selection / expansion follow the renamed folder (and anything inside it)
      const rebase = (p: string) =>
        p === pending.path || isUnder(pending.path, p) ? renamed + p.slice(pending.path.length) : p
      setExpanded((prev) => new Set([...prev].map(rebase)))
      if (selectedFolder) dispatchView({ type: 'set-folder', dir: rebase(selectedFolder) })
    }
    refresh()
  }

  const confirmDeleteFolderNow = async () => {
    const dir = confirmDeleteFolder
    setConfirmDeleteFolder(null)
    if (!dir) return
    await window.aiOffice.deleteFolder(dir)
    if (selectedFolder && (selectedFolder === dir || isUnder(dir, selectedFolder))) {
      dispatchView({ type: 'set-folder', dir: dirOf(dir) })
    }
    refresh()
  }

  const startMove = (paths: string[]) => {
    setRowMenu(null)
    setFolderMenu(null)
    if (paths.length > 0) setMovePicker(paths)
  }

  const doMove = async (paths: string[], targetDir: string, policy: MoveConflictPolicy) => {
    setMovePicker(null)
    setConflict(null)
    setSelected(new Set())
    const result = await window.aiOffice.movePaths(paths, targetDir, policy)
    if (result.moved.length > 0 && selectedFolder) {
      // a moved folder that held the selection drags the selection along
      for (const { from, to } of result.moved) {
        if (selectedFolder === from || isUnder(from, selectedFolder)) {
          dispatchView({ type: 'set-folder', dir: to + selectedFolder.slice(from.length) })
        }
      }
    }
    refresh()
    if (result.failed.length > 0) window.alert(result.failed[0].error)
    if (result.conflicts.length > 0) setConflict({ paths: result.conflicts, targetDir })
  }

  // ── Drag & drop (rows → folder rows) ──

  const dragPathsFor = (path: string, context: 'global' | 'folder'): string[] => {
    const pool = context === 'folder' ? folderSelectedPaths : selectedPaths
    return pool.includes(path) ? pool : [path]
  }

  const onRowDragStart = (event: ReactDragEvent, paths: string[]) => {
    event.dataTransfer.setData(DRAG_PATHS_MIME, JSON.stringify(paths))
    event.dataTransfer.effectAllowed = 'move'
    setRowMenu(null)
  }

  const readDragPaths = (event: ReactDragEvent): string[] => {
    try {
      const raw = JSON.parse(event.dataTransfer.getData(DRAG_PATHS_MIME)) as unknown
      return Array.isArray(raw) ? raw.filter((p): p is string => typeof p === 'string') : []
    } catch {
      return []
    }
  }

  const clearDragExpand = () => {
    if (dragExpandTimer.current !== null) {
      window.clearTimeout(dragExpandTimer.current)
      dragExpandTimer.current = null
    }
  }

  const folderDropProps = (dir: string, { autoExpand }: { autoExpand: boolean }) => ({
    onDragOver: (event: ReactDragEvent) => {
      if (rootOf(dir, roots)?.usable === false) return
      if (!event.dataTransfer.types.includes(DRAG_PATHS_MIME)) return
      event.preventDefault()
      event.dataTransfer.dropEffect = 'move'
      if (dropTarget !== dir) {
        setDropTarget(dir)
        clearDragExpand()
        if (autoExpand && !expanded.has(dir)) {
          dragExpandTimer.current = window.setTimeout(() => {
            setExpanded((prev) => new Set([...prev, dir]))
          }, DRAG_EXPAND_DELAY_MS)
        }
      }
    },
    onDragLeave: (event: ReactDragEvent) => {
      if (event.currentTarget.contains(event.relatedTarget as Node | null)) return
      if (dropTarget === dir) setDropTarget(null)
      clearDragExpand()
    },
    onDrop: (event: ReactDragEvent) => {
      event.preventDefault()
      setDropTarget(null)
      clearDragExpand()
      const paths = readDragPaths(event).filter((p) => p !== dir && !isUnder(p, dir))
      if (paths.length > 0) void doMove(paths, dir, 'ask')
    },
  })

  // ── New file (lands in the selected folder) ──
  const newFileOpts =
    selectedFolder && root && selectedFolder !== root.path && rootOf(selectedFolder, roots)?.usable
      ? { dir: selectedFolder }
      : undefined

  const NEW_ITEMS = [
    {
      ext: 'docx',
      title: t('newDoc'),
      sub: '.docx',
      action: () => window.aiOffice.newDoc(newFileOpts),
    },
    {
      ext: 'xlsx',
      title: t('newSheet'),
      sub: '.xlsx',
      action: () => window.aiOffice.newSheet(newFileOpts),
    },
    {
      ext: 'pptx',
      title: t('newSlide'),
      sub: '.pptx',
      action: () => window.aiOffice.newSlide(newFileOpts),
    },
    {
      ext: 'md',
      title: t('newMarkdown'),
      sub: '.md',
      action: () => window.aiOffice.newMarkdown(newFileOpts),
    },
    {
      ext: 'html',
      title: t('newHtml'),
      sub: '.html',
      action: () => window.aiOffice.newHtml(newFileOpts),
    },
    {
      ext: 'pdf',
      title: t('newPdf'),
      sub: '.pdf',
      action: () => window.aiOffice.newPdf(newFileOpts),
    },
  ]

  function renderQuickCards() {
    return (
      <div className="quick-cards">
        {NEW_ITEMS.map((item) => (
          <button key={item.ext} className="quick-card" onClick={() => void item.action()}>
            <FileBadge ext={item.ext} size={30} />
            <span className="quick-text">
              <span className="quick-title-row">
                <span className="quick-title">{item.title}</span>
                <span className="ai-chip">AI</span>
              </span>
              <span className="quick-sub">{item.sub}</span>
            </span>
          </button>
        ))}
        <button
          className="quick-card"
          onClick={() => setTemplatesOpen(true)}
          data-tip={TEMPLATE_GALLERY_STRINGS[lang].subtitle}
        >
          <span className="quick-folder">
            <FileBadge ext="docx" size={18} />
          </span>
          <span className="quick-text">
            <span className="quick-title-row">
              <span className="quick-title">{TEMPLATE_GALLERY_STRINGS[lang].title}</span>
            </span>
            <span className="quick-sub">{TEMPLATE_GALLERY_STRINGS[lang].subtitle}</span>
          </span>
        </button>
        <button
          className="quick-card"
          onClick={() => void window.aiOffice.browse()}
          data-tip={OPEN_LOCAL_EXTENSIONS}
        >
          <span className="quick-folder">
            <FolderIcon size={18} />
          </span>
          <span className="quick-text">
            <span className="quick-title-row">
              <span className="quick-title">{t('openLocal')}</span>
            </span>
            <span className="quick-sub">{OPEN_LOCAL_EXTENSIONS}</span>
          </span>
        </button>
      </div>
    )
  }

  // ── Sidebar folder tree ──

  const addFolderRoot = () => {
    void window.aiOffice.addFolderRoot().then((added) => {
      if (added) loadRoot()
    })
  }

  /** the folder leaves the list only; whatever it held on disk stays where it is */
  const removeFolderRoot = (path: string) => {
    setFolderMenu(null)
    if (selectedFolder && isUnder(path, selectedFolder)) {
      dispatchView({ type: 'set-folder', dir: null })
    }
    setExpanded((prev) => new Set([...prev].filter((dir) => !isUnder(path, dir))))
    void window.aiOffice.removeFolderRoot(path).then(loadRoot)
  }

  // folders dragged in from the OS join the tree in place; documents open as they do anywhere else
  const panelDropProps = {
    onDragOver: (event: ReactDragEvent) => {
      if (!event.dataTransfer.types.includes('Files')) return
      event.preventDefault()
      event.dataTransfer.dropEffect = 'link'
      if (!panelDrop) setPanelDrop(true)
    },
    onDragLeave: (event: ReactDragEvent) => {
      if (event.currentTarget.contains(event.relatedTarget as Node | null)) return
      setPanelDrop(false)
    },
    onDrop: (event: ReactDragEvent) => {
      if (!event.dataTransfer.types.includes('Files')) return
      event.preventDefault()
      setPanelDrop(false)
      const paths = Array.from(event.dataTransfer.files)
        .map((file) => window.aiOffice.pathForFile(file))
        .filter(Boolean)
      if (paths.length === 0) return
      void window.aiOffice.dropFolderRoots(paths).then((added) => {
        if (added.length > 0) loadRoot()
      })
    },
  }

  const renderFolderMenu = (entry: { path: string; name: string }, rootEntry?: FolderRoot) => (
    <div
      className="folder-menu-wrap"
      ref={menuOpenAt('tree', entry.path) ? folderMenuWrapRef : undefined}
    >
      <button
        className="folder-more-btn"
        aria-label={t('folderMoreActions', { name: entry.name })}
        aria-expanded={menuOpenAt('tree', entry.path)}
        onClick={(e) => {
          e.stopPropagation()
          if (menuOpenAt('tree', entry.path)) {
            setFolderMenu(null)
            return
          }
          const rect = e.currentTarget.getBoundingClientRect()
          setFolderMenu({
            path: entry.path,
            where: 'tree',
            x: rect.right,
            y: rect.bottom + 4,
            above: rect.top - 4,
            align: 'end',
          })
        }}
      >
        <MoreDots />
      </button>
      {folderMenu && menuOpenAt('tree', entry.path) && (
        <div className="folder-menu" role="menu" ref={folderMenuPopRef}>
          {editableAt(entry.path) && (
            <button role="menuitem" onClick={() => startCreateFolder(entry.path)}>
              {t('newSubfolder')}
            </button>
          )}
          {!rootEntry && editableAt(entry.path) && (
            <button role="menuitem" onClick={() => startRenameFolder(entry, 'tree')}>
              {t('rename')}
            </button>
          )}
          {!rootEntry && editableAt(entry.path) && (
            <button role="menuitem" onClick={() => startMove([entry.path])}>
              {t('moveToFolder')}
            </button>
          )}
          {(rootEntry?.readable ?? true) && (
            <button
              role="menuitem"
              onClick={() => {
                setFolderMenu(null)
                void window.aiOffice.revealPath(entry.path)
              }}
            >
              {t('revealInFolder')}
            </button>
          )}
          {!rootEntry && editableAt(entry.path) && (
            <>
              <div className="row-menu-divider" />
              <button
                role="menuitem"
                className="danger"
                onClick={() => {
                  setFolderMenu(null)
                  setConfirmDeleteFolder(entry.path)
                }}
              >
                {t('deleteFolder')}
              </button>
            </>
          )}
          {rootEntry?.removable && (
            <>
              <div className="row-menu-divider" />
              <button role="menuitem" onClick={() => removeFolderRoot(entry.path)}>
                {t('removeFolderRoot')}
              </button>
            </>
          )}
        </div>
      )}
    </div>
  )

  const renderNewFolderInput = (depth: number) => (
    <li className="tree-item">
      <div className="tree-row" style={{ paddingLeft: 8 + depth * 14 }}>
        <span className="tree-chevron" aria-hidden="true" />
        <span className="tree-icon" aria-hidden="true">
          <FolderIcon />
        </span>
        <input
          className="folder-rename-input inline"
          autoFocus
          placeholder={t('untitledFolder')}
          value={newFolderName}
          onChange={(e) => setNewFolderName(e.target.value)}
          onBlur={() => void commitCreateFolder()}
          onKeyDown={(e) => {
            e.stopPropagation()
            if (e.nativeEvent.isComposing) return
            if (e.key === 'Enter') void commitCreateFolder()
            if (e.key === 'Escape') {
              setCreating(null)
              setNewFolderName('')
            }
          }}
        />
      </div>
    </li>
  )

  function renderTreeNode(
    entry: { path: string; name: string; hasSubfolders: boolean },
    depth: number,
  ): ReactElement {
    const rootEntry = roots.find((r) => r.path === entry.path)
    const isRoot = rootEntry !== undefined
    const unavailable = rootEntry !== undefined && !rootEntry.readable
    const isOpen = expanded.has(entry.path) && !unavailable
    const children = listings.get(entry.path)?.folders ?? []
    const isActive = selectedFolder === entry.path
    const isRenaming = folderRenaming?.where === 'tree' && folderRenaming.path === entry.path
    const showChevron = !unavailable && (isRoot || entry.hasSubfolders || children.length > 0)
    return (
      <li key={entry.path} className="tree-item">
        <div
          className={`tree-row${isActive ? ' active' : ''}${dropTarget === entry.path ? ' drop-target' : ''}${unavailable ? ' unavailable' : ''}`}
          style={{ paddingLeft: 8 + depth * 14 }}
          role="treeitem"
          aria-selected={isActive}
          aria-expanded={showChevron ? isOpen : undefined}
          tabIndex={0}
          title={isRoot ? entry.path : undefined}
          onClick={() => {
            if (!unavailable) selectFolder(entry.path)
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !unavailable) selectFolder(entry.path)
            if (e.key === 'ArrowRight' && !isOpen) toggleExpanded(entry.path)
            if (e.key === 'ArrowLeft' && isOpen) toggleExpanded(entry.path)
          }}
          onContextMenu={(e) => {
            e.preventDefault()
            setFolderMenu({
              path: entry.path,
              where: 'tree',
              x: e.clientX + 2,
              y: e.clientY + 2,
              above: e.clientY + 2,
              align: 'start',
            })
          }}
          draggable={!isRoot && !isRenaming}
          onDragStart={(e) => onRowDragStart(e, [entry.path])}
          {...folderDropProps(entry.path, { autoExpand: true })}
        >
          <button
            className="tree-chevron"
            tabIndex={-1}
            aria-hidden="true"
            style={{ visibility: showChevron ? undefined : 'hidden' }}
            onClick={(e) => {
              e.stopPropagation()
              toggleExpanded(entry.path)
            }}
          >
            <Chevron open={isOpen} />
          </button>
          <span className="tree-icon" aria-hidden="true">
            <FolderIcon open={isOpen} />
          </span>
          {isRenaming ? (
            <input
              className="folder-rename-input inline"
              value={folderRenaming.value}
              autoFocus
              onFocus={(e) => e.target.select()}
              onClick={(e) => e.stopPropagation()}
              onChange={(e) =>
                setFolderRenaming({ path: entry.path, where: 'tree', value: e.target.value })
              }
              onBlur={() => void commitRenameFolder()}
              onKeyDown={(e) => {
                e.stopPropagation()
                if (e.nativeEvent.isComposing) return
                if (e.key === 'Enter') void commitRenameFolder()
                if (e.key === 'Escape') setFolderRenaming(null)
              }}
            />
          ) : (
            <span className="tree-name">{entry.name}</span>
          )}
          {unavailable && <span className="tree-hint">{t('rootUnavailable')}</span>}
          {renderFolderMenu(entry, rootEntry)}
        </div>
        {isOpen && (children.length > 0 || creating?.parent === entry.path) && (
          <ul className="tree-children" role="group">
            {creating?.parent === entry.path && renderNewFolderInput(depth + 1)}
            {children.map((child) => renderTreeNode(child, depth + 1))}
          </ul>
        )}
      </li>
    )
  }

  function renderFolderPanel() {
    const createIn = selectedFolder ?? root?.path
    const canCreate = createIn !== undefined && rootOf(createIn, roots)?.usable === true
    return (
      <div className={`folder-panel${panelDrop ? ' drop-target' : ''}`} {...panelDropProps}>
        <div className="folder-panel-head">
          <span className="folder-panel-title">{t('folders')}</span>
          <div className="folder-panel-actions">
            {roots.length > 0 && (
              <button
                className="folder-add-btn folder-add-root-btn"
                data-tip={t('addFolderRoot')}
                aria-label={t('addFolderRoot')}
                onClick={addFolderRoot}
              >
                <svg
                  width="14"
                  height="14"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2.4"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  aria-hidden="true"
                >
                  <path d="M2 9V5a2 2 0 0 1 2-2h3.9a2 2 0 0 1 1.69.9l.81 1.2a2 2 0 0 0 1.67.9H20a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2v-1" />
                  <path d="M2 13h10" />
                  <path d="m9 16 3-3-3-3" />
                </svg>
              </button>
            )}
            {canCreate && (
              <button
                className="folder-add-btn folder-new-btn"
                data-tip={t('newFolder')}
                aria-label={t('newFolder')}
                onClick={() => startCreateFolder(createIn)}
              >
                <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
                  <path
                    d="M7 1v12M1 7h12"
                    stroke="currentColor"
                    strokeWidth="1.7"
                    strokeLinecap="round"
                  />
                </svg>
              </button>
            )}
          </div>
        </div>
        <ul className="tree" role="tree">
          {root && !root.usable ? (
            <li className="folder-unusable">
              <p>{t('rootUnusable')}</p>
              <button
                className="btn btn-secondary"
                onClick={() => void window.aiOffice.pickDefaultSaveDir().then(() => loadRoot())}
              >
                {t('pickSaveDir')}
              </button>
            </li>
          ) : (
            root && renderTreeNode({ path: root.path, name: root.name, hasSubfolders: true }, 0)
          )}
          {roots
            .slice(1)
            .map((r) =>
              renderTreeNode({ path: r.path, name: r.name, hasSubfolders: r.readable }, 0),
            )}
        </ul>
      </div>
    )
  }

  // ── File row rendering (shared by the plain view and the folder view) ──

  // ── the row components' action bundle: Home owns the state, the rows in
  // HomeFileRows.tsx own the markup (one bundle object so both row kinds and
  // both tables share it without a prop per callback) ──
  const rowActions: FileRowActions = {
    openPath: (path) => void window.aiOffice.openPath(path),
    reportMissing: setConfirmMissing,
    toggleSelect,
    toggleStar,
    dragStart: onRowDragStart,
    startRename,
    setRenameValue: (path, value) => setRenaming({ path, value }),
    commitRename,
    cancelRename: () => setRenaming(null),
    setRowMenu,
    startMove,
    duplicateFile,
    unstar: unstarSelected,
    removeRecent,
    deleteFiles,
    pickGroup: (paths) => setGroupPicker(paths),
    dropProps: folderDropProps,
    expandTo,
    selectFolder,
    setFolderRenaming: setFolderRenaming,
    commitRenameFolder,
    setFolderMenu,
    startRenameFolder,
    confirmDeleteFolder: setConfirmDeleteFolder,
  }

  // ── Empty-state card (shared markup in home-shared.tsx) ──
  const renderEmpty = (hint: string) => <EmptyHint hint={hint} />

  // ── Search panel wiring (markup in HomeSearchPanel.tsx) ──

  const clearSearch = () => {
    setSearchQuery('')
    searchInputRef.current?.focus()
  }

  // the search box and results render through HomeSearchPanel.tsx; the query
  // state, debounce and rerank timer stay here
  const rerankApplied = rerank && rerank.key === rerankKey ? rerank.result : null
  const renderSearchBox = () => (
    <FileSearchBox
      searchQuery={searchQuery}
      onSearchQueryChange={setSearchQuery}
      searchActive={searchActive}
      clearSearch={clearSearch}
      searchInputRef={searchInputRef}
      composingRef={composingRef}
      onOpenSearchSettings={() => setSettingsRequest({ section: 'aiMedia', block: 'rerank' })}
    />
  )

  const renderSearchResults = () => (
    <FileSearchResults searchPage={searchPage} q={q} rerankApplied={rerankApplied} roots={roots} />
  )

  const renderSearchHeading = () => (
    <FileSearchHeading total={searchPage?.total ?? 0} rerankApplied={rerankApplied} />
  )

  function renderFolderContent() {
    if (roots.length === 0) return null
    const total = folderSubfolders.length + folderFiles.length
    const sortedFiles = fileSort === 'oldest' ? [...folderFiles].reverse() : folderFiles
    return (
      <main className="content">
        <section className="quick-start" aria-label={t('secQuickStart')}>
          <div className="section-head">
            <span className="section-label">{t('secQuickStart')}</span>
          </div>
          {renderQuickCards()}
        </section>

        <section className="recents" aria-label={t('folders')}>
          <div className="recents-toolbar">
            {folderSelectedPaths.length > 0 ? (
              <div className="selection-bar">
                <span className="selection-count">
                  {t('selectedCount', { n: folderSelectedPaths.length })}
                </span>
                {selectedFolder && editableAt(selectedFolder) && (
                  <>
                    <button
                      className="selection-action"
                      onClick={() => startMove(folderSelectedPaths)}
                    >
                      {t('moveToFolder')}
                    </button>
                    <button
                      className="selection-action danger"
                      onClick={() => deleteFiles(folderSelectedPaths)}
                    >
                      {t('deleteFiles')}
                    </button>
                  </>
                )}
                <button className="selection-action" onClick={() => setSelected(new Set())}>
                  {t('cancel')}
                </button>
              </div>
            ) : (
              <div className="filter-pills" role="tablist" aria-label={t('filterAria')}>
                {FILTERS.map((f) => (
                  <button
                    key={f.key}
                    className={`filter-pill${filter === f.key ? ' active' : ''}`}
                    onClick={() => changeFilter(f.key)}
                  >
                    {t(f.label)}
                  </button>
                ))}
              </div>
            )}
            {renderSearchBox()}
            {searchActive ? (
              renderSearchHeading()
            ) : (
              <div className="recents-heading folder-heading">
                <span className="file-count">
                  {t(total === 1 ? 'itemCountOne' : 'itemCount', { n: total })}
                </span>
              </div>
            )}
          </div>

          {searchActive ? (
            renderSearchResults()
          ) : total === 0 ? (
            renderEmpty(filter === 'all' ? t('emptyFolder') : t('emptyFiltered'))
          ) : (
            <div
              className={`recent-table${folderSelectedPaths.length > 0 ? ' has-selection' : ''}`}
            >
              <div className="recent-columns">
                <span className="col-check">
                  <input
                    type="checkbox"
                    checked={folderAllSelected}
                    disabled={folderFiles.length === 0}
                    onChange={toggleSelectAllFolder}
                    aria-label={t('selectAll')}
                  />
                </span>
                <span className="col-name">{t('colName')}</span>
                <span className="col-path">{t('colLocation')}</span>
                {renderModifiedHeader()}
                <span className="col-size">{t('colSize')}</span>
                <span />
                <span />
              </div>
              <ul className="recent-list">
                {folderSubfolders.map((entry) => (
                  <SubfolderRow
                    key={entry.path}
                    entry={entry}
                    dropTarget={dropTarget}
                    renaming={folderRenaming}
                    folderMenuOpen={folderMenu?.where === 'table' && folderMenu.path === entry.path}
                    folderMenuWrapRef={folderMenuWrapRef}
                    folderMenuPopRef={folderMenuPopRef}
                    roots={roots}
                    actions={rowActions}
                  />
                ))}
                {sortedFiles.map((entry) => (
                  <FileRow
                    key={entry.path}
                    entry={entry}
                    context="folder"
                    dragPool={dragPathsFor(entry.path, 'folder')}
                    selected={selected}
                    renaming={renaming}
                    rowMenu={rowMenu}
                    rowMenuWrapRef={rowMenuWrapRef}
                    roots={roots}
                    view={view}
                    canMove={canMove}
                    canDelete={folderSelectedPaths.length === 0}
                    actions={rowActions}
                  />
                ))}
              </ul>
            </div>
          )}
        </section>
      </main>
    )
  }

  // ── Plain view ────────────────────────────────────────

  function renderGlobalContent() {
    const now = new Date()
    const hour = now.getHours()
    const greetKey =
      hour < 6
        ? 'greetEvening'
        : hour < 12
          ? 'greetMorning'
          : hour < 18
            ? 'greetAfternoon'
            : 'greetEvening'
    const cjk = lang === 'zh' || lang === 'zh-TW' || lang === 'ja'
    const greeting = `${t(greetKey)}${accountName ? (cjk ? '，' : ', ') + accountName : ''}${cjk ? '。' : '. '}`
    return (
      <main className="content">
        <section className="quick-start" aria-label={t('secQuickStart')}>
          <div className="home-hero">
            <h1 className="hero-title">
              {greeting}
              <span className="hero-ask">{t(greetAskKey)}</span>
            </h1>
          </div>
          {showAiSetupEntry && (
            <button
              className="ai-setup-entry"
              onClick={() => setSettingsRequest({ section: 'aiModel' })}
            >
              <span className="ai-setup-icon" aria-hidden="true">
                <svg width="18" height="18" viewBox="0 0 16 16" fill="none">
                  <path
                    d="M8 1.8 9.5 6l4.2 1.5L9.5 9 8 13.2 6.5 9 2.3 7.5 6.5 6 8 1.8Z"
                    stroke="currentColor"
                    strokeWidth="1.3"
                    strokeLinejoin="round"
                  />
                  <path
                    d="M12.8 11.2v3M11.3 12.7h3"
                    stroke="currentColor"
                    strokeWidth="1.3"
                    strokeLinecap="round"
                  />
                </svg>
              </span>
              <span className="ai-setup-copy">
                <span className="ai-setup-title">{t('homeAiSetupTitle')}</span>
                <span className="ai-setup-body">{t('homeAiSetupBody')}</span>
              </span>
              <span className="ai-setup-action">{t('homeAiSetupAction')}</span>
            </button>
          )}
          {renderQuickCards()}
        </section>

        <section
          className="recents"
          aria-label={view === 'recent' ? t('secRecent') : t('secStarred')}
        >
          <div className="recents-toolbar">
            {selectedPaths.length > 0 ? (
              <div className="selection-bar">
                <span className="selection-count">
                  {t('selectedCount', { n: selectedPaths.length })}
                </span>
                {canMove && (
                  <button className="selection-action" onClick={() => startMove(selectedPaths)}>
                    {t('moveToFolder')}
                  </button>
                )}
                {view === 'starred' ? (
                  // every Starred-view row is a favorite, so the selection
                  // action unstars; removeRecent there kept the stars and the
                  // rows came straight back on refresh
                  <button
                    className="selection-action"
                    onClick={() => unstarSelected(selectedPaths)}
                  >
                    {t('unstar')}
                  </button>
                ) : (
                  <button className="selection-action" onClick={() => removeRecent(selectedPaths)}>
                    {t('removeFromList')}
                  </button>
                )}
                <button
                  className="selection-action danger"
                  onClick={() => deleteFiles(selectedPaths)}
                >
                  {t('deleteFiles')}
                </button>
                <button className="selection-action" onClick={() => setSelected(new Set())}>
                  {t('cancel')}
                </button>
              </div>
            ) : (
              <div className="toolbar-filters">
                <div className="filter-pills" role="tablist" aria-label={t('filterAria')}>
                  {FILTERS.map((f) => (
                    <button
                      key={f.key}
                      className={`filter-pill${filter === f.key ? ' active' : ''}`}
                      onClick={() => changeFilter(f.key)}
                    >
                      {t(f.label)}
                    </button>
                  ))}
                </div>
                {view === 'starred' && starredGroups.length > 0 && (
                  // the group pills refine the Starred view further; the type
                  // pills above them keep working on the group-scoped list
                  <div className="filter-pills" role="tablist" aria-label={t('groupFilterAria')}>
                    <button
                      className={`filter-pill${groupFilter === null ? ' active' : ''}`}
                      onClick={() => changeGroupFilter(null)}
                    >
                      {t('filterAll')}
                    </button>
                    {starredGroups.map((g) => (
                      <button
                        key={g}
                        className={`filter-pill${groupFilter === g ? ' active' : ''}`}
                        onClick={() => changeGroupFilter(g)}
                      >
                        {g}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            )}
            {renderSearchBox()}
            {searchActive ? (
              renderSearchHeading()
            ) : (
              <div className="recents-heading">
                <span className="section-label">
                  {view === 'recent' ? t('secRecent') : t('secStarred')}
                </span>
                <span className="file-count">{fileCountLabel(listTotal, lang, t)}</span>
              </div>
            )}
          </div>

          {searchActive ? (
            renderSearchResults()
          ) : entries.length === 0 ? (
            renderEmpty(
              view === 'starred'
                ? t('emptyStarred')
                : navCounts.recent === 0
                  ? t('emptyRecent')
                  : t('emptyFiltered'),
            )
          ) : (
            <div className={`recent-table${selectedPaths.length > 0 ? ' has-selection' : ''}`}>
              <div className="recent-columns">
                <span className="col-check">
                  <input
                    type="checkbox"
                    checked={allSelected}
                    onChange={toggleSelectAll}
                    aria-label={t('selectAll')}
                  />
                </span>
                <span className="col-name">{t('colName')}</span>
                <span className="col-path">{t('colLocation')}</span>
                {renderModifiedHeader()}
                <span className="col-size">{t('colSize')}</span>
                <span />
                <span />
              </div>
              <ul className="recent-list">
                {(fileSort === 'oldest' ? [...entries].reverse() : entries).map((entry) => (
                  <FileRow
                    key={entry.path}
                    entry={entry}
                    context="global"
                    dragPool={dragPathsFor(entry.path, 'global')}
                    selected={selected}
                    renaming={renaming}
                    rowMenu={rowMenu}
                    rowMenuWrapRef={rowMenuWrapRef}
                    roots={roots}
                    view={view}
                    canMove={canMove}
                    canDelete={selectedPaths.length === 0}
                    actions={rowActions}
                  />
                ))}
              </ul>
              {hasMore && (
                <div ref={sentinelRef} className="load-more" aria-hidden="true">
                  <span className="load-more-spinner" />
                </div>
              )}
            </div>
          )}
        </section>
      </main>
    )
  }

  const movingDirs = (paths: string[]) =>
    paths.filter((p) => {
      // a folder being moved: its own listing is cached, or it is a known sub-folder
      if (listings.has(p)) return true
      const parent = listings.get(dirOf(p))
      return parent?.folders.some((f) => f.path === p) ?? false
    })

  const activeNav = desktopMode
    ? 'desktop'
    : !selectedFolder && !cloudMode
      ? askMode
        ? 'ask'
        : view
      : null

  return (
    <div className={`home${desktopMode ? ' desktop-mode' : ''}`}>
      {desktopMode && (
        <div className="home-desktop-backdrop" aria-hidden="true" ref={setBackdropHost} />
      )}
      {desktopMode && <div className="sidebar-glass" aria-hidden="true" />}
      <aside className="sidebar">
        <div className="sidebar-logo">
          <img className="logo-lockup" src={logoLockup} alt="GenOffice" />
        </div>
        <nav className="sidebar-nav">
          <button
            className={`nav-item${desktopMode ? ' active' : ''}`}
            aria-current={desktopMode ? 'page' : undefined}
            onClick={showDesktop}
          >
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
              <rect
                x="1.8"
                y="1.8"
                width="5.4"
                height="5.4"
                rx="1.4"
                stroke="currentColor"
                strokeWidth="1.3"
              />
              <rect
                x="8.8"
                y="1.8"
                width="5.4"
                height="5.4"
                rx="1.4"
                stroke="currentColor"
                strokeWidth="1.3"
              />
              <rect
                x="1.8"
                y="8.8"
                width="5.4"
                height="5.4"
                rx="1.4"
                stroke="currentColor"
                strokeWidth="1.3"
              />
              <rect
                x="8.8"
                y="8.8"
                width="5.4"
                height="5.4"
                rx="1.4"
                stroke="currentColor"
                strokeWidth="1.3"
              />
            </svg>
            <span className="nav-label">{DESKTOP_STRINGS[lang].navDesktop}</span>
          </button>
          <button
            className={`nav-item${activeNav === 'recent' ? ' active' : ''}`}
            aria-current={activeNav === 'recent' ? 'page' : undefined}
            onClick={() => changeView('recent')}
          >
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
              <circle cx="8" cy="8" r="6.2" stroke="currentColor" strokeWidth="1.3" />
              <path
                d="M8 4.8V8l2.2 1.6"
                stroke="currentColor"
                strokeWidth="1.3"
                strokeLinecap="round"
              />
            </svg>
            <span className="nav-label">{t('navRecent')}</span>
            <span className="nav-count">{navCounts.recent}</span>
          </button>
          <button
            className={`nav-item${activeNav === 'starred' ? ' active' : ''}`}
            aria-current={activeNav === 'starred' ? 'page' : undefined}
            onClick={() => changeView('starred')}
          >
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
              <path
                d="M8 1.9l1.9 3.85 4.25.62-3.07 3 .72 4.23L8 11.6l-3.8 2 .72-4.23-3.07-3 4.25-.62z"
                stroke="currentColor"
                strokeWidth="1.3"
                strokeLinejoin="round"
              />
            </svg>
            <span className="nav-label">{t('navStarred')}</span>
            <span className="nav-count">{navCounts.starred}</span>
          </button>
          <button
            className={`nav-item${activeNav === 'ask' ? ' active' : ''}`}
            aria-current={activeNav === 'ask' ? 'page' : undefined}
            onClick={openAskView}
          >
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
              <path
                d="M14 10.5c0 .9-.7 1.6-1.6 1.6H5.4L2 14.5V3.6C2 2.7 2.7 2 3.6 2h8.8c.9 0 1.6.7 1.6 1.6z"
                stroke="currentColor"
                strokeWidth="1.3"
                strokeLinejoin="round"
              />
              <path
                d="M5 6.2h6M5 8.6h4"
                stroke="currentColor"
                strokeWidth="1.3"
                strokeLinecap="round"
              />
            </svg>
            <span className="nav-label">{t('navAsk')}</span>
          </button>
          {loggedIn && (
            <button
              className={`nav-item${cloudMode && !selectedFolder ? ' active' : ''}`}
              onClick={() => {
                dispatchView({ type: 'show-cloud' })
                writeLaunchView('cloud')
                setSelected(new Set())
                setRowMenu(null)
              }}
            >
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                <path
                  d="M8 1.8l1.55 4.65L14.2 8l-4.65 1.55L8 14.2 6.45 9.55 1.8 8l4.65-1.55z"
                  stroke="currentColor"
                  strokeWidth="1.3"
                  strokeLinejoin="round"
                />
              </svg>
              <span className="nav-label">{t('navCloud')}</span>
              <svg
                className="nav-external"
                width="13"
                height="13"
                viewBox="0 0 16 16"
                fill="none"
                aria-hidden="true"
              >
                <path
                  d="M6.5 3.5H4a1.5 1.5 0 0 0-1.5 1.5v7A1.5 1.5 0 0 0 4 13.5h7A1.5 1.5 0 0 0 12.5 12V9.5M9.5 2.5h4v4M13 3l-5.5 5.5"
                  stroke="currentColor"
                  strokeWidth="1.3"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            </button>
          )}
        </nav>
        <div className="sidebar-divider" />
        {renderFolderPanel()}
        <TodoWidget lang={lang} />
        <AccountEntry
          onStatusChange={handleAccountStatus}
          onFileSearchSettingsChange={() => setRerankSettingsTick((n) => n + 1)}
          openRequest={settingsRequest}
        />
      </aside>
      {desktopMode ? (
        <DesktopView
          lang={lang}
          roots={roots}
          onOpenAsk={openAskView}
          onOpenTemplates={openTemplates}
          onOpenSettings={openSettings}
          onOpenFolder={openFolderView}
          backdropHost={backdropHost}
        />
      ) : selectedFolder && rootOf(selectedFolder, roots)?.readable ? (
        renderFolderContent()
      ) : askMode ? (
        <KbAskPanel
          onOpenPath={(path) => void window.aiOffice.openPath(path)}
          onGoToRecents={() => changeView('recent')}
          onOpenModelSettings={() => setSettingsRequest({ section: 'aiModel' })}
          pendingQuery={pendingAskQuery}
          onPendingQueryConsumed={() => setPendingAskQuery(null)}
        />
      ) : cloudMode ? (
        <CloudProjectsView />
      ) : (
        renderGlobalContent()
      )}
      {confirmDelete && (
        <div className="modal-overlay" onClick={() => setConfirmDelete(null)}>
          <div
            className="modal"
            role="dialog"
            aria-modal="true"
            aria-label={t('deleteModalTitle')}
            onClick={(event) => event.stopPropagation()}
          >
            <h3>{t('deleteModalTitle')}</h3>
            <p>
              {confirmDelete.length === 1
                ? t('deleteConfirmOne', { name: fileName(confirmDelete[0]) })
                : t('deleteConfirmMany', { n: confirmDelete.length })}
            </p>
            {confirmDelete.length > 1 && (
              <ul className="modal-file-list">
                {confirmDelete.slice(0, 6).map((p) => (
                  <li key={p}>{fileName(p)}</li>
                ))}
                {confirmDelete.length > 6 && (
                  <li>{t('deleteMoreCount', { n: confirmDelete.length })}</li>
                )}
              </ul>
            )}
            <div className="modal-buttons">
              <button
                className="btn btn-secondary"
                autoFocus
                onClick={() => setConfirmDelete(null)}
              >
                {t('cancel')}
              </button>
              <button className="btn btn-danger" onClick={confirmDeleteNow}>
                {t('delete')}
              </button>
            </div>
          </div>
        </div>
      )}

      {confirmDeleteFolder && (
        <div className="modal-overlay" onClick={() => setConfirmDeleteFolder(null)}>
          <div
            className="modal"
            role="dialog"
            aria-modal="true"
            aria-label={t('deleteFolderTitle')}
            onClick={(event) => event.stopPropagation()}
          >
            <h3>{t('deleteFolderTitle')}</h3>
            <p>{t('deleteFolderConfirm', { name: fileName(confirmDeleteFolder) })}</p>
            <div className="modal-buttons">
              <button
                className="btn btn-secondary"
                autoFocus
                onClick={() => setConfirmDeleteFolder(null)}
              >
                {t('cancel')}
              </button>
              <button className="btn btn-danger" onClick={() => void confirmDeleteFolderNow()}>
                {t('delete')}
              </button>
            </div>
          </div>
        </div>
      )}

      {confirmMissing && (
        <div className="modal-overlay" onClick={() => setConfirmMissing(null)}>
          <div
            className="modal"
            role="dialog"
            aria-modal="true"
            aria-label={t('missingFileTitle')}
            onClick={(event) => event.stopPropagation()}
          >
            <h3>{t('missingFileTitle')}</h3>
            <p>{t('missingFileBody', { name: confirmMissing.name })}</p>
            <div className="modal-buttons">
              <button
                className="btn btn-secondary"
                autoFocus
                onClick={() => setConfirmMissing(null)}
              >
                {t('cancel')}
              </button>
              <button
                className="btn btn-danger"
                onClick={() => {
                  // main drops the star of an unavailable entry with the row
                  removeRecent([confirmMissing.path])
                  setConfirmMissing(null)
                }}
              >
                {t('removeFromList')}
              </button>
            </div>
          </div>
        </div>
      )}

      {movePicker && canMove && (
        <FolderPicker
          roots={roots}
          currentDirs={new Set(movePicker.map(dirOf))}
          movingDirs={movingDirs(movePicker)}
          count={movePicker.length}
          onCancel={() => setMovePicker(null)}
          onPick={(dir) => void doMove(movePicker, dir, 'ask')}
        />
      )}

      {groupPicker && (
        <GroupPicker
          groups={starredGroups}
          count={groupPicker.length}
          onCancel={() => setGroupPicker(null)}
          onPick={(group) => assignGroup(groupPicker, group)}
          onCreate={(name) => assignGroup(groupPicker, name)}
          onRemove={() => assignGroup(groupPicker, null)}
        />
      )}

      {conflict && (
        <ConflictPrompt
          names={conflict.paths.map(fileName)}
          onChoose={(policy) => {
            if (policy === 'skip') setConflict(null)
            else void doMove(conflict.paths, conflict.targetDir, policy)
          }}
        />
      )}

      <DropToOpenOverlay />
      {templatesOpen && <TemplateGallery lang={lang} onClose={() => setTemplatesOpen(false)} />}
      {paletteOpen && (
        <CommandPalette
          lang={lang}
          onClose={() => setPaletteOpen(false)}
          onAskAi={submitAskQuery}
        />
      )}
    </div>
  )
}
