import { useEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import type { CloudProjectKind, CloudProjectsSnapshot } from '../../shared/home-api'
import { useDismissablePopover } from '@genoffice/ui'
import { FileBadge } from './file-badge'
import { useI18n } from './locale'
import type { StringKey } from './locale'
import { SortCheck, formatModified } from './home-shared'

// ── Cloud (Genspark web) projects view ──────────────────

/** kind filter segments; labels shared with the recents type filter */
const CLOUD_FILTERS = [
  { key: 'all', label: 'filterAll' },
  { key: 'docs', label: 'filterDocs' },
  { key: 'sheets', label: 'filterSheets' },
  { key: 'slides', label: 'filterSlides' },
] as const satisfies readonly { key: 'all' | CloudProjectKind; label: StringKey }[]

/** module kind → file icon extension */
const CLOUD_KIND_EXT: Record<string, string> = { docs: 'docx', sheets: 'xlsx', slides: 'pptx' }

/** rows revealed per "load more" step; purely client-side over the local snapshot */
const CLOUD_REVEAL_STEP = 100

export function CloudProjectsView() {
  const i18n = useI18n()
  const { t } = i18n
  const [snapshot, setSnapshot] = useState<CloudProjectsSnapshot | null>(null)
  const [loading, setLoading] = useState(true)
  const [syncing, setSyncing] = useState(false)
  const [loginWaiting, setLoginWaiting] = useState(false)
  const [kind, setKind] = useState<'all' | CloudProjectKind>('all')
  const [query, setQuery] = useState('')
  const [sort, setSort] = useState<'recent' | 'oldest'>('recent')
  const [sortMenuOpen, setSortMenuOpen] = useState(false)
  const [revealed, setRevealed] = useState(CLOUD_REVEAL_STEP)
  const sortRef = useRef<HTMLDivElement>(null)

  // the local store paints instantly; a background sync replaces it when done.
  // a failed sync keeps whatever is shown; with nothing shown the
  // !snapshot && !loading branch below renders the retry state
  const startSync = () => {
    setSyncing(true)
    void window.aiOffice.cloudProjectsSync?.().then((synced) => {
      setSyncing(false)
      setLoading(false)
      if (synced) setSnapshot(synced)
    })
  }
  const startSyncRef = useRef(startSync)
  startSyncRef.current = startSync

  useEffect(() => {
    let cancelled = false
    void window.aiOffice.cloudProjectsCached?.().then((stored) => {
      if (cancelled || !stored) return
      setSnapshot((prev) => prev ?? stored)
      setLoading(false)
    })
    startSyncRef.current()
    return () => {
      cancelled = true
    }
  }, [])

  // the sign-in button reuses the account login flow; sync once it lands
  useEffect(() => {
    const off = window.aiOffice.onAccountLogin?.((ev) => {
      if (ev.phase === 'success') {
        setLoginWaiting(false)
        startSyncRef.current()
      } else if (ev.phase === 'error') {
        setLoginWaiting(false)
      }
    })
    return off
  }, [])

  // unified dismissal: outside press, window blur, chrome press (tab strip / window drag)
  useDismissablePopover(sortMenuOpen, () => setSortMenuOpen(false), {
    inside: () => [sortRef.current],
  })

  const startLogin = () => {
    setLoginWaiting(true)
    void window.aiOffice.accountLogin?.().then((ok) => {
      if (!ok) setLoginWaiting(false)
    })
  }

  const changeKind = (k: 'all' | CloudProjectKind) => {
    if (k === kind) return
    setKind(k)
    setRevealed(CLOUD_REVEAL_STEP)
  }

  const openProject = (projectUrl: string) => {
    void window.aiOffice.openCloudProject?.(projectUrl)
  }

  // filter / search / sort are all local over the snapshot — no requests
  const q = query.trim().toLowerCase()
  let list = snapshot?.projects.filter((proj) => kind === 'all' || proj.kind === kind) ?? []
  if (q) list = list.filter((proj) => proj.title.toLowerCase().includes(q))
  if (sort === 'oldest') list = [...list].reverse()
  const visible = list.slice(0, revealed)

  const renderRows = () => {
    const items: ReactElement[] = []
    for (const proj of visible) {
      items.push(
        <li key={proj.projectId}>
          <button
            className="cloud-row"
            data-tip={t('cloudOpenInBrowser')}
            data-tip-anchor=".cloud-row-external"
            data-tip-place="right"
            onClick={() => openProject(proj.projectUrl)}
          >
            <FileBadge ext={CLOUD_KIND_EXT[proj.kind] ?? ''} size={24} />
            <span className="cloud-row-main">
              <span className="cloud-row-title">{proj.title || t('untitled')}</span>
              <svg
                className="cloud-row-external"
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
            </span>
            <span className="cloud-row-time">
              {proj.ctimeMs ? formatModified(proj.ctimeMs, i18n) : ''}
            </span>
          </button>
        </li>,
      )
    }
    return items
  }

  const renderBody = () => {
    if (snapshot && !snapshot.available) {
      return (
        <p className="empty proj-empty">
          <span className="empty-hint">{t('cloudLoginHint')}</span>
          <button className="btn btn-secondary" disabled={loginWaiting} onClick={startLogin}>
            {loginWaiting ? t('waitingShort') : t('loginGenspark')}
          </button>
        </p>
      )
    }
    if (!snapshot) {
      if (loading || syncing) {
        return (
          <div className="load-more" aria-hidden="true">
            <span className="load-more-spinner" />
          </div>
        )
      }
      return (
        <p className="empty proj-empty">
          <span className="empty-hint">{t('cloudError')}</span>
          <button className="btn btn-secondary" onClick={() => startSync()}>
            {t('cloudRetry')}
          </button>
        </p>
      )
    }
    if (list.length === 0) {
      return (
        <p className="empty proj-empty">
          <span className="empty-hint">
            {t(q ? 'cloudNoResults' : kind === 'all' ? 'cloudEmpty' : 'emptyFiltered')}
          </span>
        </p>
      )
    }
    return (
      <div className="cloud-scroll">
        <div className="cloud-table">
          <div className="cloud-columns">
            <span className="col-name">{t('colName')}</span>
            <div className="cloud-col-sort" ref={sortRef}>
              <button
                className="cloud-col-sort-btn"
                aria-haspopup="menu"
                aria-expanded={sortMenuOpen}
                onClick={() => setSortMenuOpen((o) => !o)}
              >
                {t('colModified')}
                <svg
                  width="12"
                  height="12"
                  viewBox="0 0 16 16"
                  fill="none"
                  aria-hidden="true"
                  style={sort === 'oldest' ? { transform: 'rotate(180deg)' } : undefined}
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
              {sortMenuOpen && (
                <div className="cloud-sort-menu" role="menu">
                  {(['recent', 'oldest'] as const).map((key) => (
                    <button
                      key={key}
                      className={sort === key ? 'active' : ''}
                      role="menuitemradio"
                      aria-checked={sort === key}
                      onClick={() => {
                        setSort(key)
                        setSortMenuOpen(false)
                        setRevealed(CLOUD_REVEAL_STEP)
                      }}
                    >
                      <SortCheck visible={sort === key} />
                      {t(key === 'recent' ? 'cloudSortRecent' : 'cloudSortOldest')}
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>
          <ul className="cloud-list">{renderRows()}</ul>
        </div>
        {list.length > revealed && (
          <div className="load-more">
            <button
              className="btn btn-secondary"
              onClick={() => setRevealed((n) => n + CLOUD_REVEAL_STEP)}
            >
              {t('cloudLoadMore')}
            </button>
          </div>
        )}
      </div>
    )
  }

  return (
    <main className="content">
      <section className="cloud-projects" aria-label={t('navCloud')}>
        <header className="cloud-hero">
          <div className="cloud-hero-top">
            <h1 className="cloud-title">{t('navCloud')}</h1>
          </div>
          <p className="cloud-subtitle">{t('cloudSubtitle')}</p>
          {snapshot?.available && (
            <div className="cloud-controls">
              <div className="cloud-seg" role="tablist" aria-label={t('filterAria')}>
                {CLOUD_FILTERS.map((f) => (
                  <button
                    key={f.key}
                    className={kind === f.key ? ' active' : ''}
                    role="tab"
                    aria-selected={kind === f.key}
                    onClick={() => changeKind(f.key)}
                  >
                    {t(f.label)}
                  </button>
                ))}
              </div>
              <button
                className={`cloud-refresh-btn${syncing ? ' syncing' : ''}`}
                data-tip={t('cloudRefresh')}
                aria-label={t('cloudRefresh')}
                disabled={syncing}
                onClick={() => startSync()}
              >
                <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                  <path
                    d="M13.6 8a5.6 5.6 0 1 1-1.64-3.96M13.6 2.4v3.2h-3.2"
                    stroke="currentColor"
                    strokeWidth="1.4"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                </svg>
              </button>
              <div className="cloud-search">
                <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                  <circle cx="7" cy="7" r="4.6" stroke="currentColor" strokeWidth="1.4" />
                  <path
                    d="M10.5 10.5L14 14"
                    stroke="currentColor"
                    strokeWidth="1.4"
                    strokeLinecap="round"
                  />
                </svg>
                <input
                  value={query}
                  placeholder={t('cloudSearchPlaceholder', { n: snapshot.projects.length })}
                  onChange={(e) => {
                    setQuery(e.target.value)
                    setRevealed(CLOUD_REVEAL_STEP)
                  }}
                />
              </div>
            </div>
          )}
        </header>
        {renderBody()}
      </section>
    </main>
  )
}
