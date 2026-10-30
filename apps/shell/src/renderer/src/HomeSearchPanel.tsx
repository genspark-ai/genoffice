import type { ReactElement, Ref } from 'react'
import type {
  FileSearchHit,
  FileSearchPage,
  FileSearchRerank,
  FolderRoot,
} from '../../shared/home-api'
import { FileBadge } from './file-badge'
import { useI18n } from './locale'
import { EmptyHint, dirOf, formatModified, highlightText, locationLabel } from './home-shared'

/**
 * The file-search surface both content views (plain and folder) share: the
 * box, the settings shortcut, the results list and the heading. The query
 * state, the debounce and the rerank timer stay in Home; this file owns the
 * markup those states render.
 */

export interface FileSearchBoxProps {
  searchQuery: string
  onSearchQueryChange: (value: string) => void
  searchActive: boolean
  clearSearch: () => void
  searchInputRef: Ref<HTMLInputElement>
  /** IME composition flag, mirrored by the search effect's debounce in Home */
  composingRef: { current: boolean }
  onOpenSearchSettings: () => void
}

export function FileSearchBox({
  searchQuery,
  onSearchQueryChange,
  searchActive,
  clearSearch,
  searchInputRef,
  composingRef,
  onOpenSearchSettings,
}: FileSearchBoxProps): ReactElement {
  const { t } = useI18n()
  return (
    <div className="file-search-group">
      <div className={`file-search${searchActive ? ' active' : ''}`}>
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" aria-hidden="true">
          <circle cx="11" cy="11" r="7" stroke="currentColor" strokeWidth="2" />
          <path d="M20 20l-3.5-3.5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
        </svg>
        <input
          ref={searchInputRef}
          type="search"
          value={searchQuery}
          placeholder={t('searchFilesPlaceholder')}
          aria-label={t('searchFilesPlaceholder')}
          spellCheck={false}
          onChange={(e) => onSearchQueryChange(e.target.value)}
          onCompositionStart={() => {
            composingRef.current = true
          }}
          onCompositionEnd={(e) => {
            composingRef.current = false
            onSearchQueryChange(e.currentTarget.value)
          }}
          onKeyDown={(e) => {
            if (e.key === 'Escape' && searchQuery) {
              e.stopPropagation()
              clearSearch()
            }
          }}
        />
        {searchQuery && (
          <button className="file-search-clear" aria-label={t('searchClear')} onClick={clearSearch}>
            <svg width="12" height="12" viewBox="0 0 12 12" fill="none" aria-hidden="true">
              <path
                d="M2 2l8 8M10 2l-8 8"
                stroke="currentColor"
                strokeWidth="1.6"
                strokeLinecap="round"
              />
            </svg>
          </button>
        )}
      </div>
      <button
        className="file-search-settings"
        title={t('searchJevSettings')}
        aria-label={t('searchJevSettings')}
        onClick={onOpenSearchSettings}
      >
        <svg
          width="16"
          height="16"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
          strokeLinecap="round"
          strokeLinejoin="round"
          aria-hidden="true"
        >
          <path d="M12 20V4M12 4l-4 4M12 4l4 4" />
          <path d="M3 12h6M3 16h5M3 20h4" />
        </svg>
      </button>
    </div>
  )
}

export interface FileSearchResultsProps {
  searchPage: FileSearchPage | null
  /** the trimmed query, for the no-results copy */
  q: string
  rerankApplied: FileSearchRerank | null
  roots: readonly FolderRoot[]
}

export function FileSearchResults({
  searchPage,
  q,
  rerankApplied,
  roots,
}: FileSearchResultsProps): ReactElement {
  const i18n = useI18n()
  const { t } = i18n
  const page = searchPage
  const busy = !!page && (page.index.pending > 0 || page.index.scanning)

  /** judged hits in Jev's order, then the rest in local order */
  const orderedSearchHits = (hits: readonly FileSearchHit[]): FileSearchHit[] => {
    if (!rerankApplied) return [...hits]
    const rank = new Map(rerankApplied.order.map((path, i) => [path, i]))
    return [...hits].sort((a, b) => (rank.get(a.path) ?? Infinity) - (rank.get(b.path) ?? Infinity))
  }

  const renderSearchRow = (hit: FileSearchHit): ReactElement => (
    <li
      key={hit.path}
      className="search-row"
      role="button"
      tabIndex={0}
      onClick={() => void window.aiOffice.openPath(hit.path)}
      onKeyDown={(e) => {
        if (e.key === 'Enter') void window.aiOffice.openPath(hit.path)
      }}
    >
      <span className="recent-icon">
        <FileBadge ext={hit.ext} size={24} />
      </span>
      <div className="search-main">
        <div className="search-head">
          <span className="search-name">{highlightText(hit.name, hit.needles)}</span>
          <span className="search-path" title={dirOf(hit.path)}>
            {highlightText(locationLabel(hit.path, roots), hit.needles)}
          </span>
        </div>
        {hit.snippet && (
          <p className="search-snippet">
            {hit.snippet.map((part, i) =>
              part.hit ? (
                <mark key={i} className="search-hit">
                  {part.text}
                </mark>
              ) : (
                <span key={i}>{part.text}</span>
              ),
            )}
          </p>
        )}
      </div>
      <span className="search-time">{formatModified(hit.mtimeMs, i18n)}</span>
    </li>
  )

  return (
    <div className="search-results" aria-live="polite">
      {busy && (
        <div className="search-status">
          <span className="load-more-spinner" />
          {t('searchIndexing', { n: page.index.pending })}
        </div>
      )}
      {page && page.hits.length === 0 && !busy ? (
        <EmptyHint hint={t('searchNoResults', { q })} />
      ) : (
        page && <ul className="search-list">{orderedSearchHits(page.hits).map(renderSearchRow)}</ul>
      )}
    </div>
  )
}

export function FileSearchHeading({
  total,
  rerankApplied,
}: {
  total: number
  rerankApplied: FileSearchRerank | null
}): ReactElement {
  const { t } = useI18n()
  return (
    <div className="recents-heading">
      {rerankApplied && (
        <span className="search-rerank-badge" title={t('searchRerankedBy')}>
          Jev
        </span>
      )}
      <span className="file-count">
        {t(total === 1 ? 'searchResultCountOne' : 'searchResultCount', { n: total })}
      </span>
    </div>
  )
}
