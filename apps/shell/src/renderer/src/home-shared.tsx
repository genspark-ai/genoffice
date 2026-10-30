import type { ReactElement } from 'react'
import type { FolderRoot, RecentEntry } from '../../shared/home-api'
import { markText } from '../../shared/text-marks'
import type { I18n, StringKey } from './locale'

/**
 * The helpers, icons and constants every Home surface shares — pulled out of
 * Home.tsx so the per-surface components (FolderPicker, CloudProjectsView,
 * the file tables) and Home itself import one place instead of duplicating.
 */

/* Formats the open-local card advertises. Too long for the card at any window
   width, so it ellipsizes and a hover ScreenTip carries the full list. Keep in
   sync with the main-process open-dialog filter (OPEN_DIALOG_EXTENSIONS). */
export const OPEN_LOCAL_EXTENSIONS =
  '.docx / .odt / .ott / .xlsx / .xlsm / .xls / .csv / .tsv / .ods / .ots / .pptx / .odp / .otp / .pdf / .md / .html'

/** drag payload of home file/folder rows (JSON array of absolute paths) */
export const DRAG_PATHS_MIME = 'application/x-genoffice-paths'
/** hovering a collapsed folder this long while dragging expands it */
export const DRAG_EXPAND_DELAY_MS = 600
/** expanded folders survive a home reload; the selection is per session */
export const TREE_STATE_KEY = 'home.folderTree'

/**
 * Width of the viewport that `position: fixed` elements anchor to. The layout
 * engine owns this coordinate system, so documentElement's client rect is the
 * value to compute menu offsets from; window.innerWidth can disagree on scaled
 * Linux desktops, which pushed row menus past the window's left edge (genoffice#1946).
 */
export function fixedViewportWidth(): number {
  return document.documentElement.clientWidth
}

export function FolderIcon({ size = 16, open = false }: { size?: number; open?: boolean }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" aria-hidden="true">
      <path
        d="M1.5 4A1.5 1.5 0 0 1 3 2.5h3.1c.44 0 .85.19 1.13.52L8.4 4.4H13A1.5 1.5 0 0 1 14.5 5.9v5.6A1.5 1.5 0 0 1 13 13H3a1.5 1.5 0 0 1-1.5-1.5V4z"
        stroke="currentColor"
        strokeWidth="1.2"
        strokeLinejoin="round"
        fill={open ? 'currentColor' : 'none'}
        fillOpacity={open ? 0.12 : 0}
      />
    </svg>
  )
}

export function Chevron({ open }: { open: boolean }) {
  return (
    <svg
      width="12"
      height="12"
      viewBox="0 0 12 12"
      aria-hidden="true"
      style={{ transform: open ? 'rotate(90deg)' : undefined, transition: 'transform 0.12s' }}
    >
      <path
        d="M4.5 2.5l4 3.5-4 3.5"
        stroke="currentColor"
        strokeWidth="1.3"
        strokeLinecap="round"
        fill="none"
      />
    </svg>
  )
}

export function MoreDots() {
  return (
    <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true">
      <circle cx="3.2" cy="8" r="1.35" fill="currentColor" />
      <circle cx="8" cy="8" r="1.35" fill="currentColor" />
      <circle cx="12.8" cy="8" r="1.35" fill="currentColor" />
    </svg>
  )
}

export function formatModified(mtimeMs: number, i18n: I18n): string {
  const date = new Date(mtimeMs)
  const now = new Date()
  const startOfDay = (d: Date) => new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime()
  const days = Math.round((startOfDay(now) - startOfDay(date)) / 86400000)
  if (days <= 0) {
    return `${i18n.t('today')} · ${date.toLocaleTimeString(i18n.dateLocale, { hour: '2-digit', minute: '2-digit' })}`
  }
  if (days === 1) return i18n.t('yesterday')
  return date.toLocaleDateString(i18n.dateLocale, { month: 'short', day: 'numeric' })
}

export function formatSize(bytes: number): string {
  if (bytes >= 1048576) return `${(bytes / 1048576).toFixed(1)} MB`
  return `${Math.max(1, Math.round(bytes / 1024))} KB`
}

export function splitPath(path: string): string[] {
  return path.split(/[\\/]/).filter(Boolean)
}

export function parentDir(path: string): string {
  const parts = splitPath(path)
  return parts[parts.length - 2] ?? ''
}

export function dirOf(path: string): string {
  const cut = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  if (cut < 0) return path
  // keep the separator on a filesystem root ("/" or "C:\") instead of an empty or bare-drive string
  const dir = path.slice(0, cut)
  return dir === '' || /^[A-Za-z]:$/.test(dir) ? path.slice(0, cut + 1) : dir
}

export function isUnder(root: string, path: string): boolean {
  if (path === root) return true
  // a root that already ends with a separator ("/" or "C:\") must not get a second one
  if (/[\\/]$/.test(root)) return path.startsWith(root)
  return path.startsWith(root + '/') || path.startsWith(root + '\\')
}

export function fileName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path
}

export function baseName(entry: RecentEntry): string {
  return entry.ext ? entry.name.slice(0, -(entry.ext.length + 1)) : entry.name
}

/** the root that holds `path`; the deepest one when roots nest */
export function rootOf(path: string, roots: readonly FolderRoot[]): FolderRoot | null {
  let best: FolderRoot | null = null
  for (const root of roots) {
    if (isUnder(root.path, path) && (!best || root.path.length > best.path.length)) best = root
  }
  return best
}

/** "Clients / Contracts" for a file under a root; the parent folder name elsewhere */
export function locationLabel(path: string, roots: readonly FolderRoot[]): string {
  const dir = dirOf(path)
  const root = rootOf(dir, roots)
  if (root) {
    if (dir === root.path) return root.name
    return splitPath(dir.slice(root.path.length)).join(' / ')
  }
  return parentDir(path)
}

/** the folders between root and `dir` (inclusive) */
export function crumbsOf(root: FolderRoot, dir: string): Array<{ path: string; name: string }> {
  const crumbs = [{ path: root.path, name: root.name }]
  if (dir === root.path) return crumbs
  const sep = dir.includes('\\') && !dir.includes('/') ? '\\' : '/'
  let acc = root.path
  for (const part of splitPath(dir.slice(root.path.length))) {
    acc = `${acc}${sep}${part}`
    crumbs.push({ path: acc, name: part })
  }
  return crumbs
}

export const FILTERS: { key: string; label: StringKey }[] = [
  { key: 'all', label: 'filterAll' },
  { key: 'docx', label: 'filterDocs' },
  { key: 'xlsx', label: 'filterSheets' },
  { key: 'pptx', label: 'filterSlides' },
  { key: 'pdf', label: 'filterPdf' },
  { key: 'md', label: 'filterMd' },
  { key: 'html', label: 'filterHtml' },
]

/** sidebar filter keys that stand for a family of extensions (mirrors recent-files.ts) */
export const FILTER_FAMILY: Record<string, readonly string[]> = {
  docx: ['docx', 'doc', 'odt', 'ott'],
  xlsx: ['xlsx', 'xlsm', 'xls', 'csv', 'tsv', 'ods', 'ots'],
  pptx: ['pptx', 'ppt', 'odp', 'otp'],
  md: ['md', 'markdown', 'txt', 'json'],
  html: ['html', 'htm'],
}

export function matchesFilter(entry: RecentEntry, filter: string): boolean {
  if (filter === 'all') return true
  return (FILTER_FAMILY[filter] ?? [filter]).includes(entry.ext)
}

/** the empty-state card the list views and the search results share */
export function EmptyHint({ hint }: { hint: string }): ReactElement {
  return (
    <p className="empty proj-empty">
      <svg
        className="proj-empty-icon"
        width="48"
        height="48"
        viewBox="0 0 24 24"
        fill="none"
        aria-hidden="true"
      >
        <path
          d="M6.29297 3.75H14.1729C14.4927 3.75 14.7979 3.88392 15.0146 4.11914L18.5566 7.96387C18.7512 8.17512 18.8593 8.45208 18.8594 8.73926V19.1055C18.8593 19.7376 18.346 20.25 17.7139 20.25H6.29297C5.66091 20.2499 5.14855 19.7375 5.14844 19.1055V4.89453C5.14855 4.26247 5.66091 3.75011 6.29297 3.75Z"
          stroke="currentColor"
          strokeWidth="1.5"
        />
        <path
          d="M13.8984 4V7.11C13.8984 8.15382 14.7446 9 15.7884 9H18.8984"
          stroke="currentColor"
          strokeWidth="1.5"
        />
      </svg>
      <span className="empty-hint">{hint}</span>
    </p>
  )
}

/** wrap every matched fragment of a name or folder label in the shared search-hit mark */
export function highlightText(text: string, needles: readonly string[]): ReactElement[] | string {
  const marks = markText(text, needles)
  if (!marks.some((m) => m.hit)) return text
  return marks.map((m, i) =>
    m.hit ? (
      <mark key={i} className="search-hit">
        {m.text}
      </mark>
    ) : (
      <span key={i}>{m.text}</span>
    ),
  )
}

/** Check glyph marking the selected sort option; invisible on the others so labels stay aligned */
export function SortCheck({ visible }: { visible: boolean }): ReactElement {
  return (
    <svg
      className="cloud-sort-check"
      width="12"
      height="12"
      viewBox="0 0 16 16"
      fill="none"
      aria-hidden="true"
      style={visible ? undefined : { visibility: 'hidden' }}
    >
      <path
        d="M3 8.5L6.5 12L13 4.5"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

// ── Folder tree state (shared by the sidebar and the move picker) ──

export interface TreeState {
  expanded: string[]
  /** the root the layout was saved for; a layout for another root is not applied */
  root: string | null
}

export function readTreeState(): TreeState {
  try {
    const raw = JSON.parse(localStorage.getItem(TREE_STATE_KEY) ?? 'null') as TreeState | null
    if (raw && Array.isArray(raw.expanded)) {
      return {
        expanded: raw.expanded.filter((p): p is string => typeof p === 'string'),
        root: typeof raw.root === 'string' ? raw.root : null,
      }
    }
  } catch {
    // corrupt or absent: start collapsed
  }
  return { expanded: [], root: null }
}

export function writeTreeState(state: TreeState): void {
  try {
    localStorage.setItem(TREE_STATE_KEY, JSON.stringify(state))
  } catch {
    // quota / private mode: the tree just forgets its layout
  }
}
