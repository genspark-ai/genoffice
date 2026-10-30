/**
 * Home's view-mode state machine: the six surfaces that own the content area
 * (desktop / recent / starred / ask / cloud / folder) were five hand-written
 * flags plus a folder path, and every switch function had to remember to clear
 * the other four by hand. This reducer makes the mutual exclusion a single
 * transition: a surface switch is one action, and the old surface's mode
 * state (`list`, `folder`) is either kept or dropped per action exactly the
 * way the hand-written switchers did — zero behavior change.
 */

// ── Launch view preference ──────────────────────────────
// The desktop is the default view; the last top-level nav the person chose
// takes over on the next launch. Stored through the same localStorage
// mechanism the folder tree's expansion state uses (Home has no other
// renderer-writable settings door, and no view pref existed before).

export type LaunchView = 'desktop' | 'recent' | 'starred' | 'ask' | 'cloud'
const VIEW_PREF_KEY = 'home.view'

export function readLaunchView(): LaunchView {
  try {
    const raw = localStorage.getItem(VIEW_PREF_KEY)
    if (raw === 'recent' || raw === 'starred' || raw === 'ask' || raw === 'cloud') return raw
  } catch {
    // corrupt or absent: the desktop is the default
  }
  return 'desktop'
}

export function writeLaunchView(view: LaunchView): void {
  try {
    localStorage.setItem(VIEW_PREF_KEY, view)
  } catch {
    // quota / private mode: the choice lasts this session only
  }
}

/** which surface owns the content area; exactly one at a time */
export type HomeSurface = 'desktop' | 'list' | 'folder' | 'ask' | 'cloud'

export interface HomeViewState {
  surface: HomeSurface
  /** the list surface's own mode; remembered across surface switches */
  list: 'recent' | 'starred'
  /** the folder the folder surface shows; null whenever another surface is up */
  folder: string | null
}

export type HomeViewAction =
  /** the desktop nav item and the pin-to-desktop broadcast */
  | { type: 'show-desktop' }
  /** the Recent / Starred nav items */
  | { type: 'show-list'; list: 'recent' | 'starred' }
  /** a tree or table row click: the folder takes over the content area */
  | { type: 'select-folder'; dir: string }
  /**
   * A non-switching folder update — the path rebases a rename/move leaves
   * behind, or a folder that vanished from disk falling back (null drops the
   * folder surface back to the list, like the old setSelectedFolder(null))
   */
  | { type: 'set-folder'; dir: string | null }
  /** the Ask AI nav item, the desktop's ask tile, deep links and the palette */
  | { type: 'show-ask' }
  /** the Cloud nav item */
  | { type: 'show-cloud' }
  /** logout only: the cloud view leaves, everything else stays untouched */
  | { type: 'exit-cloud' }

export function homeViewReducer(state: HomeViewState, action: HomeViewAction): HomeViewState {
  switch (action.type) {
    case 'show-desktop':
      return { ...state, surface: 'desktop', folder: null }
    case 'show-list':
      return { ...state, surface: 'list', list: action.list, folder: null }
    case 'select-folder':
      return { ...state, surface: 'folder', folder: action.dir }
    case 'set-folder':
      if (action.dir === null) {
        // only the folder surface falls back to the list; on any other surface
        // the folder path is already null and nothing changes
        return state.surface === 'folder' ? { ...state, surface: 'list', folder: null } : state
      }
      return { ...state, surface: 'folder', folder: action.dir }
    case 'show-ask':
      return { ...state, surface: 'ask', folder: null }
    case 'show-cloud':
      return { ...state, surface: 'cloud', folder: null }
    case 'exit-cloud':
      return state.surface === 'cloud' ? { ...state, surface: 'list' } : state
  }
}

export function initialHomeViewState(launch: LaunchView): HomeViewState {
  return {
    surface:
      launch === 'desktop'
        ? 'desktop'
        : launch === 'ask'
          ? 'ask'
          : launch === 'cloud'
            ? 'cloud'
            : 'list',
    list: launch === 'starred' ? 'starred' : 'recent',
    folder: null,
  }
}
