import { useCallback, useEffect, useState } from 'react'
import type { FolderListing, FolderRoot } from '../../shared/home-api'
import { Chevron, FileBadge, FolderIcon, useFolderListings } from './folder-browse'
import { useI18n } from './locale'
import './files-pane.css'

/**
 * The in-editor files browser (issue #542): a shell-rendered tab at
 * ?mode=files that reuses the home folder APIs (folderRoots / listFolder /
 * openPath / watch). Browse the folder roots, click a file to open it in an
 * editor tab — no round trip through the Home screen. Read-only on purpose
 * for this first cut: row menus, drag-drop and file creation stay in Home.
 */

const PANE_STATE_KEY = 'files.pane'

interface PaneState {
  expanded: string[]
  selected: string | null
}

function readPaneState(): PaneState {
  try {
    const raw = JSON.parse(localStorage.getItem(PANE_STATE_KEY) ?? 'null') as PaneState | null
    if (raw && Array.isArray(raw.expanded)) {
      return {
        expanded: raw.expanded.filter((p): p is string => typeof p === 'string'),
        selected: typeof raw.selected === 'string' ? raw.selected : null,
      }
    }
  } catch {
    // corrupt or absent: start collapsed
  }
  return { expanded: [], selected: null }
}

function writePaneState(state: PaneState): void {
  try {
    localStorage.setItem(PANE_STATE_KEY, JSON.stringify(state))
  } catch {
    // quota / private mode: the pane just forgets its layout
  }
}

function basename(path: string): string {
  const idx = path.lastIndexOf('/')
  return idx >= 0 ? path.slice(idx + 1) : path
}

export function FilesScreen() {
  const { t } = useI18n()
  const [roots, setRoots] = useState<FolderRoot[]>([])
  const [expanded, setExpanded] = useState<ReadonlySet<string>>(
    () => new Set(readPaneState().expanded),
  )
  const [selected, setSelected] = useState<string | null>(() => readPaneState().selected)
  const { listings, load, invalidate, tracked } = useFolderListings()

  const refreshRoots = useCallback(() => {
    void window.aiOffice.folderRoots().then(setRoots)
  }, [])

  useEffect(() => {
    refreshRoots()
  }, [refreshRoots])

  // keep the pane live: the main process watches every listed directory and
  // broadcasts the changed dirs (same stream Home listens on)
  useEffect(() => {
    return window.aiOffice.onFolderChanged((dirs) => {
      invalidate(dirs.filter((d) => tracked(d)))
    })
  }, [invalidate, tracked])

  // load the selected dir and every expanded one
  useEffect(() => {
    for (const dir of expanded) load(dir)
    if (selected) load(selected)
  }, [expanded, selected, load])

  // first usable root becomes the selection when nothing is chosen yet
  useEffect(() => {
    if (selected) return
    const first = roots.find((r) => r.usable) ?? roots[0]
    if (first) {
      setSelected(first.path)
      setExpanded((prev) => new Set(prev).add(first.path))
    }
  }, [roots, selected])

  useEffect(() => {
    writePaneState({ expanded: [...expanded], selected })
  }, [expanded, selected])

  const addRoot = useCallback(() => {
    void window.aiOffice.addFolderRoot().then((added) => {
      refreshRoots()
      if (added) {
        setSelected(added.path)
        setExpanded((prev) => new Set(prev).add(added.path))
      }
    })
  }, [refreshRoots])

  const toggleDir = useCallback((dir: string) => {
    setExpanded((prev) => {
      const next = new Set(prev)
      if (next.has(dir)) next.delete(dir)
      else next.add(dir)
      return next
    })
    setSelected(dir)
  }, [])

  return (
    <div className="files-screen">
      <header className="files-head">
        <span className="files-title">{t('navFiles')}</span>
        <button className="files-add" onClick={addRoot}>
          {t('addFolderRoot')}
        </button>
      </header>
      {roots.length === 0 ? (
        <div className="files-empty files-empty-root">
          <p>{t('filesNoRoots')}</p>
          <button className="files-add" onClick={addRoot}>
            {t('addFolderRoot')}
          </button>
        </div>
      ) : (
        <div className="files-body">
          <nav className="files-tree" aria-label={t('navFiles')}>
            {roots.map((root) => (
              <TreeRow
                key={root.path}
                dir={root.path}
                name={root.name}
                depth={0}
                hasChildren
                expanded={expanded}
                listings={listings}
                selected={selected}
                onToggle={toggleDir}
              />
            ))}
          </nav>
          <section className="files-detail">
            {selected ? (
              <DetailList dir={selected} listing={listings.get(selected)} onOpenDir={toggleDir} />
            ) : null}
          </section>
        </div>
      )}
    </div>
  )
}

/** one folder row in the tree; children render recursively from the listing
    cache. `hasChildren` three-states the chevron: false hides it (a folder
    listing already told us it is empty), undefined keeps it (not loaded yet). */
function TreeRow(props: {
  dir: string
  name: string
  depth: number
  hasChildren?: boolean
  expanded: ReadonlySet<string>
  listings: ReadonlyMap<string, FolderListing>
  selected: string | null
  onToggle: (dir: string) => void
}) {
  const { dir, name, depth, hasChildren, expanded, listings, selected, onToggle } = props
  const listing = listings.get(dir)
  const isOpen = expanded.has(dir)
  const children = listing?.folders ?? []
  const showChevron = hasChildren !== false
  return (
    <>
      <button
        className={`files-row is-folder${selected === dir ? ' is-selected' : ''}`}
        style={{ paddingLeft: 10 + depth * 14 }}
        aria-expanded={hasChildren === false ? undefined : isOpen}
        onClick={() => onToggle(dir)}
      >
        <span className="files-row-chevron">{showChevron ? <Chevron open={isOpen} /> : null}</span>
        <FolderIcon size={15} open={isOpen} />
        <span className="files-row-name">{name}</span>
      </button>
      {isOpen
        ? children.map((child) => (
            <TreeRow
              key={child.path}
              dir={child.path}
              name={child.name}
              depth={depth + 1}
              hasChildren={child.hasSubfolders}
              expanded={expanded}
              listings={listings}
              selected={selected}
              onToggle={onToggle}
            />
          ))
        : null}
    </>
  )
}

/** the listing of the selected directory, rendered from the shared cache */
function DetailList(props: {
  dir: string
  listing: FolderListing | undefined
  onOpenDir: (dir: string) => void
}) {
  const { t } = useI18n()
  const listing = props.listing
  if (!listing) return null
  return (
    <>
      <div className="files-list-head">{basename(props.dir)}</div>
      {listing.files.length === 0 && listing.folders.length === 0 ? (
        <div className="files-empty">{t('filesEmptyDir')}</div>
      ) : (
        <div className="files-list">
          {listing.folders.map((folder) => (
            <button
              key={folder.path}
              className="files-row is-folder"
              onClick={() => props.onOpenDir(folder.path)}
            >
              <span className="files-row-chevron" />
              <FolderIcon size={15} />
              <span className="files-row-name">{folder.name}</span>
            </button>
          ))}
          {listing.files.map((file) => (
            <button
              key={file.path}
              className="files-row is-file"
              onClick={() => void window.aiOffice.openPath(file.path)}
              disabled={file.missing}
            >
              <span className="files-row-icon">
                <FileBadge ext={file.ext} size={15} />
              </span>
              <span className="files-row-name">{file.name}</span>
              {file.starred ? <span className="files-row-star">★</span> : null}
            </button>
          ))}
        </div>
      )}
    </>
  )
}
