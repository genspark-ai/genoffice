import { useCallback, useEffect, useRef, useState } from 'react'
import type { DragEvent, MouseEvent, ReactElement } from 'react'
import type { Lang } from '@genoffice/i18n'
import { useDismissablePopover } from '../popover-dismiss'
import { tFilesPane } from './strings'
import type {
  FileEntry,
  FilesPaneApi,
  FolderListing,
  FolderRoot,
  MoveConflictPolicy,
} from './types'

export interface FilesPaneProps {
  api: FilesPaneApi
  lang: Lang
  /** the document open in this editor; highlighted and revealed in the tree */
  currentPath: string | null
  onClose: () => void
}

/** drag payload shared with the shell home screen (JSON array of absolute paths) */
const DRAG_PATHS_MIME = 'application/x-genoffice-paths'
const DRAG_EXPAND_DELAY_MS = 600
const EXPANDED_KEY = 'genoffice.filesPane.expanded'

const BADGE: Record<string, { label: string; color: string }> = {
  docx: { label: 'W', color: '#2b63d9' },
  doc: { label: 'W', color: '#2b63d9' },
  xlsx: { label: 'X', color: '#1f8a4c' },
  xlsm: { label: 'X', color: '#1f8a4c' },
  xls: { label: 'X', color: '#1f8a4c' },
  csv: { label: 'X', color: '#1f8a4c' },
  pptx: { label: 'P', color: '#c9481f' },
  ppt: { label: 'P', color: '#c9481f' },
  pdf: { label: 'A', color: '#d6352d' },
  md: { label: 'M', color: '#7a4ddb' },
  markdown: { label: 'M', color: '#7a4ddb' },
  html: { label: '<>', color: '#1f8f8a' },
  htm: { label: '<>', color: '#1f8f8a' },
}

/* Badge colors are document-type brand colors (same in both themes), not chrome. */
function FileBadge({ ext }: { ext: string }) {
  const badge = BADGE[ext] ?? { label: ext ? ext.charAt(0).toUpperCase() : '?', color: '#98a2b3' }
  return (
    <svg className="fp-badge" width="16" height="16" viewBox="0 0 32 32" aria-hidden="true">
      <rect width="32" height="32" rx="7" fill={badge.color} />
      <text
        x="16"
        y="16.5"
        textAnchor="middle"
        dominantBaseline="central"
        fill="#fff"
        fontSize={badge.label.length > 1 ? 14 : 18}
        fontWeight="700"
        fontFamily="system-ui, -apple-system, 'Segoe UI', sans-serif"
      >
        {badge.label}
      </text>
    </svg>
  )
}

function FolderIcon({ open }: { open: boolean }) {
  return (
    <svg
      className="fp-folder-icon"
      width="16"
      height="16"
      viewBox="0 0 16 16"
      fill="none"
      aria-hidden="true"
    >
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

function dirOf(path: string): string {
  const cut = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  if (cut < 0) return path
  // keep the separator on a filesystem root ("/" or "C:\") instead of an empty or bare-drive string
  const dir = path.slice(0, cut)
  return dir === '' || /^[A-Za-z]:$/.test(dir) ? path.slice(0, cut + 1) : dir
}

function isUnder(root: string, path: string): boolean {
  if (path === root) return true
  // a root that already ends with a separator ("/" or "C:\") must not get a second one
  if (/[\\/]$/.test(root)) return path.startsWith(root)
  return path.startsWith(root + '/') || path.startsWith(root + '\\')
}

function nameOf(path: string): string {
  return path.split(/[\\/]/).pop() ?? path
}

function readExpanded(): string[] {
  try {
    const raw = JSON.parse(localStorage.getItem(EXPANDED_KEY) ?? '[]') as unknown
    return Array.isArray(raw) ? raw.filter((p): p is string => typeof p === 'string') : []
  } catch {
    return []
  }
}

interface MenuState {
  path: string
  kind: 'root' | 'folder' | 'file'
  name: string
  top: number
  left: number
}

/**
 * Left-docked folder tree over the default save folder, shared by every
 * editor. Files open in the shell (a new tab, or the tab already showing
 * them); folders and files can be created, renamed, moved by drag and
 * trashed right here, mirroring the home screen's Folders panel.
 */
export function FilesPane({
  api,
  lang,
  currentPath: hostPath,
  onClose,
}: FilesPaneProps): ReactElement {
  // A rename / move done here changes the open document's path before the
  // host learns about it (some hosts only receive the new name). Until the
  // host's path changes, the moved-to path stands in for it.
  const [alias, setAlias] = useState<{ from: string; to: string } | null>(null)
  const currentPath = alias && alias.from === hostPath ? alias.to : hostPath
  // once the host reports any new path the alias has served its purpose; it
  // must not resurface if the host later returns to the original name
  useEffect(() => {
    setAlias((prev) => (prev && prev.from !== hostPath ? null : prev))
  }, [hostPath])
  const followCurrent = (from: string, to: string) => {
    if (!currentPath) return
    if (currentPath === from) setAlias({ from: hostPath ?? from, to })
    else if (isUnder(from, currentPath)) {
      setAlias({ from: hostPath ?? currentPath, to: to + currentPath.slice(from.length) })
    }
  }
  const t = useCallback(
    (key: Parameters<typeof tFilesPane>[1], params?: Parameters<typeof tFilesPane>[2]) =>
      tFilesPane(lang, key, params),
    [lang],
  )
  const [root, setRoot] = useState<FolderRoot | null>(null)
  const [listings, setListings] = useState<ReadonlyMap<string, FolderListing>>(new Map())
  const [expanded, setExpanded] = useState<ReadonlySet<string>>(() => new Set(readExpanded()))
  const [menu, setMenu] = useState<MenuState | null>(null)
  const menuRef = useRef<HTMLDivElement>(null)
  const [renaming, setRenaming] = useState<{
    path: string
    value: string
    kind: 'folder' | 'file'
  } | null>(null)
  const [creatingIn, setCreatingIn] = useState<string | null>(null)
  const [newName, setNewName] = useState('')
  const [confirm, setConfirm] = useState<{ path: string; kind: 'folder' | 'file' } | null>(null)
  const [conflict, setConflict] = useState<{ paths: string[]; targetDir: string } | null>(null)
  const [dropTarget, setDropTarget] = useState<string | null>(null)
  const dragExpandTimer = useRef<number | null>(null)
  // dir → whether a reload was requested while its request was in flight (the
  // in-flight answer may predate the change, so it is fetched once more)
  const inflight = useRef(new Map<string, boolean>())

  const load = useCallback(
    (dir: string, force = false) => {
      if (inflight.current.has(dir)) {
        // only a real change re-queues; the expand effect re-asking for a
        // folder that was just invalidated rides on the request already out
        if (force) inflight.current.set(dir, true)
        return
      }
      inflight.current.set(dir, false)
      void api
        .listFolder(dir)
        .then((listing) => {
          setListings((prev) => {
            const next = new Map(prev)
            next.set(dir, listing)
            return next
          })
        })
        .finally(() => {
          const again = inflight.current.get(dir)
          inflight.current.delete(dir)
          if (again) load(dir, true)
        })
    },
    [api],
  )

  const invalidate = useCallback(
    (dirs: readonly string[]) => {
      setListings((prev) => {
        const next = new Map(prev)
        let changed = false
        for (const dir of dirs) if (next.delete(dir)) changed = true
        return changed ? next : prev
      })
      for (const dir of dirs) load(dir, true)
    },
    [load],
  )

  const refreshAll = useCallback(() => {
    void api.folderRoot().then(setRoot)
    invalidate([...listings.keys()])
  }, [api, invalidate, listings])
  const refreshRef = useRef(refreshAll)
  refreshRef.current = refreshAll

  useEffect(() => {
    void api.folderRoot().then(setRoot)
  }, [api])

  // folders shown or currently loading refresh; a change during the first
  // request queues the re-fetch through the in-flight flag
  useEffect(() => {
    return api.onFolderChanged((dirs) =>
      invalidate(dirs.filter((d) => listings.has(d) || inflight.current.has(d))),
    )
  }, [api, invalidate, listings])

  useEffect(() => {
    const onFocus = () => refreshRef.current()
    window.addEventListener('focus', onFocus)
    return () => window.removeEventListener('focus', onFocus)
  }, [])

  useEffect(() => {
    try {
      localStorage.setItem(EXPANDED_KEY, JSON.stringify([...expanded]))
    } catch {
      // private mode / quota: the tree just forgets its layout
    }
  }, [expanded])

  // root always open; every open folder is loaded; a folder that vanished
  // collapses (its parent's refresh removes the row)
  useEffect(() => {
    if (!root?.usable) return
    if (!expanded.has(root.path)) {
      setExpanded((prev) => new Set([...prev, root.path]))
      return
    }
    const gone = [...expanded].filter((dir) => dir !== root.path && listings.get(dir)?.missing)
    if (gone.length > 0) {
      setExpanded((prev) => new Set([...prev].filter((dir) => !gone.includes(dir))))
      return
    }
    for (const dir of expanded) if (!listings.has(dir)) load(dir)
  }, [root, expanded, listings, load])

  // reveal the current document: open every folder down to it
  useEffect(() => {
    if (!root?.usable || !currentPath || !isUnder(root.path, currentPath)) return
    const dirs: string[] = []
    let dir = dirOf(currentPath)
    while (isUnder(root.path, dir)) {
      dirs.push(dir)
      if (dir === root.path) break
      dir = dirOf(dir)
    }
    setExpanded((prev) => {
      if (dirs.every((d) => prev.has(d))) return prev
      return new Set([...prev, ...dirs])
    })
  }, [root, currentPath])

  useDismissablePopover(menu !== null, () => setMenu(null), { inside: () => [menuRef.current] })
  useEffect(() => {
    if (!menu && !confirm && !conflict) return
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setMenu(null)
        setConfirm(null)
        setConflict(null)
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [menu, confirm, conflict])
  useEffect(() => {
    if (!menu) return
    const close = () => setMenu(null)
    window.addEventListener('scroll', close, true)
    return () => window.removeEventListener('scroll', close, true)
  }, [menu])

  const toggle = (dir: string) =>
    setExpanded((prev) => {
      const next = new Set(prev)
      if (next.has(dir)) next.delete(dir)
      else next.add(dir)
      return next
    })

  const openMenuAt = (e: MouseEvent, state: Omit<MenuState, 'top' | 'left'>) => {
    e.preventDefault()
    e.stopPropagation()
    const maxLeft = window.innerWidth - 180
    setMenu({ ...state, top: e.clientY + 2, left: Math.min(e.clientX, maxLeft) })
  }

  const fail = (error?: string) => window.alert(error ?? t('failed'))

  // Inline inputs commit on Enter or blur and cancel on Escape. Unmounting the
  // focused input may fire one more blur, so the handlers read the edit state
  // from refs mirrored during render: once Enter or Escape has cleared the
  // state, the late blur finds nothing to commit.
  const creatingRef = useRef<{ parent: string; name: string } | null>(null)
  creatingRef.current = creatingIn ? { parent: creatingIn, name: newName } : null
  const renamingRef = useRef(renaming)
  renamingRef.current = renaming

  const commitCreate = async () => {
    const pending = creatingRef.current
    creatingRef.current = null
    setCreatingIn(null)
    setNewName('')
    if (!pending) return
    const name = pending.name.trim()
    if (!name) return
    const result = await api.createFolder(pending.parent, name)
    if (!result.ok) fail(result.error)
    invalidate([pending.parent])
  }

  const commitRename = async () => {
    const pending = renamingRef.current
    renamingRef.current = null
    setRenaming(null)
    if (!pending) return
    const value = pending.value.trim()
    if (!value || value === nameOf(pending.path)) return
    const result =
      pending.kind === 'folder'
        ? await api.renameFolder(pending.path, value)
        : await api.renameFile(pending.path, value)
    if (!result.ok) fail(result.error)
    else if (result.path) followCurrent(pending.path, result.path)
    if (result.ok && pending.kind === 'folder' && result.path) {
      const renamed = result.path
      setExpanded(
        (prev) =>
          new Set(
            [...prev].map((p) =>
              p === pending.path || isUnder(pending.path, p)
                ? renamed + p.slice(pending.path.length)
                : p,
            ),
          ),
      )
    }
    invalidate([dirOf(pending.path)])
  }

  const confirmDelete = async () => {
    const pending = confirm
    setConfirm(null)
    if (!pending) return
    if (pending.kind === 'folder') await api.deleteFolder(pending.path)
    else await api.deleteFiles([pending.path])
    invalidate([dirOf(pending.path)])
  }

  const doMove = async (paths: string[], targetDir: string, policy: MoveConflictPolicy) => {
    setConflict(null)
    const result = await api.movePaths(paths, targetDir, policy)
    for (const { from, to } of result.moved) followCurrent(from, to)
    invalidate([...new Set([targetDir, ...paths.map(dirOf)])])
    const firstFailure = result.failed[0]
    if (firstFailure) fail(firstFailure.error)
    if (result.conflicts.length > 0) setConflict({ paths: result.conflicts, targetDir })
  }

  const clearDragExpand = () => {
    if (dragExpandTimer.current !== null) {
      window.clearTimeout(dragExpandTimer.current)
      dragExpandTimer.current = null
    }
  }

  const dropProps = (dir: string) => ({
    onDragOver: (e: DragEvent) => {
      if (!e.dataTransfer.types.includes(DRAG_PATHS_MIME)) return
      e.preventDefault()
      e.dataTransfer.dropEffect = 'move'
      if (dropTarget !== dir) {
        setDropTarget(dir)
        clearDragExpand()
        if (!expanded.has(dir)) {
          dragExpandTimer.current = window.setTimeout(
            () => setExpanded((prev) => new Set([...prev, dir])),
            DRAG_EXPAND_DELAY_MS,
          )
        }
      }
    },
    onDragLeave: (e: DragEvent) => {
      if (e.currentTarget.contains(e.relatedTarget as Node | null)) return
      if (dropTarget === dir) setDropTarget(null)
      clearDragExpand()
    },
    onDrop: (e: DragEvent) => {
      e.preventDefault()
      setDropTarget(null)
      clearDragExpand()
      const dropped = ((): unknown => {
        try {
          return JSON.parse(e.dataTransfer.getData(DRAG_PATHS_MIME))
        } catch {
          return []
        }
      })()
      const paths = (Array.isArray(dropped) ? dropped : [])
        .filter((p): p is string => typeof p === 'string')
        .filter((p) => p !== dir && !isUnder(p, dir) && dirOf(p) !== dir)
      if (paths.length > 0) void doMove(paths, dir, 'ask')
    },
  })

  const dragStart = (e: DragEvent, path: string) => {
    e.dataTransfer.setData(DRAG_PATHS_MIME, JSON.stringify([path]))
    e.dataTransfer.effectAllowed = 'move'
    setMenu(null)
  }

  const renderInput = (
    value: string,
    onChange: (v: string) => void,
    onCommit: () => void,
    onCancel: () => void,
    placeholder?: string,
  ) => (
    <input
      className="fp-input"
      value={value}
      autoFocus
      placeholder={placeholder}
      onFocus={(e) => e.target.select()}
      onClick={(e) => e.stopPropagation()}
      onChange={(e) => onChange(e.target.value)}
      onBlur={onCommit}
      onKeyDown={(e) => {
        e.stopPropagation()
        if (e.nativeEvent.isComposing) return
        if (e.key === 'Enter') onCommit()
        if (e.key === 'Escape') onCancel()
      }}
    />
  )

  const renderFile = (file: FileEntry, depth: number) => {
    const isCurrent = currentPath === file.path
    const isRenaming = renaming?.path === file.path
    return (
      <li key={file.path} className="fp-item">
        <div
          className={`fp-row fp-file${isCurrent ? ' current' : ''}`}
          style={{ paddingLeft: 8 + depth * 14 }}
          role="treeitem"
          aria-selected={isCurrent}
          tabIndex={0}
          title={file.path}
          draggable={!isRenaming}
          onDragStart={(e) => dragStart(e, file.path)}
          onClick={() => {
            if (!isRenaming && !isCurrent) void api.openPath(file.path)
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !isCurrent) void api.openPath(file.path)
          }}
          onContextMenu={(e) => openMenuAt(e, { path: file.path, kind: 'file', name: file.name })}
        >
          <span className="fp-chevron" aria-hidden="true" />
          <FileBadge ext={file.ext} />
          {isRenaming ? (
            renderInput(
              renaming.value,
              (v) => setRenaming({ ...renaming, value: v }),
              () => void commitRename(),
              () => setRenaming(null),
            )
          ) : (
            <span className="fp-name">{file.name}</span>
          )}
          <button
            className="fp-more"
            aria-label={t('moreActions', { name: file.name })}
            aria-expanded={menu?.path === file.path}
            onClick={(e) => openMenuAt(e, { path: file.path, kind: 'file', name: file.name })}
          >
            <MoreDots />
          </button>
        </div>
      </li>
    )
  }

  const renderFolder = (
    entry: { path: string; name: string; hasSubfolders: boolean },
    depth: number,
  ): ReactElement => {
    const isRoot = root !== null && entry.path === root.path
    const isOpen = expanded.has(entry.path)
    const listing = listings.get(entry.path)
    const isRenaming = renaming?.path === entry.path
    const kind = isRoot ? 'root' : 'folder'
    const empty = listing && listing.folders.length === 0 && listing.files.length === 0
    return (
      <li key={entry.path} className="fp-item">
        <div
          className={`fp-row fp-dir${dropTarget === entry.path ? ' drop-target' : ''}`}
          style={{ paddingLeft: 8 + depth * 14 }}
          role="treeitem"
          aria-expanded={isOpen}
          tabIndex={0}
          title={entry.path}
          draggable={!isRoot && !isRenaming}
          onDragStart={(e) => dragStart(e, entry.path)}
          onClick={() => {
            if (!isRenaming) toggle(entry.path)
          }}
          onKeyDown={(e) => {
            if (
              e.key === 'Enter' ||
              (e.key === 'ArrowRight' && !isOpen) ||
              (e.key === 'ArrowLeft' && isOpen)
            )
              toggle(entry.path)
          }}
          onContextMenu={(e) => openMenuAt(e, { path: entry.path, kind, name: entry.name })}
          {...dropProps(entry.path)}
        >
          <span className="fp-chevron" aria-hidden="true">
            <svg
              width="12"
              height="12"
              viewBox="0 0 12 12"
              style={{ transform: isOpen ? 'rotate(90deg)' : undefined }}
            >
              <path
                d="M4.5 2.5l4 3.5-4 3.5"
                stroke="currentColor"
                strokeWidth="1.3"
                strokeLinecap="round"
                fill="none"
              />
            </svg>
          </span>
          <FolderIcon open={isOpen} />
          {isRenaming ? (
            renderInput(
              renaming.value,
              (v) => setRenaming({ ...renaming, value: v }),
              () => void commitRename(),
              () => setRenaming(null),
            )
          ) : (
            <span className="fp-name">{entry.name}</span>
          )}
          <button
            className="fp-more"
            aria-label={t('moreActions', { name: entry.name })}
            aria-expanded={menu?.path === entry.path}
            onClick={(e) => openMenuAt(e, { path: entry.path, kind, name: entry.name })}
          >
            <MoreDots />
          </button>
        </div>
        {isOpen && (
          <ul className="fp-children" role="group">
            {creatingIn === entry.path && (
              <li className="fp-item">
                <div className="fp-row fp-dir" style={{ paddingLeft: 8 + (depth + 1) * 14 }}>
                  <span className="fp-chevron" aria-hidden="true" />
                  <FolderIcon open={false} />
                  {renderInput(
                    newName,
                    setNewName,
                    () => void commitCreate(),
                    () => {
                      setCreatingIn(null)
                      setNewName('')
                    },
                    t('untitledFolder'),
                  )}
                </div>
              </li>
            )}
            {(listing?.folders ?? []).map((child) => renderFolder(child, depth + 1))}
            {(listing?.files ?? []).map((file) => renderFile(file, depth + 1))}
            {empty && creatingIn !== entry.path && (
              <li className="fp-empty" style={{ paddingLeft: 8 + (depth + 1) * 14 + 22 }}>
                {t('emptyFolder')}
              </li>
            )}
          </ul>
        )}
      </li>
    )
  }

  return (
    <aside className="files-pane" aria-label={t('title')}>
      <div className="fp-head">
        <span className="fp-title">{t('title')}</span>
        {root?.usable && (
          <button
            className="fp-head-btn"
            aria-label={t('newFolder')}
            title={t('newFolder')}
            onClick={() => {
              const parent =
                currentPath && isUnder(root.path, currentPath) ? dirOf(currentPath) : root.path
              setExpanded((prev) => new Set([...prev, parent]))
              setCreatingIn(parent)
              setNewName('')
            }}
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
        <button
          className="fp-head-btn"
          aria-label={t('close')}
          title={t('close')}
          onClick={onClose}
        >
          <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
            <path
              d="M2 2l10 10M12 2L2 12"
              stroke="currentColor"
              strokeWidth="1.6"
              strokeLinecap="round"
            />
          </svg>
        </button>
      </div>
      {root && !root.usable ? (
        <p className="fp-unusable">{t('rootUnusable')}</p>
      ) : (
        <ul className="fp-tree" role="tree">
          {root && renderFolder({ path: root.path, name: root.name, hasSubfolders: true }, 0)}
        </ul>
      )}

      {menu && (
        <div
          className="fp-menu"
          role="menu"
          ref={menuRef}
          style={{ top: menu.top, left: menu.left }}
        >
          {menu.kind === 'file' && (
            <button
              role="menuitem"
              onClick={() => {
                setMenu(null)
                if (menu.path !== currentPath) void api.openPath(menu.path)
              }}
            >
              {t('open')}
            </button>
          )}
          {menu.kind !== 'file' && (
            <button
              role="menuitem"
              onClick={() => {
                setMenu(null)
                setExpanded((prev) => new Set([...prev, menu.path]))
                setCreatingIn(menu.path)
                setNewName('')
              }}
            >
              {t('newSubfolder')}
            </button>
          )}
          {menu.kind !== 'root' && (
            <button
              role="menuitem"
              onClick={() => {
                setMenu(null)
                setRenaming({
                  path: menu.path,
                  value: menu.name,
                  kind: menu.kind === 'file' ? 'file' : 'folder',
                })
              }}
            >
              {t('rename')}
            </button>
          )}
          <button
            role="menuitem"
            onClick={() => {
              setMenu(null)
              void api.revealPath(menu.path)
            }}
          >
            {t('revealInFolder')}
          </button>
          {menu.kind !== 'root' && (
            <>
              <div className="fp-menu-divider" />
              <button
                role="menuitem"
                className="danger"
                onClick={() => {
                  setMenu(null)
                  setConfirm({ path: menu.path, kind: menu.kind === 'file' ? 'file' : 'folder' })
                }}
              >
                {menu.kind === 'file' ? t('deleteFile') : t('deleteFolder')}
              </button>
            </>
          )}
        </div>
      )}

      {confirm && (
        <div className="fp-overlay" onClick={() => setConfirm(null)}>
          <div
            className="fp-dialog"
            role="dialog"
            aria-modal="true"
            onClick={(e) => e.stopPropagation()}
          >
            <p>
              {confirm.kind === 'folder'
                ? t('deleteFolderConfirm', { name: nameOf(confirm.path) })
                : t('deleteFileConfirm', { name: nameOf(confirm.path) })}
            </p>
            <div className="fp-dialog-buttons">
              <button className="fp-btn" autoFocus onClick={() => setConfirm(null)}>
                {t('cancel')}
              </button>
              <button className="fp-btn danger" onClick={() => void confirmDelete()}>
                {t('delete')}
              </button>
            </div>
          </div>
        </div>
      )}

      {conflict && (
        <div className="fp-overlay" onClick={() => setConflict(null)}>
          <div
            className="fp-dialog"
            role="dialog"
            aria-modal="true"
            onClick={(e) => e.stopPropagation()}
          >
            <h3>{t('conflictTitle')}</h3>
            <p>{t('conflictBody')}</p>
            <ul className="fp-dialog-list">
              {conflict.paths.slice(0, 5).map((p) => (
                <li key={p}>{nameOf(p)}</li>
              ))}
            </ul>
            <div className="fp-dialog-buttons">
              <button className="fp-btn" autoFocus onClick={() => setConflict(null)}>
                {t('cancel')}
              </button>
              <button
                className="fp-btn"
                onClick={() => void doMove(conflict.paths, conflict.targetDir, 'keepBoth')}
              >
                {t('conflictKeepBoth')}
              </button>
              <button
                className="fp-btn danger"
                onClick={() => void doMove(conflict.paths, conflict.targetDir, 'replace')}
              >
                {t('conflictReplace')}
              </button>
            </div>
          </div>
        </div>
      )}
    </aside>
  )
}

function MoreDots() {
  return (
    <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true">
      <circle cx="3.2" cy="8" r="1.35" fill="currentColor" />
      <circle cx="8" cy="8" r="1.35" fill="currentColor" />
      <circle cx="12.8" cy="8" r="1.35" fill="currentColor" />
    </svg>
  )
}

/** docked tab on the container's left edge while the pane is closed; the host container must be position: relative */
export function FilesEdgeTab({ lang, onOpen }: { lang: Lang; onOpen: () => void }): ReactElement {
  const label = tFilesPane(lang, 'title')
  return (
    <button className="files-edge-tab" onClick={onOpen} aria-label={label} title={label}>
      <FolderIcon open={false} />
      <span>{label}</span>
    </button>
  )
}
