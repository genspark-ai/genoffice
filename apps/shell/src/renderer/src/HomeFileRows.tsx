import type { DragEvent as ReactDragEvent, ReactElement, Ref } from 'react'
import type { FolderEntry, FolderRoot, RecentEntry } from '../../shared/home-api'
import { useI18n } from './locale'
import { FileBadge } from './file-badge'
import {
  FolderIcon,
  MoreDots,
  dirOf,
  formatModified,
  formatSize,
  locationLabel,
  rootOf,
} from './home-shared'

/**
 * The two row kinds the plain (recent/starred) and folder tables share,
 * lifted out of Home.tsx. Everything the rows read arrives through props;
 * everything they do goes back through the action bundles, so Home keeps
 * owning the selection / rename / menu state and this file stays markup.
 */

export interface FileRowActions {
  openPath: (path: string) => void
  /** a missing recent entry was clicked: offer list removal */
  reportMissing: (entry: RecentEntry) => void
  toggleSelect: (path: string, on: boolean) => void
  toggleStar: (path: string) => void
  dragStart: (event: ReactDragEvent, paths: string[]) => void
  startRename: (entry: RecentEntry) => void
  setRenameValue: (path: string, value: string) => void
  commitRename: (entry: RecentEntry) => void
  cancelRename: () => void
  setRowMenu: (path: string | null) => void
  startMove: (paths: string[]) => void
  duplicateFile: (path: string) => void
  /** the Starred view's remove action: drop the star, keep the recents entry */
  unstar: (paths: string[]) => void
  removeRecent: (paths: string[]) => void
  deleteFiles: (paths: string[]) => void
  /** starred rows can be moved into a named group (the GroupPicker modal) */
  pickGroup: (paths: string[]) => void
  /** drag-over / drag-leave / drop handlers for one folder row */
  dropProps: (
    dir: string,
    opts: { autoExpand: boolean },
  ) => {
    onDragOver: (event: ReactDragEvent) => void
    onDragLeave: (event: ReactDragEvent) => void
    onDrop: (event: ReactDragEvent) => void
  }
  expandTo: (dir: string) => void
  selectFolder: (dir: string) => void
  setFolderRenaming: (
    value: { path: string; where: 'tree' | 'table'; value: string } | null,
  ) => void
  commitRenameFolder: () => void
  setFolderMenu: (
    menu: {
      path: string
      where: 'tree' | 'table'
      x: number
      y: number
      above: number
      align: 'start' | 'end'
    } | null,
  ) => void
  startRenameFolder: (entry: { path: string; name: string }, where: 'tree' | 'table') => void
  confirmDeleteFolder: (dir: string | null) => void
}

export interface FileRowProps {
  entry: RecentEntry
  context: 'global' | 'folder'
  /** the selected-path pool this row drags as (the row's path or the whole selection) */
  dragPool: string[]
  selected: ReadonlySet<string>
  renaming: { path: string; value: string } | null
  rowMenu: string | null
  /** actions cell of the row whose menu is open — the dismissal guard root */
  rowMenuWrapRef: Ref<HTMLSpanElement>
  roots: readonly FolderRoot[]
  /** which list the plain view shows; decides remove-from-list vs unstar */
  view: 'recent' | 'starred'
  /** a usable folder root exists, so "move to folder" can be offered */
  canMove: boolean
  /** no multi-selection under this row: per-row delete stays available */
  canDelete: boolean
  actions: FileRowActions
}

export function FileRow({
  entry,
  context,
  dragPool,
  selected,
  renaming,
  rowMenu,
  rowMenuWrapRef,
  roots,
  view,
  canMove,
  canDelete,
  actions,
}: FileRowProps): ReactElement {
  const i18n = useI18n()
  const { t } = i18n
  const isRenaming = renaming?.path === entry.path
  const editable = rootOf(entry.path, roots)?.usable !== false
  const menuOpen = rowMenu === entry.path
  return (
    <li className="recent-row" key={entry.path}>
      <div
        className={`recent-item${entry.missing ? ' missing' : ''}`}
        role="button"
        tabIndex={0}
        draggable={!isRenaming && !entry.missing}
        onDragStart={(e) => actions.dragStart(e, dragPool)}
        onClick={() => {
          if (isRenaming) return
          if (entry.missing) actions.reportMissing(entry)
          else actions.openPath(entry.path)
        }}
        onKeyDown={(event) => {
          if (event.key === 'Enter' && event.target === event.currentTarget) {
            if (entry.missing) actions.reportMissing(entry)
            else actions.openPath(entry.path)
          }
        }}
      >
        <span className="col-check" onClick={(event) => event.stopPropagation()}>
          <input
            type="checkbox"
            className="row-check"
            checked={selected.has(entry.path)}
            onChange={(event) => actions.toggleSelect(entry.path, event.target.checked)}
            aria-label={t('selectFile', { name: entry.name })}
          />
        </span>
        <span className="recent-icon">
          <FileBadge ext={entry.ext} size={24} />
        </span>
        {isRenaming ? (
          <input
            className="rename-input"
            value={renaming.value}
            autoFocus
            onFocus={(event) => event.target.select()}
            onClick={(event) => event.stopPropagation()}
            onChange={(event) => actions.setRenameValue(entry.path, event.target.value)}
            onBlur={() => actions.commitRename(entry)}
            onKeyDown={(event) => {
              event.stopPropagation()
              if (event.nativeEvent.isComposing) return
              if (event.key === 'Enter') actions.commitRename(entry)
              if (event.key === 'Escape') actions.cancelRename()
            }}
          />
        ) : (
          <span className="recent-name">{entry.name}</span>
        )}
        <span className="recent-path" title={dirOf(entry.path)}>
          {locationLabel(entry.path, roots)}
        </span>
        <span className="recent-time">
          {entry.missing ? '—' : formatModified(entry.mtimeMs, i18n)}
        </span>
        <span className="recent-size">{entry.missing ? '—' : formatSize(entry.sizeBytes)}</span>
        <button
          className={`star-btn${entry.starred ? ' starred' : ''}`}
          aria-label={entry.starred ? t('unstar') : t('star')}
          onClick={(event) => {
            event.stopPropagation()
            actions.toggleStar(entry.path)
          }}
        >
          <svg width="15" height="15" viewBox="0 0 16 16" aria-hidden="true">
            <path
              d="M8 1.9l1.9 3.85 4.25.62-3.07 3 .72 4.23L8 11.6l-3.8 2 .72-4.23-3.07-3 4.25-.62z"
              fill={entry.starred ? '#f5a623' : 'none'}
              stroke={entry.starred ? '#f5a623' : 'currentColor'}
              strokeWidth="1.2"
              strokeLinejoin="round"
            />
          </svg>
        </button>
        <span
          className="recent-actions"
          ref={menuOpen ? rowMenuWrapRef : undefined}
          onClick={(event) => event.stopPropagation()}
        >
          <button
            className="more-btn"
            aria-label={t('moreActions')}
            aria-expanded={menuOpen}
            onClick={() => actions.setRowMenu(menuOpen ? null : entry.path)}
          >
            <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
              <circle cx="3.2" cy="8" r="1.4" fill="currentColor" />
              <circle cx="8" cy="8" r="1.4" fill="currentColor" />
              <circle cx="12.8" cy="8" r="1.4" fill="currentColor" />
            </svg>
          </button>
          {menuOpen && (
            <div className="row-menu" role="menu">
              <button
                role="menuitem"
                onClick={() => {
                  actions.setRowMenu(null)
                  actions.openPath(entry.path)
                }}
              >
                {t('open')}
              </button>
              <button
                role="menuitem"
                onClick={() => {
                  actions.setRowMenu(null)
                  void window.aiOffice.revealPath(entry.path)
                }}
              >
                {t('revealInFolder')}
              </button>
              <button
                role="menuitem"
                onClick={() => {
                  actions.setRowMenu(null)
                  void navigator.clipboard.writeText(entry.path)
                }}
              >
                {t('copyPath')}
              </button>
              {canMove && editable && !entry.missing && (
                <>
                  <div className="row-menu-divider" />
                  <button role="menuitem" onClick={() => actions.startMove([entry.path])}>
                    {t('moveToFolder')}
                  </button>
                </>
              )}
              {editable && (
                <>
                  <div className="row-menu-divider" />
                  <button role="menuitem" onClick={() => actions.startRename(entry)}>
                    {t('rename')}
                  </button>
                  <button role="menuitem" onClick={() => actions.duplicateFile(entry.path)}>
                    {t('duplicate')}
                  </button>
                </>
              )}
              {entry.starred && (
                <>
                  <div className="row-menu-divider" />
                  <button
                    role="menuitem"
                    onClick={() => {
                      actions.setRowMenu(null)
                      actions.pickGroup([entry.path])
                    }}
                  >
                    {t('moveToGroup')}
                  </button>
                </>
              )}
              {canDelete && (context === 'global' || editable) && (
                <>
                  <div className="row-menu-divider" />
                  {context === 'global' &&
                    // same root cause as the selection bar: in the Starred
                    // view "remove from list" must unstar, not drop recents
                    (view === 'starred' ? (
                      <button role="menuitem" onClick={() => actions.unstar([entry.path])}>
                        {t('unstar')}
                      </button>
                    ) : (
                      <button role="menuitem" onClick={() => actions.removeRecent([entry.path])}>
                        {t('removeFromList')}
                      </button>
                    ))}
                  {editable && (
                    <button
                      role="menuitem"
                      className="danger"
                      onClick={() => actions.deleteFiles([entry.path])}
                    >
                      {t('deleteFiles')}
                    </button>
                  )}
                </>
              )}
            </div>
          )}
        </span>
      </div>
    </li>
  )
}

export interface SubfolderRowProps {
  entry: FolderEntry
  /** the folder row currently hovered by a drag (sidebar or table) */
  dropTarget: string | null
  renaming: { path: string; where: 'tree' | 'table'; value: string } | null
  folderMenuOpen: boolean
  folderMenuWrapRef: Ref<HTMLDivElement>
  folderMenuPopRef: Ref<HTMLDivElement>
  roots: readonly FolderRoot[]
  actions: FileRowActions
}

/** sub-folder row in the folder view's table: enter on click, same … menu as the tree */
export function SubfolderRow({
  entry,
  dropTarget,
  renaming,
  folderMenuOpen,
  folderMenuWrapRef,
  folderMenuPopRef,
  roots,
  actions,
}: SubfolderRowProps): ReactElement {
  const i18n = useI18n()
  const { t } = i18n
  const isRenaming = renaming?.where === 'table' && renaming.path === entry.path
  const editable = rootOf(entry.path, roots)?.usable !== false
  return (
    <li className="recent-row" key={entry.path}>
      <div
        className={`recent-item folder-item${dropTarget === entry.path ? ' drop-target' : ''}`}
        role="button"
        tabIndex={0}
        draggable={!isRenaming}
        onDragStart={(e) => actions.dragStart(e, [entry.path])}
        {...actions.dropProps(entry.path, { autoExpand: false })}
        onClick={() => {
          if (isRenaming) return
          actions.expandTo(entry.path)
          actions.selectFolder(entry.path)
        }}
        onKeyDown={(event) => {
          if (event.key === 'Enter' && event.target === event.currentTarget) {
            actions.expandTo(entry.path)
            actions.selectFolder(entry.path)
          }
        }}
        onContextMenu={(e) => {
          e.preventDefault()
          actions.setFolderMenu({
            path: entry.path,
            where: 'table',
            x: e.clientX + 2,
            y: e.clientY + 2,
            above: e.clientY + 2,
            align: 'start',
          })
        }}
      >
        <span className="col-check" aria-hidden="true" />
        <span className="recent-icon folder-badge">
          <FolderIcon size={22} />
        </span>
        {isRenaming ? (
          <input
            className="rename-input"
            value={renaming.value}
            autoFocus
            onFocus={(event) => event.target.select()}
            onClick={(event) => event.stopPropagation()}
            onChange={(event) =>
              actions.setFolderRenaming({
                path: entry.path,
                where: 'table',
                value: event.target.value,
              })
            }
            onBlur={() => void actions.commitRenameFolder()}
            onKeyDown={(event) => {
              event.stopPropagation()
              if (event.nativeEvent.isComposing) return
              if (event.key === 'Enter') void actions.commitRenameFolder()
              if (event.key === 'Escape') actions.setFolderRenaming(null)
            }}
          />
        ) : (
          <span className="recent-name">{entry.name}</span>
        )}
        <span className="recent-path">{t('folderType')}</span>
        <span className="recent-time">{formatModified(entry.mtimeMs, i18n)}</span>
        <span className="recent-size">—</span>
        <span />
        <span
          className="recent-actions"
          ref={folderMenuOpen ? folderMenuWrapRef : undefined}
          onClick={(event) => event.stopPropagation()}
        >
          <button
            className="more-btn"
            aria-label={t('folderMoreActions', { name: entry.name })}
            aria-expanded={folderMenuOpen}
            onClick={(e) => {
              if (folderMenuOpen) {
                actions.setFolderMenu(null)
                return
              }
              const rect = e.currentTarget.getBoundingClientRect()
              actions.setFolderMenu({
                path: entry.path,
                where: 'table',
                x: rect.right,
                y: rect.bottom + 4,
                above: rect.top - 4,
                align: 'end',
              })
            }}
          >
            <MoreDots />
          </button>
          {folderMenuOpen && (
            <div className="folder-menu" role="menu" ref={folderMenuPopRef}>
              <button
                role="menuitem"
                onClick={() => {
                  actions.setFolderMenu(null)
                  actions.expandTo(entry.path)
                  actions.selectFolder(entry.path)
                }}
              >
                {t('open')}
              </button>
              {editable && (
                <>
                  <button role="menuitem" onClick={() => actions.startRenameFolder(entry, 'table')}>
                    {t('rename')}
                  </button>
                  <button role="menuitem" onClick={() => actions.startMove([entry.path])}>
                    {t('moveToFolder')}
                  </button>
                </>
              )}
              <button
                role="menuitem"
                onClick={() => {
                  actions.setFolderMenu(null)
                  void window.aiOffice.revealPath(entry.path)
                }}
              >
                {t('revealInFolder')}
              </button>
              {editable && (
                <>
                  <div className="row-menu-divider" />
                  <button
                    role="menuitem"
                    className="danger"
                    onClick={() => {
                      actions.setFolderMenu(null)
                      actions.confirmDeleteFolder(entry.path)
                    }}
                  >
                    {t('deleteFolder')}
                  </button>
                </>
              )}
            </div>
          )}
        </span>
      </div>
    </li>
  )
}
