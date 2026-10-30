import { useEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import type { FolderRoot } from '../../shared/home-api'
import { useI18n } from './locale'
import { useFolderListings } from './useFolderListings'
import { Chevron, FolderIcon, isUnder, locationLabel, rootOf } from './home-shared'

// ── Move-to-folder picker ────────────────────────────────

export interface FolderPickerProps {
  roots: readonly FolderRoot[]
  /** folders the moved items already live in (greyed, not selectable) */
  currentDirs: ReadonlySet<string>
  /** folders being moved: they and their descendants cannot be targets */
  movingDirs: readonly string[]
  count: number
  onCancel: () => void
  onPick: (dir: string) => void
}

export function FolderPicker({
  roots,
  currentDirs,
  movingDirs,
  count,
  onCancel,
  onPick,
}: FolderPickerProps): ReactElement {
  const { t } = useI18n()
  const { listings, load } = useFolderListings()
  const [expanded, setExpanded] = useState<ReadonlySet<string>>(
    () => new Set(roots.map((r) => r.path)),
  )
  const [picked, setPicked] = useState<string | null>(null)
  const [query, setQuery] = useState('')
  const [creatingIn, setCreatingIn] = useState<string | null>(null)
  const [newName, setNewName] = useState('')

  useEffect(() => {
    for (const dir of expanded) load(dir)
  }, [expanded, load])

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onCancel()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onCancel])

  const disabledDir = (dir: string) =>
    currentDirs.has(dir) ||
    movingDirs.some((m) => isUnder(m, dir)) ||
    rootOf(dir, roots)?.usable === false

  // Enter or blur commits, Escape cancels; the blur an unmount may fire reads
  // the edit from a ref mirrored during render, so a finished edit is a no-op
  const creatingRef = useRef<{ parent: string; name: string } | null>(null)
  creatingRef.current = creatingIn ? { parent: creatingIn, name: newName } : null
  const commitCreate = async () => {
    const pending = creatingRef.current
    creatingRef.current = null
    setCreatingIn(null)
    setNewName('')
    const parent = pending?.parent
    const name = pending?.name.trim()
    if (!parent || !name) return
    const result = await window.aiOffice.createFolder(parent, name)
    if (!result.ok) {
      window.alert(result.error ?? t('renameFailed'))
      return
    }
    load(parent, true)
    setExpanded((prev) => new Set([...prev, parent]))
    if (result.path) setPicked(result.path)
  }

  // name search flattens the tree to every loaded folder whose name matches;
  // folders not yet expanded are loaded on the fly one level at a time
  const needle = query.trim().toLowerCase()
  const renderNode = (entry: { path: string; name: string }, depth: number): ReactElement => {
    const listing = listings.get(entry.path)
    const isOpen = expanded.has(entry.path)
    const children = listing?.folders ?? []
    const disabled = disabledDir(entry.path)
    const isCurrent = currentDirs.has(entry.path)
    return (
      <li key={entry.path}>
        <div
          className={`picker-row${picked === entry.path ? ' active' : ''}${disabled ? ' disabled' : ''}`}
          style={{ paddingLeft: 8 + depth * 16 }}
          role="treeitem"
          aria-selected={picked === entry.path}
          aria-expanded={children.length > 0 ? isOpen : undefined}
          onClick={() => {
            if (!disabled) setPicked(entry.path)
          }}
          onDoubleClick={() => {
            if (!disabled) onPick(entry.path)
          }}
        >
          <button
            className="tree-chevron"
            tabIndex={-1}
            aria-hidden="true"
            style={{ visibility: listing && children.length === 0 ? 'hidden' : undefined }}
            onClick={(e) => {
              e.stopPropagation()
              setExpanded((prev) => {
                const next = new Set(prev)
                if (next.has(entry.path)) next.delete(entry.path)
                else next.add(entry.path)
                return next
              })
            }}
          >
            <Chevron open={isOpen} />
          </button>
          <FolderIcon open={isOpen} />
          <span className="picker-name">{entry.name}</span>
          {isCurrent && <span className="picker-current">{t('currentFolder')}</span>}
        </div>
        {isOpen && (
          <ul role="group">
            {children.map((child) => renderNode(child, depth + 1))}
            {creatingIn === entry.path && (
              <li>
                <div className="picker-row" style={{ paddingLeft: 8 + (depth + 1) * 16 }}>
                  <span className="tree-chevron" aria-hidden="true" />
                  <FolderIcon />
                  <input
                    className="folder-rename-input inline"
                    autoFocus
                    placeholder={t('untitledFolder')}
                    value={newName}
                    onChange={(e) => setNewName(e.target.value)}
                    onBlur={() => void commitCreate()}
                    onKeyDown={(e) => {
                      e.stopPropagation()
                      if (e.nativeEvent.isComposing) return
                      if (e.key === 'Enter') void commitCreate()
                      if (e.key === 'Escape') {
                        setCreatingIn(null)
                        setNewName('')
                      }
                    }}
                  />
                </div>
              </li>
            )}
          </ul>
        )}
      </li>
    )
  }

  const renderSearch = (): ReactElement => {
    const hits: Array<{ path: string; name: string; rel: string }> = []
    for (const listing of listings.values()) {
      for (const f of listing.folders) {
        if (f.name.toLowerCase().includes(needle)) {
          hits.push({ path: f.path, name: f.name, rel: locationLabel(f.path + '/x', roots) })
        }
        if (!listings.has(f.path)) load(f.path)
      }
    }
    hits.sort((a, b) => a.name.localeCompare(b.name))
    if (hits.length === 0) return <p className="picker-empty">{t('noMatchingFolders')}</p>
    return (
      <ul role="tree">
        {hits.map((h) => {
          const disabled = disabledDir(h.path)
          return (
            <li key={h.path}>
              <div
                className={`picker-row${picked === h.path ? ' active' : ''}${disabled ? ' disabled' : ''}`}
                role="treeitem"
                aria-selected={picked === h.path}
                onClick={() => {
                  if (!disabled) setPicked(h.path)
                }}
                onDoubleClick={() => {
                  if (!disabled) onPick(h.path)
                }}
              >
                <span className="tree-chevron" aria-hidden="true" />
                <FolderIcon />
                <span className="picker-name">{h.name}</span>
                <span className="picker-rel">{h.rel}</span>
              </div>
            </li>
          )
        })}
      </ul>
    )
  }

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div
        className="modal picker-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('moveToFolderTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <h3>{t('moveToFolderTitle')}</h3>
        <input
          className="picker-search"
          type="search"
          placeholder={t('searchFolders')}
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
        <div className="picker-tree">
          {needle ? (
            renderSearch()
          ) : (
            <ul role="tree">{roots.map((root) => renderNode(root, 0))}</ul>
          )}
        </div>
        <div className="modal-buttons picker-buttons">
          <button
            className="btn btn-secondary picker-new"
            onClick={() => {
              const parent = picked ?? roots.find((r) => r.usable)?.path
              if (!parent) return
              setExpanded((prev) => new Set([...prev, parent]))
              setCreatingIn(parent)
              setNewName('')
            }}
          >
            {t('newFolder')}
          </button>
          <span className="picker-spacer" />
          <button className="btn btn-secondary" onClick={onCancel}>
            {t('cancel')}
          </button>
          <button
            className="btn btn-primary"
            disabled={!picked || disabledDir(picked)}
            onClick={() => {
              if (picked) onPick(picked)
            }}
          >
            {t('moveCount', { n: count })}
          </button>
        </div>
      </div>
    </div>
  )
}
