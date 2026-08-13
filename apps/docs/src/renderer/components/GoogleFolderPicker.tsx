import { useEffect, useRef, useState } from 'react'
import type { GoogleFolderSummary } from '../../shared/ipc'
import { useModalKeys } from './modal-keys'
import { showToast } from './toast-bus'

interface Crumb {
  id: string | null // null = My Drive root
  name: string
}

const ROOT_CRUMB: Crumb = { id: null, name: 'My Drive' }

/**
 * Breadcrumb folder browser used both to set the default "Send to Google
 * Docs" destination (mode="set-default") and to move an already-sent doc
 * (mode="move"). Same UI, different primary action + confirm copy.
 */
export function GoogleFolderPicker({
  mode,
  onClose,
  onPick,
  onClearDefault,
}: {
  mode: 'set-default' | 'move'
  onClose: () => void
  onPick: (folder: { id: string; name: string } | null) => Promise<void>
  /** only used in set-default mode */
  onClearDefault?: () => Promise<void>
}) {
  const modalKeys = useModalKeys(onClose)
  const [trail, setTrail] = useState<Crumb[]>([ROOT_CRUMB])
  const [folders, setFolders] = useState<GoogleFolderSummary[]>([])
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)

  const current = trail[trail.length - 1]

  // Guards against a stale response from an earlier navigation landing after
  // the user has already navigated further (e.g. rapid double-click into a
  // folder): only the response matching the request that's still current wins.
  const requestVersionRef = useRef(0)

  const load = async (parentId: string | null) => {
    const version = ++requestVersionRef.current
    setLoading(true)
    setError(null)
    const result = await window.desktop.googleListFolders(parentId ?? undefined)
    if (version !== requestVersionRef.current) return // superseded by a newer navigation
    setLoading(false)
    if (!result.ok) {
      setError(result.error)
      return
    }
    setFolders(result.data)
  }

  useEffect(() => {
    void load(current.id)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [current.id])

  const openFolder = (folder: GoogleFolderSummary) => {
    setTrail((t) => [...t, { id: folder.id, name: folder.name }])
  }

  const goToCrumb = (index: number) => {
    setTrail((t) => t.slice(0, index + 1))
  }

  const confirm = async () => {
    setBusy(true)
    try {
      await onPick(current.id ? { id: current.id, name: current.name } : null)
      onClose()
    } catch (err) {
      showToast(`${mode === 'move' ? 'Move' : 'Set folder'} failed: ${String(err)}`, 'error')
    } finally {
      setBusy(false)
    }
  }

  const clearDefault = async () => {
    if (!onClearDefault) return
    setBusy(true)
    try {
      await onClearDefault()
      onClose()
    } catch (err) {
      showToast(`Clear default failed: ${String(err)}`, 'error')
    } finally {
      setBusy(false)
    }
  }

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal gdh-google-modal gdh-folder-picker">
        <h2>{mode === 'move' ? 'Move to folder' : 'Google Drive folder'}</h2>

        <div className="gdh-folder-breadcrumb">
          {trail.map((crumb, i) => (
            <span key={`${crumb.id ?? 'root'}-${i}`} className="gdh-folder-crumb-wrap">
              {i > 0 && <span className="gdh-folder-crumb-sep">›</span>}
              <button
                type="button"
                className="gdh-folder-crumb"
                disabled={i === trail.length - 1}
                onClick={() => goToCrumb(i)}
              >
                {crumb.name}
              </button>
            </span>
          ))}
        </div>

        {loading && <p className="gdh-google-status">Loading…</p>}
        {error && <p className="gdh-google-error">{error}</p>}
        {!loading && !error && folders.length === 0 && (
          <p className="gdh-google-status">No subfolders here.</p>
        )}

        <ul className="gdh-google-list gdh-folder-list">
          {folders.map((f) => (
            <li key={f.id}>
              <button type="button" className="gdh-google-list-item" onClick={() => openFolder(f)}>
                <span className="gdh-folder-icon" aria-hidden />
                <span className="gdh-google-doc-name">{f.name}</span>
              </button>
            </li>
          ))}
        </ul>

        <div className="modal-actions gdh-folder-actions">
          <span className="gdh-folder-actions-left">
            {mode === 'set-default' && onClearDefault && (
              <button
                type="button"
                className="btn-ghost"
                disabled={busy}
                onClick={() => void clearDefault()}
              >
                Clear default
              </button>
            )}
          </span>
          <span>
            <button type="button" className="btn-ghost" onClick={onClose}>
              Cancel
            </button>
            <button
              type="button"
              className="btn-primary"
              disabled={busy}
              onClick={() => void confirm()}
            >
              {busy
                ? mode === 'move'
                  ? 'Moving…'
                  : 'Saving…'
                : mode === 'move'
                  ? `Move here${current.id ? '' : ' (My Drive)'}`
                  : 'Use this folder'}
            </button>
          </span>
        </div>
      </div>
    </div>
  )
}
