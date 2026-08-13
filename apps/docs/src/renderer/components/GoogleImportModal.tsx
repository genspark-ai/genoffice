import { useEffect, useMemo, useState } from 'react'
import type { GoogleDocSummary } from '../../shared/ipc'
import { useModalKeys } from './modal-keys'
import { showToast } from './toast-bus'

/**
 * File ▸ Import from Google Docs… picker. Handles the unconfigured/signed-out
 * states inline so the caller doesn't need to branch on auth status first.
 */
export function GoogleImportModal({
  onClose,
  onImport,
}: {
  onClose: () => void
  onImport: (fileId: string) => Promise<void>
}) {
  const modalKeys = useModalKeys(onClose)
  const [status, setStatus] = useState<'checking' | 'unconfigured' | 'signed-out' | 'ready'>(
    'checking',
  )
  const [docs, setDocs] = useState<GoogleDocSummary[]>([])
  const [query, setQuery] = useState('')
  const [loading, setLoading] = useState(false)
  const [importingId, setImportingId] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)

  const refreshList = async () => {
    setLoading(true)
    setError(null)
    const result = await window.desktop.googleListDocs()
    setLoading(false)
    if (!result.ok) {
      setError(result.error)
      return
    }
    setDocs(result.data)
  }

  useEffect(() => {
    void (async () => {
      const auth = await window.desktop.googleAuthStatus()
      if (!auth.configured) {
        setStatus('unconfigured')
        return
      }
      if (!auth.signedIn) {
        setStatus('signed-out')
        return
      }
      setStatus('ready')
      await refreshList()
    })()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const signIn = async () => {
    setLoading(true)
    setError(null)
    const result = await window.desktop.googleSignIn()
    setLoading(false)
    if (!result.ok) {
      setError(result.unconfigured ? null : (result.error ?? 'sign-in failed'))
      if (result.unconfigured) setStatus('unconfigured')
      return
    }
    setStatus('ready')
    await refreshList()
  }

  const filtered = useMemo(
    () => docs.filter((d) => d.name.toLowerCase().includes(query.trim().toLowerCase())),
    [docs, query],
  )

  const runImport = async (fileId: string) => {
    setImportingId(fileId)
    try {
      await onImport(fileId)
      onClose()
    } catch (err) {
      showToast(`Import failed: ${String(err)}`, 'error')
    } finally {
      setImportingId(null)
    }
  }

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal gdh-google-modal">
        <h2>Import from Google Docs</h2>

        {status === 'checking' && <p className="gdh-google-status">Checking Google account…</p>}

        {status === 'unconfigured' && (
          <p className="gdh-google-status">
            Google account not configured. Ask an administrator to set up Google sign-in for this
            app.
          </p>
        )}

        {status === 'signed-out' && (
          <div className="gdh-google-signin">
            <p className="gdh-google-status">
              Sign in to your Google account to import a document.
            </p>
            {error && <p className="gdh-google-error">{error}</p>}
            <button
              type="button"
              className="btn-primary"
              disabled={loading}
              onClick={() => void signIn()}
            >
              {loading ? 'Opening browser…' : 'Sign in with Google'}
            </button>
          </div>
        )}

        {status === 'ready' && (
          <>
            <input
              className="gdh-google-search"
              placeholder="Search documents…"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              autoFocus
            />
            {loading && <p className="gdh-google-status">Loading…</p>}
            {error && <p className="gdh-google-error">{error}</p>}
            {!loading && !error && filtered.length === 0 && (
              <p className="gdh-google-status">No documents found.</p>
            )}
            <ul className="gdh-google-list">
              {filtered.map((d) => (
                <li key={d.id}>
                  <button
                    type="button"
                    className="gdh-google-list-item"
                    disabled={importingId !== null}
                    onClick={() => void runImport(d.id)}
                  >
                    {d.iconLink ? (
                      <img src={d.iconLink} alt="" className="gdh-google-doc-icon" />
                    ) : (
                      <span className="gdh-google-doc-icon-placeholder" />
                    )}
                    <span className="gdh-google-doc-name">{d.name}</span>
                    <span className="gdh-google-doc-time">
                      {new Date(d.modifiedTime).toLocaleDateString()}
                    </span>
                    {importingId === d.id && (
                      <span className="gdh-google-doc-loading">Importing…</span>
                    )}
                  </button>
                </li>
              ))}
            </ul>
          </>
        )}

        <div className="modal-actions">
          <button type="button" className="btn-ghost" onClick={onClose}>
            Cancel
          </button>
        </div>
      </div>
    </div>
  )
}
