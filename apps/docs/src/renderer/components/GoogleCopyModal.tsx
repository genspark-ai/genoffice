import { useState } from 'react'
import { GoogleFolderPicker } from './GoogleFolderPicker'
import { useModalKeys } from './modal-keys'
import { showToast } from './toast-bus'

/**
 * File ▸ Make a copy… dialog for a document already linked to Google Docs.
 * Reuses GoogleFolderPicker (mode="set-default") for the optional destination
 * folder — same breadcrumb browser used everywhere else, just for picking a
 * one-off destination instead of the app default.
 */
export function GoogleCopyModal({
  docTitle,
  onClose,
  onMakeCopy,
}: {
  docTitle: string
  onClose: () => void
  onMakeCopy: (name: string, folderId?: string) => Promise<void>
}) {
  const modalKeys = useModalKeys(onClose)
  const [name, setName] = useState(`Copy of ${docTitle}`)
  const [folder, setFolder] = useState<{ id: string; name: string } | null>(null)
  const [showFolderPicker, setShowFolderPicker] = useState(false)
  const [busy, setBusy] = useState(false)

  const confirm = async () => {
    if (!name.trim()) return
    setBusy(true)
    try {
      await onMakeCopy(name.trim(), folder?.id)
      onClose()
    } catch (err) {
      showToast(`Make a copy failed: ${String(err)}`, 'error')
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
      <div className="modal gdh-google-modal">
        <h2>Make a copy</h2>

        <label htmlFor="gdh-copy-name">Name</label>
        <input
          id="gdh-copy-name"
          className="gdh-google-search"
          value={name}
          onChange={(e) => setName(e.target.value)}
          autoFocus
        />

        <div className="gdh-copy-folder-row">
          <span className="gdh-copy-folder-label">
            Folder: {folder ? `/${folder.name}` : 'My Drive'}
          </span>
          <button type="button" className="btn-ghost" onClick={() => setShowFolderPicker(true)}>
            Choose folder…
          </button>
        </div>

        <div className="modal-actions">
          <button type="button" className="btn-ghost" onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            disabled={busy || !name.trim()}
            onClick={() => void confirm()}
          >
            {busy ? 'Copying…' : 'Make a copy'}
          </button>
        </div>
      </div>

      {showFolderPicker && (
        <GoogleFolderPicker
          mode="set-default"
          onClose={() => setShowFolderPicker(false)}
          onPick={async (picked) => {
            setFolder(picked)
          }}
        />
      )}
    </div>
  )
}
