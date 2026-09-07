import { useCallback, useEffect, useState } from 'react'
import { HwpStudio } from './HwpStudio'

export default function App() {
  const [path, setPath] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [saveError, setSaveError] = useState<string | null>(null)

  const markDirty = useCallback((dirty: boolean) => {
    window.hwpApi.setDirty(dirty)
  }, [])

  useEffect(() => {
    const offRename = window.hwpApi.onFileRenamed(setPath)
    return offRename
  }, [])

  if (error) {
    return (
      <div className="hwp-shell">
        <main className="hwp-page">
          <h1 className="hwp-title">Hangul</h1>
          <p className="hwp-error">{error}</p>
        </main>
      </div>
    )
  }

  return (
    <div className="hwp-shell">
      {saveError && (
        <div className="hwp-save-error" role="alert">
          <p className="hwp-save-error-text">{saveError}</p>
          <button type="button" className="hwp-save-error-dismiss" onClick={() => setSaveError(null)}>
            Dismiss
          </button>
        </div>
      )}
      <HwpStudio
        path={path}
        onPath={setPath}
        onDirty={markDirty}
        onError={setError}
        onSaveError={setSaveError}
      />
    </div>
  )
}
