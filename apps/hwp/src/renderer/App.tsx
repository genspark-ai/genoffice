import { useCallback, useEffect, useState } from 'react'
import { AiPanel, GensparkMark } from './ai/AiPanel'
import { HwpStudio } from './HwpStudio'
import { useI18n } from './i18n/locale'
import type { HangulStudioFacade } from './studio-text'

export default function App() {
  const { t } = useI18n()
  const [path, setPath] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [studio, setStudio] = useState<HangulStudioFacade | null>(null)
  const [aiOpen, setAiOpen] = useState(() => localStorage.getItem('hwp.showAi') !== '0')

  const markDirty = useCallback((dirty: boolean) => {
    window.hwpApi.setDirty(dirty)
  }, [])

  useEffect(() => {
    const offRename = window.hwpApi.onFileRenamed(setPath)
    return offRename
  }, [])

  useEffect(() => {
    localStorage.setItem('hwp.showAi', aiOpen ? '1' : '0')
  }, [aiOpen])

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
          <button
            type="button"
            className="hwp-save-error-dismiss"
            onClick={() => setSaveError(null)}
          >
            Dismiss
          </button>
        </div>
      )}
      <div className="hwp-main">
        <div className={`ai-dock${aiOpen ? '' : ' collapsed'}`}>
          {!aiOpen && (
            <button
              className="ai-rail"
              data-tip={t('aiOpenAssistant')}
              aria-label={t('aiOpenAssistant')}
              onClick={() => setAiOpen(true)}
            >
              <GensparkMark size={22} />
            </button>
          )}
          <AiPanel facade={studio} filePath={path} onCollapse={() => setAiOpen(false)} />
        </div>
        <HwpStudio
          path={path}
          onPath={setPath}
          onDirty={markDirty}
          onError={setError}
          onSaveError={setSaveError}
          onReady={setStudio}
        />
      </div>
    </div>
  )
}
