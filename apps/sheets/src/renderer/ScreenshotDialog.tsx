import { useEffect, useState } from 'react'

import type { ScreenSourcesResult } from '../shared/desktop-api'
import { useI18n } from './i18n/locale'

/// Excel's Insert → Screenshot: a picker over the OS's capturable surfaces.
/// Screens first, then other windows (our own window is excluded by the main
/// process). Picking one grabs a full-resolution frame and inserts it as a
/// floating picture at the selection.

type LoadState =
  | { phase: 'loading' }
  | { phase: 'denied' }
  | { phase: 'ready'; sources: ScreenSourcesResult['sources'] }

export function ScreenshotDialog({
  onInsert,
  onClose,
}: {
  readonly onInsert: (dataUrl: string, width: number, height: number) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [state, setState] = useState<LoadState>({ phase: 'loading' })
  const [refreshTick, setRefreshTick] = useState(0)
  const [capturingId, setCapturingId] = useState<string | null>(null)
  const [failed, setFailed] = useState(false)

  useEffect(() => {
    let stale = false
    setState({ phase: 'loading' })
    setFailed(false)
    window.desktopApi
      .captureScreenSources()
      .then((result) => {
        if (stale) return
        if (result.status === 'denied') setState({ phase: 'denied' })
        else setState({ phase: 'ready', sources: result.sources })
      })
      .catch(() => {
        if (stale) return
        setFailed(true)
        setState({ phase: 'ready', sources: [] })
      })
    return () => {
      stale = true
    }
  }, [refreshTick])

  const capture = (id: string): void => {
    if (capturingId !== null) return
    setCapturingId(id)
    setFailed(false)
    window.desktopApi
      .captureScreenSource({ id })
      .then((result) => {
        if (!result) {
          setFailed(true)
          setCapturingId(null)
          setRefreshTick((tick) => tick + 1)
          return
        }
        onInsert(`data:${result.mediaType};base64,${result.base64}`, result.width, result.height)
        onClose()
      })
      .catch(() => {
        setFailed(true)
        setCapturingId(null)
      })
  }

  const section = (
    label: string,
    sources: ScreenSourcesResult['sources'],
  ): React.JSX.Element | null =>
    sources.length === 0 ? null : (
      <>
        <p className="dialog-note">{label}</p>
        <div className="screenshot-grid">
          {sources.map((source) => (
            <button
              key={source.id}
              type="button"
              disabled={capturingId !== null}
              className={capturingId === source.id ? 'capturing' : ''}
              title={source.name}
              onClick={() => capture(source.id)}
            >
              {source.thumbnail !== '' && <img src={source.thumbnail} alt="" draggable={false} />}
              <span>{source.name}</span>
            </button>
          ))}
        </div>
      </>
    )

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog screenshot-dialog"
        role="dialog"
        aria-label={t('dlgScreenshotTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgScreenshotTitle')}</header>
        <section className="dialog-body">
          {state.phase === 'loading' && <p className="dialog-note">{t('dlgScreenshotLoading')}</p>}
          {state.phase === 'denied' && <p className="dialog-note">{t('dlgScreenshotDenied')}</p>}
          {state.phase === 'ready' && (
            <>
              {section(
                t('dlgScreenshotScreens'),
                state.sources.filter((source) => source.kind === 'screen'),
              )}
              {section(
                t('dlgScreenshotWindows'),
                state.sources.filter((source) => source.kind === 'window'),
              )}
              {state.sources.length === 0 && !failed && (
                <p className="dialog-note">{t('dlgScreenshotEmpty')}</p>
              )}
            </>
          )}
          {failed && <p className="dialog-note">{t('dlgScreenshotFailed')}</p>}
        </section>
        <div className="dialog-actions">
          <button
            className="secondary"
            disabled={state.phase === 'loading' || capturingId !== null}
            onClick={() => setRefreshTick((tick) => tick + 1)}
          >
            {t('dlgScreenshotRefresh')}
          </button>
          <button className="secondary" onClick={onClose}>
            {t('dlgClose')}
          </button>
        </div>
      </div>
    </div>
  )
}
