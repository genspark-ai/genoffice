/**
 * Export as Video: resolution, seconds per slide, whether saved timings and narrations drive
 * the show; then the encoding progress with a cancel button, in the same sheet.
 */
import React, { useEffect, useRef, useState } from 'react'
import { useEscOverlay } from '../esc-overlay'
import { useI18n } from '../i18n/locale'
import type { VideoExportOptions, VideoExportResult } from '../video-export/export-video'

const DEFAULT_SECONDS = 5

export function ExportVideoDialog({
  slideCount,
  run,
  onClose,
}: {
  slideCount: number
  /** Starts the export once the user confirms; resolves when it ends however it ends */
  run: (
    opts: VideoExportOptions,
    onProgress: (fraction: number) => void,
    signal: AbortSignal,
  ) => Promise<VideoExportResult>
  onClose: () => void
}) {
  useEscOverlay(true)
  const { t } = useI18n()
  const [height, setHeight] = useState<1080 | 720>(1080)
  const [seconds, setSeconds] = useState(String(DEFAULT_SECONDS))
  const [useTimings, setUseTimings] = useState(true)
  const [hasTimings, setHasTimings] = useState(false)
  const [progress, setProgress] = useState<number | null>(null)
  const abortRef = useRef<AbortController | null>(null)

  useEffect(() => {
    let cancelled = false
    void Promise.all(
      Array.from({ length: slideCount }, (_, i) => window.slidesApi.getShowTiming(i)),
    ).then((list) => {
      if (!cancelled)
        setHasTimings(list.some((x) => x.advanceMs != null || x.narrationIds.length > 0))
    })
    return () => {
      cancelled = true
    }
  }, [slideCount])

  const secs = Math.min(600, Math.max(1, Number(seconds) || DEFAULT_SECONDS))

  const start = async () => {
    const ac = new AbortController()
    abortRef.current = ac
    setProgress(0)
    try {
      await run(
        { height, defaultSlideMs: secs * 1000, useTimings: hasTimings && useTimings },
        setProgress,
        ac.signal,
      )
    } finally {
      onClose()
    }
  }

  const cancel = () => {
    if (progress == null) onClose()
    else abortRef.current?.abort()
  }

  return (
    <div className="modal-backdrop" onClick={cancel}>
      <div className="modal export-video-dialog" onClick={(e) => e.stopPropagation()}>
        <h2>{t('appExportVideoTitle')}</h2>
        {progress == null ? (
          <>
            <label>
              {t('appExportVideoResolution')}
              <select
                value={height}
                onChange={(e) => setHeight(Number(e.target.value) === 720 ? 720 : 1080)}
              >
                <option value={1080}>Full HD (1080p)</option>
                <option value={720}>HD (720p)</option>
              </select>
            </label>
            <label>
              {t('appExportVideoSeconds')}
              <input
                type="number"
                min={1}
                max={600}
                step={1}
                value={seconds}
                onChange={(e) => setSeconds(e.target.value)}
              />
            </label>
            <label className="export-video-check">
              <input
                type="checkbox"
                disabled={!hasTimings}
                checked={hasTimings && useTimings}
                onChange={(e) => setUseTimings(e.target.checked)}
              />
              {t('appExportVideoUseTimings')}
            </label>
          </>
        ) : (
          <div className="export-video-progress" role="progressbar" aria-valuenow={progress * 100}>
            <div className="export-video-progress-bar" style={{ width: `${progress * 100}%` }} />
            <span>{t('appExportVideoProgress', { percent: Math.floor(progress * 100) })}</span>
          </div>
        )}
        <div className="modal-actions">
          <button onClick={cancel}>{t('appExportVideoCancel')}</button>
          {progress == null && (
            <button className="primary" onClick={() => void start()}>
              {t('appExportVideoCreate')}
            </button>
          )}
        </div>
      </div>
    </div>
  )
}
