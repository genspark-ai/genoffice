import { useI18n } from '../i18n/locale'
import { formatClock } from '../slideshow-utils'
import type { RecordingPhase } from './use-recording'

/** The show's recording bar: what is being recorded, the time on this slide and in all, pause. */
export function RecordingBar({
  phase,
  slideElapsedMs,
  elapsedMs,
  onTogglePause,
}: {
  phase: RecordingPhase
  slideElapsedMs: number
  elapsedMs: number
  onTogglePause: () => void
}) {
  const { t } = useI18n()
  if (phase === 'failed') {
    return <div className="ss-record ss-record-failed">{t('paneShowRecordFailed')}</div>
  }
  return (
    <div className="ss-record" onClick={(e) => e.stopPropagation()}>
      <span className={`ss-record-dot${phase === 'recording' ? ' on' : ''}`} aria-hidden="true" />
      <span className="ss-record-state">
        {t(phase === 'paused' ? 'paneShowRecordPaused' : 'paneShowRecording')}
      </span>
      <span className="ss-record-cur">{formatClock(slideElapsedMs)}</span>
      <span className="ss-record-total">
        {t('paneShowRehearseTotal', { time: formatClock(elapsedMs) })}
      </span>
      <button
        className="ss-record-btn"
        disabled={phase === 'starting'}
        onClick={onTogglePause}
        data-tip={t(phase === 'paused' ? 'paneShowRecordResume' : 'paneShowRecordPause')}
      >
        {t(phase === 'paused' ? 'paneShowRecordResume' : 'paneShowRecordPause')}
      </button>
    </div>
  )
}
