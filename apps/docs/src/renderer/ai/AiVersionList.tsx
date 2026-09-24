import { useI18n } from '../i18n/locale'
import { IconClock } from '../components/icons'
import { canRestore, type DocVersion } from './version-history'

export interface AiVersionListProps {
  versions: DocVersion[]
  busy: boolean
  onRollback: (id: number) => void
  onUndo: (id: number) => void
}

/**
 * The document's evolution as a list of AI turns (#293), each one restorable
 * (#543). Same shape as the Markdown / HTML snapshot footers; only the button
 * semantics differ, because here a roll back can be taken back.
 *
 * Discarded versions stay listed — they are part of the record — but lose their
 * button, since the future they describe is no longer reachable while an earlier
 * roll back stands. Versions whose stored snapshot is gone say so instead of
 * quietly losing their action.
 */
export function AiVersionList({
  versions,
  busy,
  onRollback,
  onUndo,
}: AiVersionListProps): React.ReactElement | null {
  const { t } = useI18n()
  if (versions.length === 0) return null
  return (
    <div className="ai-versions">
      <div className="ai-versions-title">
        <IconClock size={12} />
        {t('aiSnapshotsTitle')}
      </div>
      {versions.map((v) => (
        <div
          key={v.id}
          className={
            'ai-version-row' +
            (v.discarded ? ' ai-version-discarded' : '') +
            (v.rolledBack ? ' ai-version-rolled-back' : '') +
            (canRestore(v) ? '' : ' ai-version-expired')
          }
        >
          <span className="ai-version-label" data-tip={v.label}>
            <span className="ai-version-time">{v.time}</span>
            {v.label}
          </span>
          {v.rolledBack ? (
            <button
              type="button"
              className="ai-version-rollback"
              disabled={busy}
              onClick={() => onUndo(v.id)}
            >
              {t('aiRollbackUndo')}
            </button>
          ) : !canRestore(v) ? (
            <span className="ai-version-note">{t('aiVersionExpired')}</span>
          ) : (
            <button
              type="button"
              className="ai-version-rollback"
              disabled={busy}
              onClick={() => onRollback(v.id)}
            >
              {t('aiRollback')}
            </button>
          )}
        </div>
      ))}
    </div>
  )
}
