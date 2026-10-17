import { useCallback, useEffect, useState } from 'react'
import type { ReactElement } from 'react'
import { useI18n } from '../i18n/locale'
import type { OfflineStatus, OutboxRecord, OutboxSyncApi } from './OutboxApi'
import { OfflineIndicator } from './OfflineIndicator'
import './outbox.css'

/**
 * Outbox panel (Lane O owns). Lists the account outbox newest-first with
 * the five spec states, manual retry/cancel affordances, a sync-now
 * trigger, and the offline indicator. Refresh is event-driven (sync-state
 * events, offline flips) plus post-action reloads — the renderer runs no
 * retry timers (spec section 5.3: one main-process timer per account).
 * `lastError` is main-side sanitized and shown verbatim, never logged.
 */
export interface OutboxPanelProps {
  api: OutboxSyncApi
  accountId: string
  folderId?: string
}

function isRelevantAccount(eventAccountId: string, accountId: string): boolean {
  return eventAccountId === accountId
}

export function OutboxPanel({ api, accountId, folderId }: OutboxPanelProps): ReactElement {
  const { t } = useI18n()
  const [records, setRecords] = useState<OutboxRecord[]>([])
  const [loading, setLoading] = useState(true)
  const [failed, setFailed] = useState(false)
  const [syncing, setSyncing] = useState(false)
  const [pendingId, setPendingId] = useState<string | null>(null)

  const reload = useCallback(async () => {
    setLoading(true)
    try {
      setRecords(await api.outboxList(accountId))
      setFailed(false)
    } catch {
      setFailed(true)
    } finally {
      setLoading(false)
    }
  }, [api, accountId])

  useEffect(() => {
    let cancelled = false
    void (async () => {
      if (cancelled) return
      await reload()
    })()
    const offSync = api.onSyncEvent((event) => {
      if (!isRelevantAccount(event.accountId, accountId)) return
      if (event.kind === 'sync-state') setSyncing(event.state === 'syncing')
      void reload()
    })
    const offOffline = api.onOfflineStatus((status: OfflineStatus) => {
      if (status.accountId !== undefined && !isRelevantAccount(status.accountId, accountId)) return
      void reload()
    })
    return () => {
      cancelled = true
      offSync()
      offOffline()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api, accountId])

  const runSyncNow = useCallback(() => {
    void api
      .syncNow(accountId, folderId)
      .then(() => reload())
      .catch(() => setFailed(true))
  }, [api, accountId, folderId, reload])

  const runRetry = useCallback(
    (outboxId: string) => {
      setPendingId(outboxId)
      void api
        .outboxRetry(accountId, outboxId)
        .then(() => reload())
        .catch(() => setFailed(true))
        .finally(() => setPendingId(null))
    },

    [api, accountId, reload],
  )

  const runCancel = useCallback(
    (outboxId: string) => {
      setPendingId(outboxId)
      void api
        .outboxCancel(accountId, outboxId)
        .then(() => reload())
        .catch(() => setFailed(true))
        .finally(() => setPendingId(null))
    },

    [api, accountId, reload],
  )

  return (
    <section className="mail-outbox-panel" aria-label={t('mail.outbox.title')}>
      <div className="mail-outbox-head">
        <h2 className="mail-outbox-title">{t('mail.outbox.title')}</h2>
        <OfflineIndicator api={api} accountId={accountId} />
        <button type="button" className="mail-outbox-sync" disabled={syncing} onClick={runSyncNow}>
          {syncing ? t('mail.outbox.syncing') : t('mail.outbox.syncNow')}
        </button>
      </div>
      {failed && (
        <div className="mail-outbox-load-error" role="alert">
          {t('mail.errors.ipcRejected')}
        </div>
      )}
      {loading ? (
        <div className="mail-outbox-empty">{t('mail.list.loading')}</div>
      ) : records.length === 0 ? (
        <div className="mail-outbox-empty">{t('mail.outbox.empty')}</div>
      ) : (
        <ul className="mail-outbox-list">
          {records.map((record) => (
            <OutboxRow
              key={record.id}
              record={record}
              busy={pendingId === record.id}
              onRetry={() => runRetry(record.id)}
              onCancel={() => runCancel(record.id)}
            />
          ))}
        </ul>
      )}
    </section>
  )
}

function OutboxRow({
  record,
  busy,
  onRetry,
  onCancel,
}: {
  record: OutboxRecord
  busy: boolean
  onRetry(): void
  onCancel(): void
}): ReactElement {
  const { t } = useI18n()
  const subject = record.req.subject.trim()
  const recipientCount = record.req.to.length + (record.req.cc?.length ?? 0)
  return (
    <li className={`mail-outbox-item is-${record.state}`}>
      <span className={`mail-outbox-badge is-${record.state}`}>
        {t(`mail.outbox.state.${record.state}`)}
      </span>
      <div className="mail-outbox-main">
        <div className="mail-outbox-subject">
          {subject.length > 0 ? subject : t('mail.outbox.noSubject')}
        </div>
        <div className="mail-outbox-meta">
          {t('mail.outbox.recipients', { count: recipientCount })}
        </div>
        {record.state === 'failed' && (
          <div className="mail-outbox-failed-note">{t('mail.outbox.failedNote')}</div>
        )}
        {record.state === 'failed' && record.lastError && (
          <div className="mail-outbox-error" role="alert">
            {record.lastError}
          </div>
        )}
      </div>
      <div className="mail-outbox-actions">
        {record.state === 'failed' && (
          <button type="button" disabled={busy} onClick={onRetry}>
            {t('mail.outbox.retry')}
          </button>
        )}
        {(record.state === 'queued' || record.state === 'failed') && (
          <button type="button" disabled={busy} onClick={onCancel}>
            {t('mail.outbox.cancel')}
          </button>
        )}
      </div>
    </li>
  )
}
