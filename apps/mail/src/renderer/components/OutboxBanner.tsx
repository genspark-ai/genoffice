import { useCallback, useEffect, useState } from 'react'
import type { ReactElement } from 'react'
import { useI18n } from '../i18n/locale'
import type { OfflineStatus, OutboxSyncApi } from './OutboxApi'
import './outbox.css'

/**
 * Outbox banner (Lane O owns). This EXTENDS the P3 outbox banner — the same
 * `.mail-notice` root App.tsx already renders for `mail.compose.outboxQueued`
 * — it does not fork a second banner. Drop-in replacement for the App.tsx
 * block:
 *
 *   {notice && <div className="mail-notice" role="status">{notice}</div>}
 *
 * becomes:
 *
 *   <OutboxBanner api={resolvedApi} accountId={selection?.accountId ?? ''}
 *     notice={notice} onOpenPanel={() => setOutboxOpen(true)} />
 *
 * The banner keeps the P3 notice text, adds live queued/failed counts plus
 * the offline chip, and renders nothing when there is no notice, no outbox
 * rows, and the account is online (same empty behavior as P3).
 */
export interface OutboxBannerProps {
  api: OutboxSyncApi
  accountId: string
  notice: string | null
  onOpenPanel(): void
}

export function OutboxBanner({
  api,
  accountId,
  notice,
  onOpenPanel,
}: OutboxBannerProps): ReactElement | null {
  const { t } = useI18n()
  const [queued, setQueued] = useState(0)
  const [failedCount, setFailedCount] = useState(0)
  const [offline, setOffline] = useState(false)

  const reload = useCallback(async () => {
    try {
      const records = await api.outboxList(accountId)
      setQueued(
        records.filter((item) => item.state === 'queued' || item.state === 'sending').length,
      )
      setFailedCount(records.filter((item) => item.state === 'failed').length)
    } catch {
      // Banner stays quiet on IPC rejection; the panel surfaces the error.
    }
  }, [api, accountId])

  useEffect(() => {
    void reload()
    const offSync = api.onSyncEvent((event) => {
      if (event.accountId !== accountId) return
      void reload()
    })
    const offOffline = api.onOfflineStatus((status: OfflineStatus) => {
      if (status.accountId !== undefined && status.accountId !== accountId) return
      setOffline(!status.online)
      void reload()
    })
    return () => {
      offSync()
      offOffline()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api, accountId])

  const total = queued + failedCount
  if (notice === null && total === 0 && !offline) return null

  return (
    <div className="mail-notice mail-outbox-banner" role="status">
      {notice && <span className="mail-outbox-banner-notice">{notice}</span>}
      {offline && <span className="mail-outbox-banner-offline">{t('mail.offline.offline')}</span>}
      {total > 0 && (
        <button type="button" className="mail-outbox-banner-open" onClick={onOpenPanel}>
          {t('mail.outbox.viewOutbox', { count: total })}
        </button>
      )}
    </div>
  )
}
