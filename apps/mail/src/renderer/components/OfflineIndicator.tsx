import { useEffect, useState } from 'react'
import type { ReactElement } from 'react'
import { useI18n } from '../i18n/locale'
import type { OfflineStatus, OutboxSyncApi } from './OutboxApi'
import './outbox.css'

/**
 * Offline indicator (Lane O owns). Subscribes via `onOfflineStatus` (the
 * bridge replays the current value on subscribe); renders nothing until the
 * first status arrives so no wrong affordance flashes. When `accountId` is
 * set, global events (no accountId) still apply while other accounts'
 * events are ignored. Status text only — no secrets or bodies involved.
 */
export interface OfflineIndicatorProps {
  api: OutboxSyncApi
  accountId?: string
}

export function OfflineIndicator({ api, accountId }: OfflineIndicatorProps): ReactElement | null {
  const { t } = useI18n()
  const [status, setStatus] = useState<OfflineStatus | null>(null)

  useEffect(
    () =>
      api.onOfflineStatus((next) => {
        if (accountId !== undefined && next.accountId !== undefined && next.accountId !== accountId)
          return
        setStatus(next)
      }),
    [api, accountId],
  )

  if (status === null) return null
  const online = status.online
  return (
    <span
      className={online ? 'mail-offline-chip is-online' : 'mail-offline-chip is-offline'}
      role="status"
    >
      <span className="mail-offline-dot" aria-hidden="true" />
      {online ? t('mail.offline.online') : t('mail.offline.offline')}
    </span>
  )
}
