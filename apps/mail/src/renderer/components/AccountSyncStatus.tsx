import { useEffect, useState } from 'react'
import type { ReactElement } from 'react'
import { useI18n } from '../i18n/locale'
import type { OutboxSyncApi } from './OutboxApi'
import './outbox.css'

/**
 * Per-account sync status + send caps (Lane O owns). Sync state comes from
 * the existing `sync-state`/`resync`/`conflict` mail events filtered by
 * account; nothing here reads bodies, secrets, or staged paths.
 *
 * Caps are restated P1 driver ceilings (the renderer cannot import the
 * node-only driver modules — ComposeHost precedent): IMAP
 * `MAX_TOTAL_BYTES`/`MAX_ATTACHMENT_BYTES`/`MAX_RECIPIENTS` from
 * `imap-smtp-driver.ts`, Gmail `GMAIL_MAX_MESSAGE_BYTES`/
 * `GMAIL_MAX_RECIPIENTS_PER_MESSAGE` from `gmail.ts` (Gmail enforces only
 * the per-message ceiling, so the per-file figure is that same ceiling,
 * effective), Graph `GRAPH_MAX_TOTAL_BYTES`/`GRAPH_MAX_ATTACHMENT_BYTES`/
 * `GRAPH_MAX_RECIPIENTS` from `graph-driver.ts`. Main re-enforces the real
 * caps at enqueue and at send; this display is informational only.
 */
export type AccountDriver = 'imap' | 'gmail' | 'graph'

export interface AccountSyncStatusProps {
  api: OutboxSyncApi
  accountId: string
  label: string
  driver?: AccountDriver
}

interface DriverCaps {
  name: string
  totalBytes: number
  fileBytes: number
  recipients: number
}

const MIB = 1024 * 1024

export const SEND_CAPS: Record<AccountDriver, DriverCaps> = {
  imap: { name: 'IMAP', totalBytes: 50 * MIB, fileBytes: 25 * MIB, recipients: 100 },
  gmail: { name: 'Gmail', totalBytes: 35 * MIB, fileBytes: 35 * MIB, recipients: 500 },
  graph: { name: 'Microsoft Graph', totalBytes: 4 * MIB, fileBytes: 3 * MIB, recipients: 500 },
}

function formatMegaBytes(bytes: number): string {
  const whole = bytes / MIB
  return `${Number.isInteger(whole) ? String(whole) : whole.toFixed(1)} MB`
}

type SyncNote = { kind: 'resync' } | { kind: 'conflict' } | { kind: 'error'; detail?: string }

export function AccountSyncStatus({
  api,
  accountId,
  label,
  driver = 'imap',
}: AccountSyncStatusProps): ReactElement {
  const { t } = useI18n()
  const [syncing, setSyncing] = useState(false)
  const [note, setNote] = useState<SyncNote | null>(null)
  const caps = SEND_CAPS[driver]

  useEffect(
    () =>
      api.onSyncEvent((event) => {
        if (event.accountId !== accountId) return
        if (event.kind === 'sync-state') {
          setSyncing(event.state === 'syncing')
          setNote(event.state === 'error' ? { kind: 'error', detail: event.detail } : null)
        } else if (event.kind === 'resync') {
          setSyncing(false)
          setNote({ kind: 'resync' })
        } else if (event.kind === 'conflict') {
          setNote({ kind: 'conflict' })
        }
      }),
    [api, accountId],
  )

  return (
    <section className="mail-sync-status" aria-label={label}>
      <div className="mail-sync-state">
        {syncing ? (
          <span className="mail-sync-badge is-syncing">{t('mail.outbox.syncing')}</span>
        ) : (
          <span className="mail-sync-badge is-idle">{t('mail.sync.upToDate')}</span>
        )}
      </div>
      {note && (
        <div className="mail-sync-note" role="status">
          {note.kind === 'resync' && t('mail.errors.resyncNotice')}
          {note.kind === 'conflict' && t('mail.errors.conflictNotice')}
          {note.kind === 'error' && (note.detail ?? t('mail.errors.syncError'))}
        </div>
      )}
      <div className="mail-sync-caps">
        <span className="mail-sync-caps-title">{t('mail.sync.capsTitle')}</span>
        <span className="mail-sync-caps-detail">
          {t('mail.sync.capsDetail', {
            driver: caps.name,
            total: formatMegaBytes(caps.totalBytes),
            file: formatMegaBytes(caps.fileBytes),
            recipients: caps.recipients,
          })}
        </span>
      </div>
    </section>
  )
}
