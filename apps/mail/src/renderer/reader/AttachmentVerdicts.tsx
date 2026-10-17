import { useState } from 'react'
import type { ReactElement } from 'react'
import { useI18n } from '../i18n/locale'
import type { QuarantineReason, QuarantineVerdict } from '../../shared/quarantine-channels'

/**
 * Receive-side attachment verdict rows (H5 owns). Every inbound attachment
 * enters the main-process staged-temp --> scan --> store pipeline before
 * the reader sees it; these rows surface the verdict:
 * - clean: inline text preview allowed (via the file-parse path in main).
 * - quarantined (executables, scripts, macros, MIME mismatch): NEVER
 *   inline-previewed — save + OS-open only with an explicit confirm.
 * - blocked: non-rendered placeholder — no preview, no open control.
 * Without the quarantine bridge (old preloads, unit tests) the rows still
 * render as informative placeholders with no actions.
 */
export interface ReaderAttachmentRef {
  filename: string
  mime: string
  size: number
  verdict: QuarantineVerdict
  reason?: QuarantineReason
  /** Store path for preview/open; absent for blocked placeholders. */
  path?: string
}

function formatBytes(size: number): string {
  if (size < 1024) return `${size} B`
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`
  return `${(size / (1024 * 1024)).toFixed(1)} MB`
}

export function AttachmentVerdicts({
  attachments,
}: {
  attachments: ReaderAttachmentRef[]
}): ReactElement {
  const { t } = useI18n()
  const [previews, setPreviews] = useState<Record<string, string>>({})
  const [confirmPath, setConfirmPath] = useState<string | null>(null)
  const bridge = typeof window !== 'undefined' ? window.mailQuarantine : undefined

  async function showPreview(att: ReaderAttachmentRef): Promise<void> {
    if (!bridge || !att.path) return
    try {
      const result = await bridge.preview(att.path)
      if (result.allowed && typeof result.text === 'string') {
        setPreviews((current) => ({ ...current, [att.path as string]: result.text as string }))
      }
    } catch {
      // Preview refusal stays a quiet placeholder; main already logged it.
    }
  }

  async function openStored(att: ReaderAttachmentRef): Promise<void> {
    if (!bridge || !att.path) return
    if (confirmPath !== att.path) {
      setConfirmPath(att.path)
      return
    }
    setConfirmPath(null)
    try {
      await bridge.openAttachment(att.path, true)
    } catch {
      // Rejection (unwired opener, missing file) stays on the row.
    }
  }

  return (
    <ul className="mail-reader-flags">
      {attachments.map((att) => {
        const key = `${att.filename}-${att.size}`
        const shown = att.path !== undefined ? previews[att.path] : undefined
        return (
          <li key={key} className="mail-reader-flag">
            <span className="mail-reader-addr">
              {att.filename} ({formatBytes(att.size)})
            </span>{' '}
            {att.verdict === 'blocked' ? (
              <span role="status">
                {t('mail.quarantine.blocked', { reason: att.reason ?? 'suspicious-content' })}
              </span>
            ) : null}
            {att.verdict === 'quarantined' ? (
              <span role="status">
                {t('mail.quarantine.quarantined')} {t('mail.quarantine.execNote')}
              </span>
            ) : null}
            {att.verdict === 'clean' ? (
              <span role="status">{t('mail.quarantine.clean')}</span>
            ) : null}{' '}
            {att.verdict === 'clean' && bridge && att.path ? (
              <button type="button" onClick={() => void showPreview(att)}>
                {t('mail.quarantine.preview')}
              </button>
            ) : null}
            {att.verdict === 'quarantined' && bridge && att.path ? (
              <button type="button" onClick={() => void openStored(att)}>
                {confirmPath === att.path
                  ? t('mail.quarantine.confirmTitle')
                  : t('mail.quarantine.saveOpen')}
              </button>
            ) : null}
            {shown !== undefined ? <pre className="mq-text">{shown}</pre> : null}
          </li>
        )
      })}
    </ul>
  )
}
