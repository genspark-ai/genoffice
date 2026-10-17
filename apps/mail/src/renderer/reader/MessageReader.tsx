import { useEffect, useMemo, useState } from 'react'
import type { MouseEvent, ReactElement } from 'react'
import type { MailBody, MailEnvelope } from '@genoffice/mail-protocols'
import type { MailApi } from '../../shared/ipc'
import { useI18n } from '../i18n/locale'
import { extractDomain, lookalikeDomain, replyToMismatch, sanitizeMailHtml } from './quarantine'
import { AttachmentVerdicts } from './AttachmentVerdicts'
import type { ReaderAttachmentRef } from './AttachmentVerdicts'
import { allowSender, isSenderAllowed } from './allowlist'
import '../mail-quarantine.css'
import './reader.css'

/**
 * Quarantined message reader (Lane R owns). Bodies render through
 * `sanitizeMailHtml` into `.mq-body` — authored colors, never chrome
 * tokens. Remote images stay blocked until the per-message "Show
 * images" action or an explicit per-sender allowlist entry. Links
 * outside an allowlisted sender need an inline confirm; phishing
 * signals (Reply-To divergence, lookalike domain, auth failure)
 * surface as reader-chrome flags, never inside the body layer.
 */

export interface ReaderSecurity {
  replyTo?: string
  auth?: 'pass' | 'fail' | 'unknown'
}

export interface MessageReaderProps {
  api: MailApi
  envelope: MailEnvelope | null
  security?: ReaderSecurity
  trustedDomains?: string[]
  onLinkOpen?: (url: string) => void
  /**
   * Receive-side quarantine verdicts (H5): when present, each inbound
   * attachment renders its staged-temp --> scan --> store verdict row
   * (preview for clean, confirm-gated save+open for quarantined, a
   * non-rendered placeholder for blocked). Absent, the header keeps the
   * plain attachments label.
   */
  attachments?: ReaderAttachmentRef[]
}

function escapeToHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

export function MessageReader({
  api,
  envelope,
  security,
  trustedDomains = [],
  onLinkOpen,
  attachments,
}: MessageReaderProps): ReactElement {
  const { t } = useI18n()
  const [body, setBody] = useState<MailBody | null>(null)
  const [loading, setLoading] = useState(false)
  const [failed, setFailed] = useState(false)
  const [showImagesOnce, setShowImagesOnce] = useState(false)
  const [senderAllowed, setSenderAllowed] = useState(false)
  const [linkConfirm, setLinkConfirm] = useState<string | null>(null)

  const envelopeRef = envelope?.ref ?? null

  useEffect(() => {
    setBody(null)
    setFailed(false)
    setShowImagesOnce(false)
    setLinkConfirm(null)
    if (!envelope || !envelopeRef) {
      setLoading(false)
      return
    }
    setSenderAllowed(isSenderAllowed(envelope.from[0] ?? ''))
    setLoading(true)
    let cancelled = false
    api
      .fetchBody(envelopeRef)
      .then((next) => {
        if (!cancelled) {
          setBody(next)
          setLoading(false)
        }
      })
      .catch(() => {
        if (!cancelled) {
          setFailed(true)
          setLoading(false)
        }
      })
    return () => {
      cancelled = true
    }
  }, [api, envelope, envelopeRef])

  const sanitized = useMemo(() => {
    if (!body) return { html: '', blockedCount: 0 }
    return sanitizeMailHtml(bodyHtml(body), {
      allowRemoteImages: showImagesOnce || senderAllowed,
      blockedLabel: t('mail.reader.blockedImagesNote'),
    })
    // t is not referentially stable; translate at render time only.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [body, showImagesOnce, senderAllowed])

  if (!envelope) {
    return (
      <div className="mail-reader mail-reader-empty" aria-live="polite">
        {t('mail.openMessage')}
      </div>
    )
  }

  const imagesVisible = showImagesOnce || senderAllowed
  const showImageBar = sanitized.blockedCount > 0 && !imagesVisible
  const flags: string[] = []
  if (replyToMismatch(envelope.from, security?.replyTo))
    flags.push(t('mail.reader.replyToMismatch'))
  const senderDomain = extractDomain(envelope.from[0] ?? '')
  if (senderDomain && lookalikeDomain(senderDomain, trustedDomains)) {
    flags.push(t('mail.reader.lookalikeDomain'))
  }
  if (security?.auth === 'fail') flags.push(t('mail.reader.authFail'))

  const onBodyClick = (event: MouseEvent<HTMLDivElement>): void => {
    const anchor = (event.target as HTMLElement).closest?.('a[href]')
    if (!anchor) return
    const href = anchor.getAttribute('href') ?? ''
    if (!/^https?:\/\//i.test(href)) {
      event.preventDefault()
      return
    }
    if (senderAllowed) {
      if (onLinkOpen) {
        event.preventDefault()
        onLinkOpen(href)
      }
      return
    }
    event.preventDefault()
    setLinkConfirm(href)
  }

  const openConfirmedLink = (): void => {
    if (!linkConfirm) return
    if (onLinkOpen) onLinkOpen(linkConfirm)
    else window.open(linkConfirm, '_blank', 'noopener,noreferrer')
    setLinkConfirm(null)
  }

  const trustSender = (): void => {
    allowSender(envelope.from[0] ?? '')
    setSenderAllowed(true)
  }

  return (
    <article className="mail-reader" aria-label={envelope.subject}>
      <header className="mail-reader-head">
        <h2 className="mail-reader-subject">{envelope.subject}</h2>
        <dl className="mail-reader-meta">
          <div>
            <dt>{t('mail.reader.from')}</dt>
            <dd>
              <span className="mail-reader-addr">{envelope.from.join(', ')}</span>
            </dd>
          </div>
          <div>
            <dt>{t('mail.reader.to')}</dt>
            <dd>{envelope.to.join(', ')}</dd>
          </div>
          {envelope.cc.length > 0 && (
            <div>
              <dt>{t('mail.reader.cc')}</dt>
              <dd>{envelope.cc.join(', ')}</dd>
            </div>
          )}
          <div>
            <dt>{t('mail.reader.date')}</dt>
            <dd>{envelope.date}</dd>
          </div>
        </dl>
        {flags.length > 0 && (
          <ul className="mail-reader-flags">
            {flags.map((flag) => (
              <li key={flag} className="mail-reader-flag" role="status">
                {flag}
              </li>
            ))}
          </ul>
        )}
        {envelope.hasAttach && (
          <div className="mail-reader-attachments">
            <span>{t('mail.reader.attachments')}</span>
            {attachments !== undefined && attachments.length > 0 ? (
              <AttachmentVerdicts attachments={attachments} />
            ) : null}
          </div>
        )}
      </header>

      {showImageBar && (
        <div className="mail-reader-imagebar" role="status">
          <span>{t('mail.reader.blockedImagesNote')}</span>
          <button type="button" onClick={() => setShowImagesOnce(true)}>
            {t('mail.reader.showImages')}
          </button>
          <button type="button" onClick={trustSender}>
            {t('mail.reader.alwaysAllowSender')}
          </button>
        </div>
      )}

      {linkConfirm && (
        <div
          className="mail-reader-linkconfirm"
          role="alertdialog"
          aria-label={t('mail.reader.linkConfirmTitle')}
        >
          <p className="mail-reader-linkconfirm-title">{t('mail.reader.linkConfirmTitle')}</p>
          <p>{t('mail.reader.linkConfirmMsg')}</p>
          <p className="mail-reader-linkconfirm-url">{linkConfirm}</p>
          <button type="button" onClick={openConfirmedLink}>
            {t('mail.reader.linkConfirmTitle')}
          </button>
          <button type="button" onClick={() => setLinkConfirm(null)}>
            {t('mail.btnCancel')}
          </button>
        </div>
      )}

      <div className="mail-reader-body">
        {loading && <p className="mail-reader-status">{t('mail.list.loading')}</p>}
        {failed && !loading && (
          <p className="mail-reader-status" role="alert">
            {t('mail.errors.ipcRejected')}
          </p>
        )}
        {!loading && !failed && body && (
          <div
            className="mq-body"
            onClickCapture={onBodyClick}
            dangerouslySetInnerHTML={{ __html: sanitized.html }}
          />
        )}
      </div>
    </article>
  )
}

/**
 * HTML source for a body. P3 fakes inline message HTML on the body
 * object (fake-only `html` extension, never on the wire); production
 * resolves `htmlRef` through main-process handlers in a later phase,
 * so the fallback renders the inline text part only.
 */
function bodyHtml(body: MailBody): string {
  const fakeHtml = (body as Partial<{ html: string }>).html
  if (typeof fakeHtml === 'string' && fakeHtml) return fakeHtml
  return body.text ? `<pre class="mq-text">${escapeToHtml(body.text)}</pre>` : ''
}
