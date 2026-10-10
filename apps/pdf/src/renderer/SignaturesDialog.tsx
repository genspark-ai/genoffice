import type { ReactElement } from 'react'
import type { CertificateInfo, PdfSignatureInfo, PdfSignatureProblem } from '../shared/ipc'
import type { TFunc } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

type Key = Parameters<TFunc>[0]

const PROBLEM_KEYS: Record<PdfSignatureProblem, Key> = {
  'digest-mismatch': 'sigProblemDigestMismatch',
  'signature-invalid': 'sigProblemSignatureInvalid',
  malformed: 'sigProblemMalformed',
  'unsupported-algorithm': 'sigProblemUnsupported',
  'signer-certificate-missing': 'sigProblemNoSignerCert',
  'document-changed-after-signing': 'sigProblemChanged',
  'self-signed': 'sigProblemSelfSigned',
  'untrusted-issuer': 'sigProblemUntrusted',
  'certificate-expired': 'sigProblemExpired',
  'certificate-not-yet-valid': 'sigProblemNotYetValid',
}

const STATUS_KEYS: Record<PdfSignatureInfo['status'], Key> = {
  valid: 'sigStatusValid',
  warning: 'sigStatusWarning',
  invalid: 'sigStatusInvalid',
}

const LEVEL_KEYS = { 1: 'certLevel1', 2: 'certLevel2', 3: 'certLevel3' } as const

/** Worst status wins: what the banner and the ribbon button summarize */
export function overallSignatureStatus(
  signatures: readonly PdfSignatureInfo[],
): PdfSignatureInfo['status'] {
  if (signatures.some((s) => s.status === 'invalid')) return 'invalid'
  if (signatures.some((s) => s.status === 'warning')) return 'warning'
  return 'valid'
}

const formatTime = (iso: string): string => new Date(iso).toLocaleString()

function CertificateRows({ cert, t }: { cert: CertificateInfo; t: TFunc }): ReactElement {
  const row = (label: string, value: string): ReactElement => (
    <div className="pdf-prop-row">
      <span>{label}</span>
      <em data-tip={value}>{value || '—'}</em>
    </div>
  )
  return (
    <div className="pdf-prop-table">
      {row(t('sigSubject'), cert.commonName)}
      {row(t('sigIssuedBy'), cert.selfSigned ? t('sigSelfSigned') : cert.issuerCommonName)}
      {row(t('sigValidFrom'), formatTime(cert.validFrom))}
      {row(t('sigValidTo'), formatTime(cert.validTo))}
      {row(t('sigKey'), cert.keyAlgorithm)}
      {row(t('sigSerial'), cert.serialNumber)}
      {row(t('sigFingerprint'), cert.fingerprint)}
    </div>
  )
}

function SignatureCard({
  sig,
  t,
  onGoToPage,
}: {
  sig: PdfSignatureInfo
  t: TFunc
  onGoToPage: (pageIndex: number) => void
}): ReactElement {
  const name = sig.signer?.commonName ?? sig.signerName ?? sig.fieldName
  const detail = (label: string, value?: string): ReactElement | null =>
    value ? (
      <div className="pdf-prop-row">
        <span>{label}</span>
        <em data-tip={value}>{value}</em>
      </div>
    ) : null

  return (
    <div className={`pdf-sig-card ${sig.status}`}>
      <div className="pdf-sig-head">
        <span className={`pdf-sig-pill ${sig.status}`}>{t(STATUS_KEYS[sig.status])}</span>
        <strong data-tip={name}>
          {sig.kind === 'timestamp' ? t('sigKindTimestamp') : t('sigSignedBy', { name })}
        </strong>
        {sig.pageIndex >= 0 && (
          <button className="pdf-modal-btn" onClick={() => onGoToPage(sig.pageIndex)}>
            {t('sigOnPage', { page: sig.pageIndex + 1 })} · {t('sigGoToPage')}
          </button>
        )}
      </div>
      {sig.problems.length > 0 && (
        <ul className="pdf-sig-problems">
          {sig.problems.map((problem) => (
            <li key={problem}>{t(PROBLEM_KEYS[problem])}</li>
          ))}
        </ul>
      )}
      <div className="pdf-prop-table">
        {detail(
          sig.signingTimeSigned ? t('sigSignedAt') : t('sigClaimedAt'),
          sig.signingTime ? formatTime(sig.signingTime) : undefined,
        )}
        {detail(t('sigReason'), sig.reason)}
        {detail(t('sigLocation'), sig.location)}
        {detail(t('sigContact'), sig.contactInfo)}
        {detail(t('sigHash'), sig.digestAlgorithm)}
        {detail(t('sigCovers'), sig.coversWholeDocument ? t('sigCoversAll') : t('sigCoversPart'))}
        {sig.certifiedLevel &&
          detail(t('certSign'), t('sigCertified', { level: t(LEVEL_KEYS[sig.certifiedLevel]) }))}
      </div>
      {sig.chain.length > 0 && (
        <details className="pdf-sig-chain">
          <summary>
            {t('sigChain')} · {sig.trusted ? t('sigChainTrusted') : t('sigChainUntrusted')}
          </summary>
          {sig.chain.map((cert) => (
            <CertificateRows key={cert.fingerprint} cert={cert} t={t} />
          ))}
        </details>
      )}
    </div>
  )
}

/** Lists every certificate signature of the open PDF with the verdict and the signer's certificate */
export function SignaturesDialog({
  t,
  signatures,
  onClose,
  onGoToPage,
}: {
  t: TFunc
  signatures: readonly PdfSignatureInfo[]
  onClose: () => void
  onGoToPage: (pageIndex: number) => void
}): ReactElement {
  const dialogRef = useModalDialog(onClose)
  return (
    <div className="pdf-modal-mask" onClick={onClose}>
      <div
        ref={dialogRef}
        className="pdf-modal pdf-modal-wide pdf-sig-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('sigPanelTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="pdf-modal-title">{t('sigPanelTitle')}</div>
        {signatures.map((sig) => (
          <SignatureCard
            key={sig.fieldName}
            sig={sig}
            t={t}
            onGoToPage={(index) => {
              onGoToPage(index)
              onClose()
            }}
          />
        ))}
        <div className="pdf-modal-hint">{t('sigNoRevocation')}</div>
        <div className="pdf-modal-actions">
          <button className="pdf-modal-btn primary" onClick={onClose}>
            {t('ok')}
          </button>
        </div>
      </div>
    </div>
  )
}
