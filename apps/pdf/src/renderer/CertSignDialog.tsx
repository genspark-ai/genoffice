import { useCallback, useEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import type {
  CertificateInfo,
  SignWithCertificateResult,
  SignaturePosition,
  SystemCertificateInfo,
  SystemStoreIssue,
} from '../shared/ipc'
import type { TFunc } from './i18n/locale'
import { useModalDialog } from './modal-dialog'
import { renderSignatureAppearance } from './cert-appearance'

const BOX_WIDTH = 210
const BOX_HEIGHT = 70
const POSITIONS: { value: SignaturePosition; key: Parameters<TFunc>[0] }[] = [
  { value: 'bottom-right', key: 'certPosBottomRight' },
  { value: 'bottom-left', key: 'certPosBottomLeft' },
  { value: 'top-right', key: 'certPosTopRight' },
  { value: 'top-left', key: 'certPosTopLeft' },
  { value: 'center', key: 'certPosCenter' },
]

type FailedResult = Extract<SignWithCertificateResult, { ok: false; cancelled?: false }>

export function certErrorKey(error: FailedResult['error']): Parameters<TFunc>[0] {
  switch (error) {
    case 'cert-password':
      return 'certErrPassword'
    case 'cert-invalid':
      return 'certErrInvalid'
    case 'cert-unsupported-key':
      return 'certErrUnsupported'
    case 'cert-no-key':
      return 'certErrNoKey'
    case 'cert-not-exportable':
      return 'certErrNotExportable'
    case 'cert-store-locked':
      return 'certErrStoreLocked'
    case 'cert-store-tool-missing':
      return 'certErrNoTools'
    case 'pdf-encrypted':
      return 'certErrEncrypted'
    default:
      return 'certErrFailed'
  }
}

const formatDate = (iso: string): string => new Date(iso).toLocaleDateString()

/** Sign the open PDF with a PKCS#12 certificate; the result is always a new file */
export function CertSignDialog({
  t,
  filePath,
  pageCount,
  currentPage,
  hasSignatures,
  onCancel,
  onSigned,
}: {
  t: TFunc
  filePath: string
  pageCount: number
  currentPage: number
  /** The document already carries signatures: it can still be signed, but not certified */
  hasSignatures: boolean
  onCancel: () => void
  onSigned: (path: string) => void
}): ReactElement {
  const dialogRef = useModalDialog(onCancel)
  const [source, setSource] = useState<'system' | 'file'>('system')
  const [cert, setCert] = useState<{ certId: string; fileName: string } | null>(null)
  const [systemCerts, setSystemCerts] = useState<SystemCertificateInfo[]>([])
  const [systemId, setSystemId] = useState('')
  const [issues, setIssues] = useState<SystemStoreIssue[]>([])
  const [storePassword, setStorePassword] = useState('')
  const [listing, setListing] = useState(false)
  const activeId = source === 'file' ? (cert?.certId ?? '') : systemId
  const activeIdRef = useRef('')
  activeIdRef.current = activeId
  const [password, setPassword] = useState('')
  const [signer, setSigner] = useState<CertificateInfo | null>(null)
  const [chain, setChain] = useState<CertificateInfo[]>([])
  const [checking, setChecking] = useState(false)
  const [reason, setReason] = useState('')
  const [location, setLocation] = useState('')
  const [contact, setContact] = useState('')
  const [visible, setVisible] = useState(true)
  const [page, setPage] = useState(currentPage)
  const [position, setPosition] = useState<SignaturePosition>('bottom-right')
  const [certify, setCertify] = useState(false)
  const [level, setLevel] = useState<1 | 2 | 3>(2)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)

  // A different password or certificate invalidates what was checked
  useEffect(() => {
    setSigner(null)
    setChain([])
  }, [cert, password, source, systemId])

  const loadSystem = useCallback(async (unlockWith?: string) => {
    setListing(true)
    try {
      const result = await window.pdfApi.listSystemCertificates(unlockWith)
      setSystemCerts(result.certs)
      setIssues(result.issues)
      setSystemId((current) =>
        result.certs.some((c) => c.certId === current) ? current : (result.certs[0]?.certId ?? ''),
      )
    } catch {
      setSystemCerts([])
      setIssues([])
    } finally {
      setListing(false)
    }
  }, [])

  useEffect(() => {
    if (source === 'system') void loadSystem()
  }, [source, loadSystem])

  // Whatever the main process still holds for the chosen identity goes when the dialog does
  useEffect(
    () => () => {
      if (activeIdRef.current) window.pdfApi.releaseCertificate(activeIdRef.current)
    },
    [],
  )

  const choose = async () => {
    setError(null)
    const picked = await window.pdfApi.pickCertificate()
    if (picked) setCert(picked)
  }

  const check = async () => {
    if (!activeId || checking) return
    setChecking(true)
    setError(null)
    try {
      const result = await window.pdfApi.inspectCertificate(
        activeId,
        source === 'file' ? password : '',
        source === 'system' && storePassword ? storePassword : undefined,
      )
      if (result.ok) {
        setSigner(result.signer)
        setChain(result.chain)
      } else {
        setError(t(certErrorKey(result.error)))
      }
    } finally {
      setChecking(false)
    }
  }

  const sign = async () => {
    if (!activeId || !signer || busy) return
    setBusy(true)
    setError(null)
    try {
      const png = visible
        ? renderSignatureAppearance(
            {
              heading: t('certApprSignedBy'),
              name: signer.commonName,
              date: t('certApprDate', { date: new Date().toLocaleString() }),
              reason: reason.trim() ? t('certApprReason', { reason: reason.trim() }) : undefined,
              location: location.trim()
                ? t('certApprLocation', { location: location.trim() })
                : undefined,
            },
            BOX_WIDTH,
            BOX_HEIGHT,
          )
        : null
      const result = await window.pdfApi.signWithCertificate({
        path: filePath,
        certId: activeId,
        password: source === 'file' ? password : '',
        reason: reason.trim() || undefined,
        location: location.trim() || undefined,
        contactInfo: contact.trim() || undefined,
        certifyLevel: certify && !hasSignatures ? level : undefined,
        visible: png
          ? {
              pageIndex: Math.min(Math.max(page, 1), pageCount) - 1,
              position,
              width: BOX_WIDTH,
              height: BOX_HEIGHT,
              png,
            }
          : undefined,
      })
      if (result.ok) onSigned(result.path)
      else if (!result.cancelled) setError(t(certErrorKey(result.error)))
    } catch {
      setError(t('certErrFailed'))
    } finally {
      setBusy(false)
    }
  }

  const levelKeys = { 1: 'certLevel1', 2: 'certLevel2', 3: 'certLevel3' } as const

  return (
    <div className="pdf-modal-mask" onClick={busy ? undefined : onCancel}>
      <div
        ref={dialogRef}
        className="pdf-modal pdf-modal-wide pdf-cert-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('certSignTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="pdf-modal-title">{t('certSignTitle')}</div>
        <div className="pdf-modal-hint">{t('certSignIntro')}</div>

        <div className="pdf-field">
          <span>{t('certSource')}</span>
          <select
            className="pdf-modal-input"
            value={source}
            disabled={busy}
            onChange={(e) => setSource(e.target.value as 'system' | 'file')}
          >
            <option value="system">{t('certSourceSystem')}</option>
            <option value="file">{t('certSourceFile')}</option>
          </select>
        </div>

        {source === 'system' ? (
          <>
            <div className="pdf-field">
              <span>{t('certSigner')}</span>
              <select
                className="pdf-modal-input"
                value={systemId}
                disabled={busy || systemCerts.length === 0}
                onChange={(e) => setSystemId(e.target.value)}
              >
                {systemCerts.map((c) => (
                  <option key={c.certId} value={c.certId}>
                    {c.commonName}
                    {c.issuerCommonName && c.issuerCommonName !== c.commonName
                      ? ` — ${c.issuerCommonName}`
                      : ''}
                    {c.expired ? ` (${t('certExpiredTag')})` : ''}
                  </option>
                ))}
              </select>
              <button
                className="pdf-modal-btn"
                onClick={() => void loadSystem(storePassword || undefined)}
                disabled={busy || listing}
              >
                {t('certRefresh')}
              </button>
            </div>
            {!listing && systemCerts.length === 0 && issues.length === 0 && (
              <div className="pdf-modal-hint">{t('certSystemEmpty')}</div>
            )}
            {issues.includes('nss-tools-missing') && (
              <div className="pdf-modal-hint">{t('certStoreNoTools')}</div>
            )}
            {issues.includes('store-locked') && (
              <div className="pdf-field">
                <span>{t('certStorePassword')}</span>
                <input
                  className="pdf-modal-input"
                  type="password"
                  autoComplete="off"
                  value={storePassword}
                  onChange={(e) => setStorePassword(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') void loadSystem(storePassword || undefined)
                  }}
                />
                <button
                  className="pdf-modal-btn"
                  onClick={() => void loadSystem(storePassword || undefined)}
                  disabled={listing}
                >
                  {t('certStoreUnlock')}
                </button>
              </div>
            )}
            <div className="pdf-field">
              <span />
              <button
                className="pdf-modal-btn"
                onClick={() => void check()}
                disabled={!systemId || checking || busy}
              >
                {checking ? t('certChecking') : t('certCheck')}
              </button>
              <em className="pdf-cert-file">{t('certSystemHint')}</em>
            </div>
          </>
        ) : (
          <>
            <div className="pdf-field">
              <span>{t('certSigner')}</span>
              <button className="pdf-modal-btn" onClick={() => void choose()} disabled={busy}>
                {t('certChoose')}
              </button>
              <em className="pdf-cert-file" data-tip={cert?.fileName}>
                {cert?.fileName ?? t('certNoFile')}
              </em>
            </div>
            <label className="pdf-field">
              <span>{t('certPassword')}</span>
              <input
                className="pdf-modal-input"
                type="password"
                autoComplete="off"
                value={password}
                disabled={!cert || busy}
                onChange={(e) => setPassword(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') void (signer ? sign() : check())
                }}
              />
              <button
                className="pdf-modal-btn"
                onClick={() => void check()}
                disabled={!cert || checking || busy}
              >
                {checking ? t('certChecking') : t('certCheck')}
              </button>
            </label>
          </>
        )}

        {signer && (
          <div className="pdf-prop-table">
            <div className="pdf-prop-row">
              <span>{t('certSigner')}</span>
              <em data-tip={signer.subject}>{signer.commonName}</em>
            </div>
            <div className="pdf-prop-row">
              <span>{t('certIssuer')}</span>
              <em data-tip={signer.issuer}>
                {signer.selfSigned ? t('sigSelfSigned') : signer.issuerCommonName}
              </em>
            </div>
            <div className="pdf-prop-row">
              <span>{t('certValidity')}</span>
              <em>
                {t('certValidityRange', {
                  from: formatDate(signer.validFrom),
                  to: formatDate(signer.validTo),
                })}
              </em>
            </div>
            <div className="pdf-prop-row">
              <span>{t('sigKey')}</span>
              <em>
                {signer.keyAlgorithm}
                {chain.length > 1 ? ` · ${t('sigChain')} ${chain.length}` : ''}
              </em>
            </div>
          </div>
        )}

        <label className="pdf-field">
          <span>{t('certReason')}</span>
          <input
            className="pdf-modal-input"
            value={reason}
            maxLength={200}
            disabled={busy}
            onChange={(e) => setReason(e.target.value)}
          />
        </label>
        <label className="pdf-field">
          <span>{t('certLocation')}</span>
          <input
            className="pdf-modal-input"
            value={location}
            maxLength={200}
            disabled={busy}
            onChange={(e) => setLocation(e.target.value)}
          />
        </label>
        <label className="pdf-field">
          <span>{t('certContact')}</span>
          <input
            className="pdf-modal-input"
            value={contact}
            maxLength={200}
            disabled={busy}
            onChange={(e) => setContact(e.target.value)}
          />
        </label>

        <label className="pdf-modal-check">
          <input
            type="checkbox"
            checked={visible}
            disabled={busy}
            onChange={(e) => setVisible(e.target.checked)}
          />
          {t('certVisible')}
        </label>
        {visible && (
          <div className="pdf-field">
            <span>{t('certPage')}</span>
            <input
              className="pdf-modal-input pdf-cert-page"
              type="number"
              min={1}
              max={pageCount}
              value={page}
              disabled={busy}
              onChange={(e) => setPage(Number(e.target.value) || 1)}
            />
            <span>{t('certPosition')}</span>
            <select
              className="pdf-modal-input"
              value={position}
              disabled={busy}
              onChange={(e) => setPosition(e.target.value as SignaturePosition)}
            >
              {POSITIONS.map((p) => (
                <option key={p.value} value={p.value}>
                  {t(p.key)}
                </option>
              ))}
            </select>
          </div>
        )}

        <label
          className="pdf-modal-check"
          data-tip={hasSignatures ? t('certCertifyUnavailable') : undefined}
        >
          <input
            type="checkbox"
            checked={certify && !hasSignatures}
            disabled={busy || hasSignatures}
            onChange={(e) => setCertify(e.target.checked)}
          />
          {t('certCertify')}
        </label>
        {certify && !hasSignatures && (
          <div className="pdf-field">
            <span />
            <select
              className="pdf-modal-input"
              value={level}
              disabled={busy}
              onChange={(e) => setLevel(Number(e.target.value) as 1 | 2 | 3)}
            >
              {([1, 2, 3] as const).map((value) => (
                <option key={value} value={value}>
                  {t(levelKeys[value])}
                </option>
              ))}
            </select>
          </div>
        )}

        {error && (
          <div className="pdf-cert-error" role="alert">
            {error}
          </div>
        )}
        <div className="pdf-modal-actions">
          <button className="pdf-modal-btn" onClick={onCancel} disabled={busy}>
            {t('cancel')}
          </button>
          <button
            className="pdf-modal-btn primary"
            onClick={() => void sign()}
            disabled={!signer || busy}
          >
            {busy ? t('certSigning') : t('certSignButton')}
          </button>
        </div>
      </div>
    </div>
  )
}
