/**
 * Word-style "Protect Document" dialog (Review > Protect), modeled on Word for
 * Mac's Password Protect sheet. One dialog covers all protection features:
 *
 * - Security: password to open (ECMA-376 whole-file encryption) and password
 *   to modify (settings.xml w:writeProtection, honor-system).
 * - Protection: editing restriction (w:documentProtection) with the four Word
 *   modes (tracked changes / comments / read only / forms) + optional password.
 * - Privacy: remove known author and organization metadata on save.
 *
 * The dialog only computes a diff (ProtectDialogResult); applying it (IPC for
 * the open password, dirty-state for the rest) is the caller's job.
 */
import { useState } from 'react'
import {
  hashProtectionPassword,
  verifyProtectionPassword,
  type DocProtection,
  type WriteProtection,
} from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import { FieldError, PasswordInput } from './PasswordInput'
import { useModalKeys } from './modal-keys'

/** every field: undefined = unchanged; null = remove; value = set */
export interface ProtectDialogResult {
  openPassword?: string | null
  writeProtection?: WriteProtection | null
  protection?: DocProtection | null
  removePersonalInfo?: boolean
  /** per section: protected under forms protection (w:formProt); present when it changed */
  sectionsProtected?: boolean[]
}

export const PROTECTION_MODES = ['trackedChanges', 'comments', 'readOnly', 'forms'] as const
export type ProtectionMode = (typeof PROTECTION_MODES)[number]

const MODE_LABEL_KEYS = {
  trackedChanges: 'appProtectModeTracked',
  comments: 'appProtectModeComments',
  readOnly: 'appProtectModeReadOnly',
  forms: 'appProtectModeForms',
} as const

/**
 * Existing passwords are only known as hashes, so the fields are prefilled
 * with an untypable sentinel that renders as dots: submitting it unchanged
 * keeps the password, clearing the field removes it.
 */
const KEEP = '\u0001'.repeat(8)

type ErrorKey = '' | 'appEncMismatch' | 'appWrongPassword'

export function ProtectDialog({
  encrypted,
  writeProtection,
  protection,
  sectionsProtected = [],
  removePersonalInfo,
  onCancel,
  onApply,
}: {
  /** an open password is desired for the next save */
  encrypted: boolean
  writeProtection: WriteProtection | null
  protection: DocProtection | null
  /** per section: protected when the document is protected for forms */
  sectionsProtected?: boolean[]
  removePersonalInfo: boolean
  onCancel: () => void
  onApply: (result: ProtectDialogResult) => void
}) {
  const { t } = useI18n()
  const hadModifyPwd = !!writeProtection?.hash
  const wasEnforced = !!protection?.enforced
  /** changing/removing an enforced password-protected restriction needs the password */
  const locked = wasEnforced && !!protection?.hash

  const [openPwd, setOpenPwd] = useState(encrypted ? KEEP : '')
  const [openPwd2, setOpenPwd2] = useState(encrypted ? KEEP : '')
  const [modifyPwd, setModifyPwd] = useState(hadModifyPwd ? KEEP : '')
  const [modifyPwd2, setModifyPwd2] = useState(hadModifyPwd ? KEEP : '')
  const [protectOn, setProtectOn] = useState(wasEnforced)
  const [mode, setMode] = useState<ProtectionMode>(
    PROTECTION_MODES.includes(protection?.edit as ProtectionMode)
      ? (protection?.edit as ProtectionMode)
      : 'trackedChanges',
  )
  const [protectPwd, setProtectPwd] = useState(locked ? KEEP : '')
  const [unlockPwd, setUnlockPwd] = useState('')
  const [removePersonal, setRemovePersonal] = useState(removePersonalInfo)
  const [sectionChecks, setSectionChecks] = useState(sectionsProtected)
  const [showSections, setShowSections] = useState(false)
  const sectionsChanged = sectionChecks.some((on, i) => on !== sectionsProtected[i])
  const [errorKey, setErrorKey] = useState<ErrorKey>('')
  const [busy, setBusy] = useState(false)

  // which sections a forms protection covers is part of the restriction: it takes the password too
  const protectionChanged =
    protectOn !== wasEnforced ||
    (protectOn &&
      (mode !== protection?.edit ||
        protectPwd !== (locked ? KEEP : '') ||
        (mode === 'forms' && sectionsChanged)))

  const submit = async () => {
    if (busy) return
    if (openPwd !== openPwd2 || modifyPwd !== modifyPwd2) {
      setErrorKey('appEncMismatch')
      return
    }
    setBusy(true)
    try {
      const result: ProtectDialogResult = {}

      if (openPwd !== (encrypted ? KEEP : '')) {
        result.openPassword = openPwd === '' ? null : openPwd
      }

      if (modifyPwd !== (hadModifyPwd ? KEEP : '')) {
        const recommended = writeProtection?.recommended ? { recommended: true } : {}
        result.writeProtection =
          modifyPwd === ''
            ? writeProtection?.recommended
              ? { recommended: true }
              : null
            : { ...recommended, ...(await hashProtectionPassword(modifyPwd)) }
      }

      if (protectionChanged) {
        if (locked && !(await verifyProtectionPassword(unlockPwd, protection!))) {
          setErrorKey('appWrongPassword')
          return
        }
        if (!protectOn) {
          result.protection = null
        } else {
          const creds =
            protectPwd === KEEP && protection?.hash
              ? {
                  hash: protection.hash,
                  salt: protection.salt,
                  spinCount: protection.spinCount,
                  algorithmSid: protection.algorithmSid,
                }
              : protectPwd && protectPwd !== KEEP
                ? await hashProtectionPassword(protectPwd)
                : {}
          result.protection = { edit: mode, enforced: true, ...creds }
        }
      }

      if (removePersonal !== removePersonalInfo) {
        result.removePersonalInfo = removePersonal
      }
      // a forms protection that starts now covers every section left checked, whatever the file said
      const formsStarts =
        protectOn && mode === 'forms' && !(wasEnforced && protection?.edit === 'forms')
      if (
        sectionChecks.length > 0 &&
        (formsStarts || (protectOn && mode === 'forms' && sectionsChanged))
      ) {
        result.sectionsProtected = sectionChecks
      }

      onApply(result)
    } finally {
      setBusy(false)
    }
  }

  const onEnter = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') void submit()
  }
  const clearError = () => setErrorKey('')

  const mismatch = errorKey === 'appEncMismatch'

  return (
    <div className="modal-backdrop">
      <div className="modal protect-dialog gs-form">
        <h2>{t('appProtectTitle')}</h2>
        <p className="modal-desc">{t('appProtectDesc')}</p>

        <h3 className="protect-section-title">{t('appProtectSecurity')}</h3>
        <div className="fld-row">
          <label className="fld">
            {t('appProtectOpenPwd')} <span className="fld-opt">{t('appOptional')}</span>
            <PasswordInput
              autoFocus
              maxLength={255}
              value={openPwd}
              hideReveal={openPwd === KEEP}
              onChange={(v) => {
                setOpenPwd(v)
                clearError()
              }}
              onKeyDown={onEnter}
            />
          </label>
          <label className="fld">
            {t('appEncConfirmLabel')}
            <PasswordInput
              maxLength={255}
              value={openPwd2}
              invalid={mismatch}
              hideReveal={openPwd2 === KEEP}
              onChange={(v) => {
                setOpenPwd2(v)
                clearError()
              }}
              onKeyDown={onEnter}
            />
          </label>
        </div>
        <div className="fld-row">
          <label className="fld">
            {t('appProtectModifyPwd')} <span className="fld-opt">{t('appOptional')}</span>
            <PasswordInput
              maxLength={255}
              value={modifyPwd}
              hideReveal={modifyPwd === KEEP}
              onChange={(v) => {
                setModifyPwd(v)
                clearError()
              }}
              onKeyDown={onEnter}
            />
          </label>
          <label className="fld">
            {t('appEncConfirmLabel')}
            <PasswordInput
              maxLength={255}
              value={modifyPwd2}
              invalid={mismatch}
              hideReveal={modifyPwd2 === KEEP}
              onChange={(v) => {
                setModifyPwd2(v)
                clearError()
              }}
              onKeyDown={onEnter}
            />
          </label>
        </div>
        {(encrypted || hadModifyPwd) && <p className="fld-hint">{t('appProtectPwdKeepHint')}</p>}

        <h3 className="protect-section-title">{t('appProtectSectionTitle')}</h3>
        <label className="protect-check">
          <input
            type="checkbox"
            checked={protectOn}
            onChange={(e) => {
              setProtectOn(e.target.checked)
              clearError()
            }}
          />
          <span className="ctl" aria-hidden="true" />
          {t('appProtectFor')}
        </label>
        <div className="protect-modes">
          {PROTECTION_MODES.map((m) => (
            <label key={m} className="protect-check">
              <input
                type="radio"
                name="protect-mode"
                disabled={!protectOn}
                checked={mode === m}
                onChange={() => {
                  setMode(m)
                  clearError()
                }}
              />
              <span className="ctl" aria-hidden="true" />
              {t(MODE_LABEL_KEYS[m])}
            </label>
          ))}
          <button
            type="button"
            className="protect-sections"
            disabled={!protectOn || mode !== 'forms' || sectionChecks.length === 0}
            onClick={() => setShowSections(true)}
          >
            {t('appProtectSections')}
          </button>
        </div>
        <label className="fld">
          {t('appProtectPwdOptional')} <span className="fld-opt">{t('appOptionalBlank')}</span>
          <PasswordInput
            maxLength={255}
            disabled={!protectOn}
            value={protectOn ? protectPwd : ''}
            hideReveal={protectPwd === KEEP}
            onChange={(v) => {
              setProtectPwd(v)
              clearError()
            }}
            onKeyDown={onEnter}
          />
        </label>
        {locked && protectionChanged && (
          <label className="fld">
            {t('appProtectUnlockPwd')}
            <PasswordInput
              value={unlockPwd}
              invalid={errorKey === 'appWrongPassword'}
              onChange={(v) => {
                setUnlockPwd(v)
                clearError()
              }}
              onKeyDown={onEnter}
            />
          </label>
        )}

        <h3 className="protect-section-title">{t('appProtectPrivacy')}</h3>
        <label className="protect-check">
          <input
            type="checkbox"
            checked={removePersonal}
            onChange={(e) => setRemovePersonal(e.target.checked)}
          />
          <span className="ctl" aria-hidden="true" />
          {t('appProtectRemovePersonal')}
        </label>

        {errorKey && <FieldError>{t(errorKey)}</FieldError>}
        <div className="modal-actions">
          <button onClick={onCancel}>{t('appCancel')}</button>
          <button className="btn-primary" disabled={busy} onClick={() => void submit()}>
            {t('appOk')}
          </button>
        </div>
      </div>
      {showSections && (
        <SectionProtectionDialog
          checks={sectionChecks}
          onCancel={() => setShowSections(false)}
          onOk={(next) => {
            setSectionChecks(next)
            setShowSections(false)
            clearError()
          }}
        />
      )}
    </div>
  )
}

/** Word's Section Protection: the sections a forms protection covers */
function SectionProtectionDialog({
  checks,
  onCancel,
  onOk,
}: {
  checks: boolean[]
  onCancel: () => void
  onOk: (checks: boolean[]) => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onCancel)
  const [draft, setDraft] = useState(checks)
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={(e) => {
        e.stopPropagation()
        keys.onKeyDown(e)
      }}
    >
      <div
        className="modal protect-dialog protect-sections-dialog gs-form"
        role="dialog"
        aria-label={t('appProtectSectionsTitle')}
      >
        <h2>{t('appProtectSectionsTitle')}</h2>
        <p className="modal-desc">{t('appProtectSectionsLabel')}</p>
        <div className="protect-sections-list">
          {draft.map((on, i) => (
            <label key={i} className="protect-check">
              <input
                type="checkbox"
                checked={on}
                autoFocus={i === 0}
                onChange={(e) => setDraft(draft.map((v, k) => (k === i ? e.target.checked : v)))}
              />
              <span className="ctl" aria-hidden="true" />
              {t('appProtectSectionN', { n: i + 1 })}
            </label>
          ))}
        </div>
        <div className="modal-actions">
          <button onClick={onCancel}>{t('appCancel')}</button>
          <button className="btn-primary" onClick={() => onOk(draft)}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
