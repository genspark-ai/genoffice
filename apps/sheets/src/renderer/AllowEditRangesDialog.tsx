/**
 * Review > Allow Edit Ranges: manage the active sheet's <protectedRanges>
 * entries (ranges that stay editable while the sheet is protected). The
 * session edits a working copy; OK hands the full replacement set back.
 */
import { useState } from 'react'

import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

export interface AllowEditRange {
  readonly name: string
  readonly sqref: string
  /// Plaintext password for a new range; hashed before anything touches disk,
  /// so this is never a stored secret.
  readonly password?: string
  /// Set for a range that carries a password; true together with `locked` only
  /// when the hash is not known to this session (a file's own hash), which is
  /// what makes the row unchangeable here.
  readonly hasPassword?: boolean
  readonly locked?: boolean
}

const RANGE_REF = /^\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?$/

export function AllowEditRangesDialog({
  ranges,
  defaultRef,
  onApply,
  onClose,
}: {
  readonly ranges: readonly AllowEditRange[]
  /// Seed for the range input (the active cell), '' when unknown.
  readonly defaultRef: string
  /// Returns an error message, or null on success (async allowed: hashing).
  readonly onApply: (ranges: readonly AllowEditRange[]) => string | null | Promise<string | null>
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [working, setWorking] = useState<readonly AllowEditRange[]>(ranges)
  const [name, setName] = useState(() => {
    for (let index = ranges.length + 1; ; index += 1) {
      const candidate = `Range${index}`
      if (!ranges.some((range) => range.name === candidate)) return candidate
    }
  })
  const [reference, setReference] = useState(defaultRef)
  const [password, setPassword] = useState('')
  const [error, setError] = useState<string | null>(null)

  const add = (): void => {
    const trimmedName = name.trim()
    const trimmedRef = reference.trim().toUpperCase()
    if (trimmedName === '' || working.some((range) => range.name === trimmedName)) {
      setError(t('dlgRangeNameTaken'))
      return
    }
    if (!RANGE_REF.test(trimmedRef)) {
      setError(t('dlgRangeRefInvalid'))
      return
    }
    const next = [
      ...working,
      {
        name: trimmedName,
        sqref: trimmedRef.replaceAll('$', ''),
        ...(password.trim() === '' ? {} : { password: password.trim() }),
      },
    ]
    setWorking(next)
    setError(null)
    setName(`Range${next.length + 1}`)
    setReference('')
    setPassword('')
  }

  const [applying, setApplying] = useState(false)

  const apply = (): void => {
    setApplying(true)
    void Promise.resolve(onApply(working))
      .then((failed) => {
        if (failed !== null) {
          setError(failed)
          setApplying(false)
          return
        }
        onClose()
      })
      .catch(() => setApplying(false))
  }

  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        {...modal}
        aria-label={t('dlgAllowEditRangesTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgAllowEditRangesTitle')}</header>
        <section className="dialog-body">
          <p className="dialog-note">{t('dlgAllowEditRangesHint')}</p>
          {working.length > 0 && (
            <ul className="allow-edit-list">
              {working.map((range) => (
                <li key={range.name}>
                  <span className="allow-edit-name">
                    {range.name}
                    {range.locked === true && (
                      <span className="allow-edit-lock" aria-hidden="true">
                        {' 🔒'}
                      </span>
                    )}
                  </span>
                  <span className="allow-edit-ref">{range.sqref}</span>
                  <button
                    className="secondary"
                    disabled={range.locked === true}
                    title={range.locked === true ? t('dlgRangeLocked') : undefined}
                    onClick={() => setWorking(working.filter((entry) => entry.name !== range.name))}
                  >
                    {t('dlgRangeDelete')}
                  </button>
                </li>
              ))}
            </ul>
          )}
          <div className="dialog-grid">
            <label>
              {t('dlgRangeName')}
              <input value={name} onChange={(event) => setName(event.target.value)} />
            </label>
            <label>
              {t('dlgRangeRef')}
              <input
                autoFocus
                placeholder="A1:B4"
                value={reference}
                onChange={(event) => setReference(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') add()
                }}
              />
            </label>
            <label>
              {t('dlgRangePassword')}
              <input
                type="password"
                value={password}
                onChange={(event) => setPassword(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') add()
                }}
              />
            </label>
          </div>
          <div className="dialog-actions-inline">
            <button className="secondary" onClick={add}>
              {t('dlgRangeAdd')}
            </button>
          </div>
          {error && (
            <p className="dialog-note dialog-error" role="alert">
              {error}
            </p>
          )}
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary-action" disabled={applying} onClick={apply}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
