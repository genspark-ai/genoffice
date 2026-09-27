import { useState } from 'react'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/** Edit Field…: Word's Field dialog reduced to the instruction string (the result is recomputed on OK) */
export function FieldDialog({
  instr,
  onSubmit,
  onClose,
}: {
  instr: string
  onSubmit: (instr: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [code, setCode] = useState(instr)
  const submit = () => {
    if (!code.trim()) return
    onSubmit(code.trim())
    onClose()
  }
  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal">
        <h2>{t('appFieldDialogTitle')}</h2>
        <label>
          {t('appFieldCode')}
          <input
            autoFocus
            value={code}
            spellCheck={false}
            onChange={(e) => setCode(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && submit()}
          />
        </label>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" disabled={!code.trim()} onClick={submit}>
            {t('ribbonOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
