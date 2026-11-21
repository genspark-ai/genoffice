import { useState } from 'react'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'
import { TRANSLATE_LANGS, ribbonLangKey, type TranslateLang } from './translate-langs'

/** Word's Tools > Translate > Set Document Translation Language: what Translate turns text into */
export function TranslationLanguageDialog({
  current,
  onApply,
  onClose,
}: {
  current: TranslateLang
  onApply: (lang: TranslateLang) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [lang, setLang] = useState(current)
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog translation-language-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          onApply(lang)
        }}
      >
        <h2>{t('ribbonTranslationLanguageTitle')}</h2>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonTranslationLanguageLabel')}
            <select value={lang} onChange={(e) => setLang(e.target.value as TranslateLang)}>
              {TRANSLATE_LANGS.map(({ code }) => (
                <option key={code} value={code}>
                  {t(ribbonLangKey(code))}
                </option>
              ))}
            </select>
          </label>
        </div>
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary">
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}
