import { useMemo, useState } from 'react'
import { detectLanguageEnabled, setDetectLanguageEnabled } from '../editor/language-detect'
import { PROOFING_LANGUAGES } from '../editor/proofing-language'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/** a language's name in the UI language, falling back to its tag */
export function languageName(tag: string, uiLang: string): string {
  try {
    return new Intl.DisplayNames([uiLang], { type: 'language' }).of(tag) ?? tag
  } catch {
    return tag
  }
}

/** Word's Tools > Language: the language the selected text is proofed as */
export function LanguageDialog({
  current,
  noProof,
  onApply,
  onClose,
}: {
  /** the tag the selection is marked as; '' when it is not marked */
  current: string
  /** whether the selection is marked as not to be checked */
  noProof: boolean
  /** `skip` is passed only when the box was changed */
  onApply: (tag: string, skip?: boolean) => void
  onClose: () => void
}) {
  const { t, lang } = useI18n()
  const keys = useModalKeys(onClose)
  const tags = useMemo(
    () =>
      current && !(PROOFING_LANGUAGES as readonly string[]).includes(current)
        ? [current, ...PROOFING_LANGUAGES]
        : [...PROOFING_LANGUAGES],
    [current],
  )
  const [opened] = useState(current || tags[0])
  const [picked, setPicked] = useState(opened)
  const [skip, setSkip] = useState(noProof)
  const [detect, setDetect] = useState(detectLanguageEnabled)
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog language-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          const onlyDetect =
            detect !== detectLanguageEnabled() && picked === opened && skip === noProof
          setDetectLanguageEnabled(detect)
          // turning detection on or off alone leaves the selection's language as it is
          if (!onlyDetect) onApply(picked, skip === noProof ? undefined : skip)
        }}
      >
        <h2>{t('ribbonProofingLanguage')}</h2>
        <select
          className="language-list"
          size={10}
          autoFocus
          aria-label={t('ribbonProofingLanguage')}
          value={picked}
          onChange={(e) => setPicked(e.target.value)}
          onDoubleClick={(e) => e.currentTarget.form?.requestSubmit()}
        >
          {tags.map((tag) => (
            <option key={tag} value={tag}>
              {languageName(tag, lang)}
            </option>
          ))}
        </select>
        <label className="tco-check">
          <input
            type="checkbox"
            name="noProof"
            checked={skip}
            onChange={(e) => setSkip(e.target.checked)}
          />
          {t('ribbonProofingNoCheck')}
        </label>
        <label className="tco-check">
          <input
            type="checkbox"
            name="detect"
            checked={detect}
            onChange={(e) => setDetect(e.target.checked)}
          />
          {t('ribbonProofingDetect')}
        </label>
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
