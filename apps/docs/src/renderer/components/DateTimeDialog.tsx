import { useMemo, useState } from 'react'
import { dateTimePicturesFor, formatDatePicture, toFullWidth } from '../editor/date-picture'
import { PROOFING_LANGUAGES } from '../editor/proofing-language'
import { languageName } from './LanguageDialog'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

export interface DateTimePick {
  picture: string
  asField: boolean
  /** the language the names in the date are written in */
  lang: string
  fullWidth: boolean
}

/** Word's Insert > Date and Time: a format to pick, as a field that updates or as plain text */
export function DateTimeDialog({
  language,
  onInsert,
  onClose,
}: {
  /** the language of the text at the caret; '' when it is not marked */
  language: string
  onInsert: (pick: DateTimePick) => void
  onClose: () => void
}) {
  const { t, lang: uiLang } = useI18n()
  const keys = useModalKeys(onClose)
  const [now] = useState(() => new Date())
  const tags = useMemo(
    () =>
      language && !(PROOFING_LANGUAGES as readonly string[]).includes(language)
        ? [language, ...PROOFING_LANGUAGES]
        : [...PROOFING_LANGUAGES],
    [language],
  )
  const [lang, setLang] = useState(language || tags[0])
  const pictures = useMemo(() => dateTimePicturesFor(lang), [lang])
  const [picked, setPicked] = useState(0)
  const picture = pictures[Math.min(picked, pictures.length - 1)]
  const [asField, setAsField] = useState(false)
  const [fullWidth, setFullWidth] = useState(false)
  const shown = (p: string) => {
    const text = formatDatePicture(now, p, lang)
    return fullWidth ? toFullWidth(text) : text
  }
  const submit = () => {
    onClose()
    onInsert({ picture, asField, lang, fullWidth })
  }
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog date-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          submit()
        }}
      >
        <h2>{t('ribbonDateTimeTitle')}</h2>
        <div className="doc-dialog-caption">{t('ribbonDateTimeFormats')}</div>
        <div className="date-dialog-list" role="listbox" aria-label={t('ribbonDateTimeFormats')}>
          {pictures.map((p, i) => (
            <button
              key={p}
              type="button"
              role="option"
              aria-selected={picture === p}
              className={picture === p ? 'active' : ''}
              onClick={() => setPicked(i)}
              onDoubleClick={submit}
            >
              {shown(p)}
            </button>
          ))}
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonProofingLanguage')}
            <select
              name="language"
              value={lang}
              onChange={(e) => {
                setLang(e.target.value)
                setPicked(0)
              }}
            >
              {tags.map((tag) => (
                <option key={tag} value={tag}>
                  {languageName(tag, uiLang)}
                </option>
              ))}
            </select>
          </label>
        </div>
        <label className="tco-check">
          <input type="checkbox" checked={asField} onChange={(e) => setAsField(e.target.checked)} />
          {t('ribbonDateTimeUpdate')}
        </label>
        <label className="tco-check">
          <input
            type="checkbox"
            name="fullWidth"
            checked={fullWidth}
            onChange={(e) => setFullWidth(e.target.checked)}
          />
          {t('ribbonDateTimeFullWidth')}
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
