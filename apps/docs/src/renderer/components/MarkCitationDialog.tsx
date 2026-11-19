import { useState } from 'react'
import { CITATION_CATEGORIES, type Citation } from '../editor/citations'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/** Word's names for the table of authorities categories, by the number a citation carries */
export const CITATION_CATEGORY_LABELS: Record<number, StringKey> = {
  1: 'ribbonToaCases',
  2: 'ribbonToaStatutes',
  3: 'ribbonToaOther',
  4: 'ribbonToaRules',
  5: 'ribbonToaTreatises',
  6: 'ribbonToaRegulations',
  7: 'ribbonToaConstitutional',
}

/** Word's References > Mark Citation: the text the table of authorities lists, and its category */
export function MarkCitationDialog({
  selected,
  onMark,
  onClose,
}: {
  /** the selected text: what the long citation starts as */
  selected: string
  /** `all`: every other mention of the short citation is marked too */
  onMark: (citation: Citation, all: boolean) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [long, setLong] = useState(selected)
  const [short, setShort] = useState(selected)
  const [category, setCategory] = useState<number>(CITATION_CATEGORIES[0])
  const text = long.trim()
  const mark = (all: boolean) => {
    if (!text) return
    onClose()
    onMark({ long: text, short: short.trim(), category }, all)
  }
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog mark-citation-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          mark(false)
        }}
      >
        <h2>{t('ribbonMarkCitation')}</h2>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonToaLong')}
            <input
              type="text"
              name="long"
              value={long}
              autoFocus
              spellCheck={false}
              onChange={(e) => setLong(e.target.value)}
            />
          </label>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonToaShort')}
            <input
              type="text"
              name="short"
              value={short}
              spellCheck={false}
              onChange={(e) => setShort(e.target.value)}
            />
          </label>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonToaCategory')}
            <select value={category} onChange={(e) => setCategory(Number(e.target.value))}>
              {CITATION_CATEGORIES.map((n) => (
                <option key={n} value={n}>
                  {t(CITATION_CATEGORY_LABELS[n])}
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
            <button type="button" disabled={!text || !short.trim()} onClick={() => mark(true)}>
              {t('ribbonXeMarkAll')}
            </button>
            <button type="submit" className="primary" disabled={!text}>
              {t('ribbonXeMark')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}
