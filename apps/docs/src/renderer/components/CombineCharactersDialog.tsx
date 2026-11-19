import { useState } from 'react'
import { COMBINE_MAX, combineLines, type EqLayoutFont } from '../editor/eq-layout'
import { fontFamiliesFor } from '../font-list'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]

/** Word's Format > Asian Layout > Combine Characters: up to six characters on two lines in one character's place */
export function CombineCharactersDialog({
  text,
  combined,
  font,
  sizePt,
  onApply,
  onClose,
}: {
  text: string
  /** the text is combined already: Remove is offered */
  combined: boolean
  /** the font and size the text names itself; empty and 0 when it inherits them */
  font: string
  sizePt: number
  /** null takes the combination off; a font or size left alone is not passed */
  onApply: (text: string | null, picked: EqLayoutFont) => void
  onClose: () => void
}) {
  const { t, lang } = useI18n()
  const keys = useModalKeys(onClose)
  const [chars, setChars] = useState(text)
  const [pickedFont, setPickedFont] = useState(font)
  const [pickedSize, setPickedSize] = useState(sizePt)
  const families = fontFamiliesFor(lang)
  const fonts = !font || families.includes(font) ? families : [font, ...families]
  const sizes =
    SIZES.includes(sizePt) || sizePt <= 0 ? SIZES : [...SIZES, sizePt].sort((a, b) => a - b)
  const lines = combineLines(chars)
  const ready = chars.trim() !== ''
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog eq-layout-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          if (!ready) return
          onClose()
          onApply(chars, {
            ...(pickedFont !== font ? { font: pickedFont } : {}),
            ...(pickedSize !== sizePt ? { sizePt: pickedSize } : {}),
          })
        }}
      >
        <h2>{t('ribbonCcTitle')}</h2>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonCcCharacters')}
            <input
              type="text"
              value={chars}
              autoFocus
              spellCheck={false}
              onChange={(e) => setChars(Array.from(e.target.value).slice(0, COMBINE_MAX).join(''))}
            />
          </label>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonFontFamilyTip')}
            <select name="font" value={pickedFont} onChange={(e) => setPickedFont(e.target.value)}>
              {!font && <option value="">{t('ribbonCcSameAsText')}</option>}
              {fonts.map((name) => (
                <option key={name} value={name}>
                  {name}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t('appFontSizeLabel')}
            <select
              name="size"
              value={pickedSize}
              onChange={(e) => setPickedSize(Number(e.target.value))}
            >
              {sizePt <= 0 && <option value={0}>{t('ribbonCcSameAsText')}</option>}
              {sizes.map((pt) => (
                <option key={pt} value={pt}>
                  {pt}
                </option>
              ))}
            </select>
          </label>
        </div>
        <div className="eq-layout-preview">
          <span
            className="doc-combine-chars"
            style={{ fontFamily: pickedFont ? `"${pickedFont}"` : undefined }}
          >
            <span>{lines.a}</span>
            <span>{lines.b}</span>
          </span>
        </div>
        <div className="modal-actions">
          <button
            type="button"
            disabled={!combined}
            onClick={() => {
              onClose()
              onApply(null, {})
            }}
          >
            {t('ribbonFitRemove')}
          </button>
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary" disabled={!ready}>
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}
