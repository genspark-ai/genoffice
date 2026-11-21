import { useState } from 'react'
import { ENCLOSURES, type EncloseStyle } from '@genoffice/docx-engine'
import { ENCLOSE_MAX, encloseMarkClass } from '../editor/eq-layout'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

const STYLES: Array<{ style: EncloseStyle; labelKey: StringKey }> = [
  { style: 'none', labelKey: 'ribbonNone' },
  { style: 'shrink', labelKey: 'ribbonEcShrink' },
  { style: 'enlarge', labelKey: 'ribbonEcEnlarge' },
]

/** Word's Format > Asian Layout > Enclose Characters: a character inside a circle, square, triangle or diamond */
export function EncloseCharactersDialog({
  text,
  enclosure,
  style,
  onApply,
  onClose,
}: {
  text: string
  /** the enclosure and style it already has; null for plain text */
  enclosure: string | null
  style: EncloseStyle | null
  /** null takes the enclosure off (Word's style None) */
  onApply: (
    layout: { text: string; enclosure: string; style: Exclude<EncloseStyle, 'none'> } | null,
  ) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [chars, setChars] = useState(text)
  const [mark, setMark] = useState(enclosure ?? ENCLOSURES[0])
  const [fit, setFit] = useState<EncloseStyle>(style ?? 'shrink')
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
          onApply(fit === 'none' ? null : { text: chars, enclosure: mark, style: fit })
        }}
      >
        <h2>{t('ribbonEncloseCharacters')}</h2>
        <fieldset className="eq-layout-choices">
          <legend>{t('ribbonPbStyle')}</legend>
          {STYLES.map((item) => (
            <button
              key={item.style}
              type="button"
              role="radio"
              aria-checked={fit === item.style}
              className={`eq-layout-choice${fit === item.style ? ' active' : ''}`}
              onClick={() => setFit(item.style)}
            >
              <span className="eq-layout-sample">
                {item.style === 'none' ? (
                  chars || ' '
                ) : (
                  <span
                    className={`doc-enclose-chars doc-enclose-${item.style}${Array.from(chars).length > 1 ? ' doc-enclose-two' : ''}`}
                  >
                    <span className={`doc-enclose-mark ${encloseMarkClass(mark)}`}>{mark}</span>
                    <span className="doc-enclose-text">{chars}</span>
                  </span>
                )}
              </span>
              <span>{t(item.labelKey)}</span>
            </button>
          ))}
        </fieldset>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonEcText')}
            <input
              type="text"
              value={chars}
              autoFocus
              spellCheck={false}
              onChange={(e) => setChars(Array.from(e.target.value).slice(0, ENCLOSE_MAX).join(''))}
            />
          </label>
        </div>
        <fieldset className="eq-layout-choices">
          <legend>{t('ribbonEcEnclosure')}</legend>
          {ENCLOSURES.map((symbol) => (
            <button
              key={symbol}
              type="button"
              role="radio"
              aria-checked={mark === symbol}
              aria-label={symbol}
              className={`eq-layout-choice eq-layout-mark${mark === symbol ? ' active' : ''}`}
              onClick={() => setMark(symbol)}
            >
              <span className="eq-layout-sample">
                <span className="doc-enclose-chars">
                  <span className={`doc-enclose-mark ${encloseMarkClass(symbol)}`}>{symbol}</span>
                </span>
              </span>
            </button>
          ))}
        </fieldset>
        <div className="modal-actions">
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
