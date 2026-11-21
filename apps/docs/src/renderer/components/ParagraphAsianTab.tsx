import { useState } from 'react'
import { Dropdown } from '@genoffice/ui'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'
import type { AsianTypographyState, TextAlignment } from './paragraph-asian-typography'
import { Unavailable } from './Unavailable'

/** the document-wide typography setting the Options button edits */
export interface TypographyOptions {
  /** settings.xml w:characterSpacingControl compresses punctuation */
  compressPunctuation: boolean
  onChange: (compressPunctuation: boolean) => void
  /** settings.xml w:strictFirstAndLastChars: the Strict level (it changes Japanese text only) */
  strictKinsoku: boolean
  onStrictKinsoku: (strict: boolean) => void
}

/** Word's Typography options, reached from Asian Typography > Options… */
function TypographyDialog({
  options,
  onClose,
}: {
  options: TypographyOptions
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [compress, setCompress] = useState(options.compressPunctuation)
  const [strict, setStrict] = useState(options.strictKinsoku)
  const radio = (name: string, label: StringKey, checked: boolean, pick?: () => void) => (
    <label className="font-check">
      <input
        type="radio"
        name={name}
        checked={checked}
        disabled={!pick}
        // the dialog opens on the kinsoku level in force: with the focus on the other radio,
        // Space or an arrow key would change the level by accident
        autoFocus={name === 'typo-kinsoku' && checked}
        onChange={pick}
        readOnly
      />
      {t(label)}
    </label>
  )
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={(e) => {
        e.stopPropagation()
        keys.onKeyDown(e)
      }}
      onMouseDown={(e) => {
        e.stopPropagation()
        if (e.target === e.currentTarget) onClose()
      }}
    >
      <div className="modal typography-dialog" role="dialog" aria-label={t('appTypoTitle')}>
        <h2>{t('appTypoTitle')}</h2>
        <h3 className="para-section-title">{t('appTypoKinsoku')}</h3>
        <div className="para-checks para-asian">
          {radio('typo-kinsoku', 'appTypoStandard', !strict, () => setStrict(false))}
          {radio('typo-kinsoku', 'appTypoStrict', strict, () => setStrict(true))}
          <Unavailable id="typography.kinsoku-custom">
            {radio('typo-kinsoku', 'appTypoCustom', false)}
          </Unavailable>
        </div>
        <h3 className="para-section-title">{t('appTypoSpacingControl')}</h3>
        <div className="para-checks para-asian">
          {radio('typo-spacing', 'appTypoNoCompress', !compress, () => setCompress(false))}
          {radio('typo-spacing', 'appTypoCompressPunct', compress, () => setCompress(true))}
          <Unavailable id="typography.compress-kana">
            {radio('typo-spacing', 'appTypoCompressKana', false)}
          </Unavailable>
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button
            className="btn-primary"
            onClick={() => {
              if (compress !== options.compressPunctuation) options.onChange(compress)
              if (strict !== options.strictKinsoku) options.onStrictKinsoku(strict)
              onClose()
            }}
          >
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

const TEXT_ALIGNMENTS: Array<{ value: TextAlignment; nameKey: StringKey }> = [
  { value: 'top', nameKey: 'ribbonMarginTop' },
  { value: 'center', nameKey: 'appScAlignCenter' },
  { value: 'baseline', nameKey: 'appParaAsianAlignBaseline' },
  { value: 'bottom', nameKey: 'ribbonMarginBottom' },
  { value: 'auto', nameKey: 'ribbonAuto' },
]

/** Word's Format > Paragraph > Asian Typography */
export function ParagraphAsianTab({
  state,
  onChange,
  typography,
}: {
  state: AsianTypographyState
  onChange: (next: AsianTypographyState) => void
  typography?: TypographyOptions
}) {
  const { t } = useI18n()
  const [showOptions, setShowOptions] = useState(false)
  const live = (key: Exclude<keyof AsianTypographyState, 'textAlignment'>, label: StringKey) => (
    <label className="font-check">
      <input
        type="checkbox"
        checked={state[key]}
        onChange={(e) => onChange({ ...state, [key]: e.target.checked })}
      />
      {t(label)}
    </label>
  )
  return (
    <>
      <h3 className="para-section-title">{t('appParaAsianLineBreak')}</h3>
      <div className="para-checks para-asian">
        {live('kinsoku', 'appParaAsianKinsoku')}
        {/* as in Word, w:wordWrap off lets hangul words split; Latin words stay whole */}
        {live('latinWrap', 'appParaAsianWordWrap')}
        {live('hangingPunct', 'appParaAsianHangPunct')}
      </div>
      <h3 className="para-section-title">{t('appParaAsianCharSpacing')}</h3>
      <div className="para-checks para-asian">
        {live('topLinePunct', 'appParaAsianTopLinePunct')}
        {live('autoSpaceLatin', 'appParaAsianAutoSpaceLatin')}
        {live('autoSpaceDigits', 'appParaAsianAutoSpaceDigits')}
      </div>
      <div className="font-dialog-row">
        <label>
          {t('appParaAsianTextAlign')}
          <Dropdown
            value={state.textAlignment}
            ariaLabel={t('appParaAsianTextAlign')}
            options={TEXT_ALIGNMENTS.map((a) => ({ value: a.value, label: t(a.nameKey) }))}
            onPick={(value) => onChange({ ...state, textAlignment: value })}
          />
        </label>
      </div>
      <div className="para-asian-options">
        <button type="button" disabled={!typography} onClick={() => setShowOptions(true)}>
          {t('ribbonPbOptions')}
        </button>
      </div>
      {showOptions && typography && (
        <TypographyDialog options={typography} onClose={() => setShowOptions(false)} />
      )}
    </>
  )
}
