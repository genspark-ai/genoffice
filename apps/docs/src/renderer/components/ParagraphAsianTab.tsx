import { useState } from 'react'
import { Dropdown } from '@genoffice/ui'
import type { KinsokuCustom } from '@genoffice/docx-engine'
import type { SpacingControl } from '../editor/cjk-punct-shrink'
import {
  KINSOKU_LANGS,
  STANDARD_KINSOKU,
  kinsokuLang,
  type KinsokuLang,
} from '../editor/kinsoku-tables'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'
import type { AsianTypographyState, TextAlignment } from './paragraph-asian-typography'

const KINSOKU_LANG_KEYS: Record<KinsokuLang, StringKey> = {
  'ja-JP': 'appLangJapanese',
  'zh-CN': 'appLangSimplifiedChinese',
  'zh-TW': 'appLangTraditionalChinese',
  'ko-KR': 'appLangKorean',
}

/** the document-wide typography setting the Options button edits */
export interface TypographyOptions {
  /** settings.xml w:characterSpacingControl */
  spacingControl: SpacingControl
  onChange: (spacingControl: SpacingControl) => void
  /** settings.xml w:strictFirstAndLastChars: the Strict level (it changes Japanese text only) */
  strictKinsoku: boolean
  /** settings.xml w:noLineBreaksBefore / After: the Custom level */
  kinsokuCustom: KinsokuCustom | null
  onKinsoku: (strict: boolean, custom: KinsokuCustom | null) => void
}

type KinsokuLevel = 'standard' | 'strict' | 'custom'

/** a list is a set of characters: line breaks and repeats typed into the box do not count */
function cleanLists(lists: { before: string; after: string }) {
  const clean = (chars: string) => [...new Set(chars.replace(/\s/g, ''))].join('')
  return { before: clean(lists.before), after: clean(lists.after) }
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
  const [control, setControl] = useState(options.spacingControl)
  const initialLevel: KinsokuLevel = options.strictKinsoku
    ? 'strict'
    : options.kinsokuCustom
      ? 'custom'
      : 'standard'
  const [level, setLevel] = useState(initialLevel)
  // a list in the file may name a language the dropdown does not offer: it is kept as it is
  const fileLang = options.kinsokuCustom?.lang ?? ''
  const [lang, setLang] = useState(
    options.kinsokuCustom ? (kinsokuLang(fileLang || 'x') ?? fileLang) : 'ja-JP',
  )
  const standardOf = (l: string) =>
    (STANDARD_KINSOKU as Record<string, { before: string; after: string }>)[l] ?? {
      before: '',
      after: '',
    }
  const [lists, setLists] = useState(options.kinsokuCustom ?? { ...standardOf('ja-JP') })
  const custom: KinsokuCustom = { lang, before: lists.before, after: lists.after }
  const customChanged =
    level === 'custom' &&
    (!options.kinsokuCustom ||
      lists.before !== options.kinsokuCustom.before ||
      lists.after !== options.kinsokuCustom.after ||
      (lang !== fileLang && lang !== kinsokuLang(fileLang || 'x')))
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
          {radio('typo-kinsoku', 'appTypoStandard', level === 'standard', () =>
            setLevel('standard'),
          )}
          {radio('typo-kinsoku', 'appTypoStrict', level === 'strict', () => setLevel('strict'))}
          {radio('typo-kinsoku', 'appTypoCustom', level === 'custom', () => setLevel('custom'))}
        </div>
        {level === 'custom' && (
          <div className="typography-custom">
            <label>
              {t('appTypoKinsokuLang')}
              <Dropdown
                value={lang}
                ariaLabel={t('appTypoKinsokuLang')}
                options={[
                  ...KINSOKU_LANGS.map((l) => ({
                    value: l as string,
                    label: t(KINSOKU_LANG_KEYS[l]),
                  })),
                  ...(KINSOKU_LANGS.includes(lang as KinsokuLang)
                    ? []
                    : [{ value: lang, label: lang }]),
                ]}
                onPick={(value) => {
                  // as in Word, another language starts from its own standard lists
                  setLang(value)
                  setLists({ ...standardOf(value) })
                }}
              />
            </label>
            <label>
              {t('appTypoNoStart')}
              <textarea
                rows={2}
                spellCheck={false}
                value={lists.before}
                onChange={(e) => setLists({ ...lists, before: e.target.value })}
              />
            </label>
            <label>
              {t('appTypoNoEnd')}
              <textarea
                rows={2}
                spellCheck={false}
                value={lists.after}
                onChange={(e) => setLists({ ...lists, after: e.target.value })}
              />
            </label>
          </div>
        )}
        <h3 className="para-section-title">{t('appTypoSpacingControl')}</h3>
        <div className="para-checks para-asian">
          {radio('typo-spacing', 'appTypoNoCompress', control === 'none', () => setControl('none'))}
          {radio('typo-spacing', 'appTypoCompressPunct', control === 'punct', () =>
            setControl('punct'),
          )}
          {radio('typo-spacing', 'appTypoCompressKana', control === 'kana', () =>
            setControl('kana'),
          )}
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button
            className="btn-primary"
            onClick={() => {
              if (control !== options.spacingControl) options.onChange(control)
              if (level !== initialLevel || customChanged)
                options.onKinsoku(
                  level === 'strict',
                  level === 'custom' ? { ...custom, ...cleanLists(lists) } : null,
                )
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
