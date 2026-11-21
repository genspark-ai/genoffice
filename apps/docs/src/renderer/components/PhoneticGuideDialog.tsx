import { useMemo, useState } from 'react'
import { RUBY_ALIGNS, type RubyAlign } from '@genoffice/docx-engine'
import { RUBY_BASE_MAX, type RubyEdit, type RubyReading } from '../editor/phonetic-guide'
import { fontFamiliesFor } from '../font-list'
import { cssFontFamily } from '../text-layout/font-rules/css-chains'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/** Word names the two distributed placements by their spacing pattern, in every language */
const ALIGN_LABELS: Record<RubyAlign, StringKey | { literal: string }> = {
  center: 'ribbonPgCentered',
  distributeLetter: { literal: '0-1-0' },
  distributeSpace: { literal: '1-2-1' },
  left: 'appAlignLeft',
  right: 'appAlignRight',
}

/** Word's Format > Asian Layout > Phonetic Guide: the reading shown above the base text */
export function PhoneticGuideDialog({
  base,
  rt,
  align,
  reading: set,
  onApply,
  onClose,
}: {
  /** how the reading is set now, or would be on new text */
  reading: RubyReading
  base: string
  /** the reading the text already has; empty when it has none */
  rt: string
  /** the alignment the guide already has; null for plain text or one not offered here */
  align: RubyAlign | null
  /** an empty reading takes the guide off */
  onApply: (edit: RubyEdit) => void
  onClose: () => void
}) {
  const { t, lang } = useI18n()
  const keys = useModalKeys(onClose)
  const [font, setFont] = useState(set.font)
  const [size, setSize] = useState(String(set.sizePt))
  const [offset, setOffset] = useState(String(set.offsetPt))
  const fonts = useMemo(() => {
    const list = [...fontFamiliesFor(lang)]
    return set.font && !list.includes(set.font) ? [set.font, ...list] : list
  }, [lang, set.font])
  // half-point steps, within what Word's Size and distance boxes take
  const points = (text: string, min: number, max: number, fallback: number): number => {
    const n = Math.round(Number(text) * 2) / 2
    return Number.isFinite(n) && text.trim() !== '' ? Math.min(max, Math.max(min, n)) : fallback
  }
  const sizePt = points(size, 1, 100, set.sizePt)
  const offsetPt = points(offset, 0, 100, set.offsetPt)
  const restyled: Partial<RubyReading> = {
    ...(font !== set.font ? { font } : {}),
    ...(sizePt !== set.sizePt ? { sizePt } : {}),
    ...(offsetPt !== set.offsetPt ? { offsetPt } : {}),
  }
  const hasRestyle = Object.keys(restyled).length > 0
  const [reading, setReading] = useState(rt)
  const [baseText, setBaseText] = useState(base)
  // a guide with a placement not offered here keeps it until another is picked
  const kept = rt !== '' && align === null
  const [placement, setPlacement] = useState<RubyAlign | ''>(align ?? (kept ? '' : 'center'))
  const changed =
    reading.trim() !== rt || baseText !== base || placement !== (align ?? '') || hasRestyle
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog phonetic-guide-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          // nothing to put a reading on: the dialog stays, with what was typed
          if (!baseText.trim()) return
          onClose()
          if (!changed) return
          onApply({
            rt: reading,
            ...(baseText !== base ? { base: baseText } : {}),
            ...(placement ? { align: placement } : {}),
            ...(hasRestyle ? { reading: restyled } : {}),
          })
        }}
      >
        <h2>{t('ribbonPhoneticGuide')}</h2>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonPgBase')}
            <input
              type="text"
              name="base"
              value={baseText}
              maxLength={RUBY_BASE_MAX}
              spellCheck={false}
              onChange={(e) => setBaseText(e.target.value)}
            />
          </label>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonPgRuby')}
            <input
              type="text"
              value={reading}
              autoFocus
              spellCheck={false}
              onChange={(e) => setReading(e.target.value)}
            />
          </label>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('appAlignment')}
            <select
              name="align"
              value={placement}
              onChange={(e) => setPlacement(e.target.value as RubyAlign)}
            >
              {kept && <option value="" />}
              {RUBY_ALIGNS.map((value) => {
                const label = ALIGN_LABELS[value]
                return (
                  <option key={value} value={value}>
                    {typeof label === 'string' ? t(label) : label.literal}
                  </option>
                )
              })}
            </select>
          </label>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('appFontTabFont')}
            <select name="font" value={font} onChange={(e) => setFont(e.target.value)}>
              <option value="" />
              {fonts.map((f) => (
                <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t('appFontSizeLabel')}
            <input
              type="number"
              name="size"
              min={1}
              max={100}
              step={0.5}
              value={size}
              onChange={(e) => setSize(e.target.value)}
            />
          </label>
          <label>
            {t('ribbonPgDistance')}
            <input
              type="number"
              name="offset"
              min={0}
              max={100}
              step={0.5}
              value={offset}
              onChange={(e) => setOffset(e.target.value)}
            />
          </label>
        </div>
        <div className="phonetic-guide-preview">
          <ruby className="doc-ruby" data-ruby-align={placement || undefined}>
            {placement === 'right' ? <span className="doc-ruby-base">{baseText}</span> : baseText}
            <rt style={font ? { fontFamily: cssFontFamily(font) } : undefined}>{reading}</rt>
          </ruby>
        </div>
        <div className="modal-actions">
          <button
            type="button"
            disabled={!rt}
            onClick={() => {
              onClose()
              onApply({ rt: '' })
            }}
          >
            {t('ribbonPgClear')}
          </button>
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary" disabled={!baseText.trim()}>
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}
