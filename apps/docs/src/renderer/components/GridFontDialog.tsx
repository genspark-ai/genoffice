import { useState } from 'react'
import { Dropdown, type DropdownOption } from '@genoffice/ui'
import { useI18n } from '../i18n/locale'
import { fontFamiliesFor } from '../font-list'
import { fontSizeLabel, fontSizeOptions, parseFontSize } from '../font-sizes'
import { cssFontFamily } from '../text-layout/font-rules'
import { useModalKeys } from './modal-keys'

/** the font of the default paragraph style, which the document grid counts its cells from */
export interface GridFont {
  latin?: string
  eastAsia?: string
  sizeHalfPoints?: number
  bold?: boolean
  italic?: boolean
}

const STYLES = [
  ['regular', 'appFontRegular'],
  ['italic', 'appFontItalic'],
  ['bold', 'appFontBold'],
  ['boldItalic', 'appFontBoldItalic'],
] as const

/** Document Grid > Set Font…: only what the user changes is handed back */
export function GridFontDialog({
  font,
  onApply,
  onClose,
}: {
  font: Required<GridFont>
  onApply: (changed: GridFont) => void
  onClose: () => void
}) {
  const { t, lang } = useI18n()
  const keys = useModalKeys(onClose)
  const [value, setValue] = useState(font)
  const families = fontFamiliesFor(lang)
  const sizes = fontSizeOptions(lang)
  const sizePt = value.sizeHalfPoints / 2
  const style = value.bold
    ? value.italic
      ? 'boldItalic'
      : 'bold'
    : value.italic
      ? 'italic'
      : 'regular'
  const facePicker = (key: 'latin' | 'eastAsia', label: string) => (
    <label>
      {label}
      <Dropdown
        value={value[key]}
        ariaLabel={label}
        options={[
          ...(families.includes(value[key]) || !value[key] ? [] : [value[key]]),
          ...families,
        ].map((f): DropdownOption => ({
          value: f,
          label: f,
          render: <span style={{ fontFamily: cssFontFamily(f) }}>{f}</span>,
        }))}
        onPick={(f) => setValue((v) => ({ ...v, [key]: f }))}
      />
    </label>
  )
  const submit = () => {
    const changed: GridFont = {}
    if (value.latin !== font.latin) changed.latin = value.latin
    if (value.eastAsia !== font.eastAsia) changed.eastAsia = value.eastAsia
    if (value.sizeHalfPoints !== font.sizeHalfPoints) changed.sizeHalfPoints = value.sizeHalfPoints
    if (value.bold !== font.bold) changed.bold = value.bold
    if (value.italic !== font.italic) changed.italic = value.italic
    onApply(changed)
    onClose()
  }
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal font-dialog grid-font-dialog"
        onSubmit={(e) => {
          e.preventDefault()
          e.stopPropagation()
          submit()
        }}
      >
        <h2>{t('appFontDialogTitle')}</h2>
        <div className="font-dialog-row">
          {facePicker('latin', t('ribbonFontLatin'))}
          {facePicker('eastAsia', t('ribbonFontEastAsia'))}
          <label>
            {t('appFontStyleLabel')}
            <Dropdown
              value={style}
              ariaLabel={t('appFontStyleLabel')}
              options={STYLES.map(([key, nameKey]) => ({ value: key, label: t(nameKey) }))}
              onPick={(s) =>
                setValue((v) => ({
                  ...v,
                  bold: s === 'bold' || s === 'boldItalic',
                  italic: s === 'italic' || s === 'boldItalic',
                }))
              }
            />
          </label>
          <label>
            {t('appFontSizeLabel')}
            <Dropdown
              value={fontSizeLabel(sizePt, lang)}
              ariaLabel={t('appFontSizeLabel')}
              options={[
                ...(sizes.some((o) => o.name === fontSizeLabel(sizePt, lang))
                  ? []
                  : [fontSizeLabel(sizePt, lang)]),
                ...sizes.map((o) => o.name),
              ].map((s) => ({ value: s, label: s }))}
              onPick={(s) => {
                const pt = parseFontSize(s, lang)
                if (pt !== null) setValue((v) => ({ ...v, sizeHalfPoints: Math.round(pt * 2) }))
              }}
            />
          </label>
        </div>
        <div className="modal-actions">
          <button type="button" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button type="submit" className="primary">
            {t('ribbonOk')}
          </button>
        </div>
      </form>
    </div>
  )
}
