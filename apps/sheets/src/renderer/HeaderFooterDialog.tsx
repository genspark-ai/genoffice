import { useRef, useState } from 'react'

import { useI18n, type StringKey } from './i18n/locale'

import type { HeaderFooterParts } from './edit-journal'

/// Excel's Insert → Header & Footer as a dialog (the app has no Page Layout
/// view): three sections each for the printed header and footer. Field codes
/// travel verbatim — the quick-insert row types them into the focused
/// section. OK hands both halves to App, which journals them as page-setup
/// state; all-empty sections clear that half (null).

export interface HeaderFooterResult {
  readonly header: HeaderFooterParts | null
  readonly footer: HeaderFooterParts | null
}

type SectionKey =
  | 'header-left'
  | 'header-center'
  | 'header-right'
  | 'footer-left'
  | 'footer-center'
  | 'footer-right'

const SECTION_KEYS: readonly SectionKey[] = [
  'header-left',
  'header-center',
  'header-right',
  'footer-left',
  'footer-center',
  'footer-right',
]

const SECTION_LABELS: Record<SectionKey, StringKey> = {
  'header-left': 'dlgHfHeaderLeft',
  'header-center': 'dlgHfHeaderCenter',
  'header-right': 'dlgHfHeaderRight',
  'footer-left': 'dlgHfFooterLeft',
  'footer-center': 'dlgHfFooterCenter',
  'footer-right': 'dlgHfFooterRight',
}

const FIELD_CODES: readonly { readonly code: string; readonly labelKey: StringKey }[] = [
  { code: '&P', labelKey: 'dlgHfPageNumber' },
  { code: '&N', labelKey: 'dlgHfPageCount' },
  { code: '&D', labelKey: 'dlgHfDate' },
  { code: '&T', labelKey: 'dlgHfTime' },
  { code: '&F', labelKey: 'dlgHfFileName' },
  { code: '&A', labelKey: 'dlgHfSheetName' },
]

/// Non-empty sections only; all empty means "clear this half" (null).
function toParts(left: string, center: string, right: string): HeaderFooterParts | null {
  if (left === '' && center === '' && right === '') return null
  return {
    ...(left === '' ? {} : { left }),
    ...(center === '' ? {} : { center }),
    ...(right === '' ? {} : { right }),
  }
}

export function HeaderFooterDialog({
  initialHeader,
  initialFooter,
  onApply,
  onClose,
}: {
  readonly initialHeader: HeaderFooterParts | null
  readonly initialFooter: HeaderFooterParts | null
  /// Returns an error message, or null on success.
  readonly onApply: (result: HeaderFooterResult) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [values, setValues] = useState<Record<SectionKey, string>>({
    'header-left': initialHeader?.left ?? '',
    'header-center': initialHeader?.center ?? '',
    'header-right': initialHeader?.right ?? '',
    'footer-left': initialFooter?.left ?? '',
    'footer-center': initialFooter?.center ?? '',
    'footer-right': initialFooter?.right ?? '',
  })
  const [error, setError] = useState<string | null>(null)
  const focusedKey = useRef<SectionKey>('header-center')
  const inputs = useRef(new Map<SectionKey, HTMLInputElement>())

  const insertCode = (code: string): void => {
    const key = focusedKey.current
    const input = inputs.current.get(key)
    const value = values[key]
    const start = input?.selectionStart ?? value.length
    const end = input?.selectionEnd ?? value.length
    setValues((prev) => ({ ...prev, [key]: value.slice(0, start) + code + value.slice(end) }))
    // Refocus so consecutive quick-inserts land in the same section.
    requestAnimationFrame(() => {
      input?.focus()
      input?.setSelectionRange(start + code.length, start + code.length)
    })
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={t('dlgHfTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgHfTitle')}</header>
        <div className="dialog-grid" style={{ gridTemplateColumns: '1fr 1fr 1fr' }}>
          {SECTION_KEYS.map((key) => (
            <label key={key}>
              {t(SECTION_LABELS[key])}
              <input
                value={values[key]}
                ref={(element) => {
                  if (element) inputs.current.set(key, element)
                  else inputs.current.delete(key)
                }}
                onFocus={() => {
                  focusedKey.current = key
                }}
                onChange={(event) => {
                  const next = event.target.value
                  setValues((prev) => ({ ...prev, [key]: next }))
                }}
              />
            </label>
          ))}
          <div className="dialog-span" style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
            {FIELD_CODES.map((field) => (
              <button
                key={field.code}
                className="secondary"
                data-tip={t('dlgHfInsertCode', { code: field.code })}
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => insertCode(field.code)}
              >
                {t(field.labelKey)} {field.code}
              </button>
            ))}
          </div>
          <p className="dialog-note dialog-span">{t('dlgHfNote')}</p>
        </div>
        {error && (
          <p className="dialog-note" role="alert">
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button
            className="primary-action"
            onClick={() => {
              const failure = onApply({
                header: toParts(
                  values['header-left'],
                  values['header-center'],
                  values['header-right'],
                ),
                footer: toParts(
                  values['footer-left'],
                  values['footer-center'],
                  values['footer-right'],
                ),
              })
              setError(failure)
              if (failure === null) onClose()
            }}
          >
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
