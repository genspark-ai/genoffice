import { useState } from 'react'

import { useI18n, type StringKey } from './i18n/locale'
import { useModalDialog } from './modal-dialog'
import {
  PASTE_SPECIAL_OPERATIONS,
  PASTE_SPECIAL_TYPES,
  pasteSpecialAvailability,
  type PasteSpecialOperation,
  type PasteSpecialOptions,
  type PasteSpecialSource,
  type PasteSpecialType,
} from './paste-special'

const TYPE_LABELS: Readonly<Record<PasteSpecialType, StringKey>> = {
  all: 'dlgPasteSpecialAll',
  formulas: 'dlgPasteSpecialFormulas',
  values: 'dlgPasteSpecialValues',
  formats: 'dlgPasteSpecialFormats',
  comments: 'dlgPasteSpecialComments',
  validation: 'dlgPasteSpecialValidation',
  'all-source-theme': 'dlgPasteSpecialAllSourceTheme',
  'all-except-borders': 'dlgPasteSpecialAllExceptBorders',
  'col-widths': 'dlgPasteSpecialColWidths',
  'formulas-numfmt': 'dlgPasteSpecialFormulasNumfmt',
  'values-numfmt': 'dlgPasteSpecialValuesNumfmt',
  'all-merge-cf': 'dlgPasteSpecialAllMergeCf',
}

const OPERATION_LABELS: Readonly<Record<PasteSpecialOperation, StringKey>> = {
  none: 'dlgPasteSpecialOpNone',
  add: 'dlgPasteSpecialOpAdd',
  subtract: 'dlgPasteSpecialOpSubtract',
  multiply: 'dlgPasteSpecialOpMultiply',
  divide: 'dlgPasteSpecialOpDivide',
}

/// Excel's Paste Special (Ctrl+Alt+V): paste type, arithmetic operation,
/// skip blanks / transpose, and Paste Link. Choices the copy source cannot
/// serve (a cut, text from another app) are disabled rather than hidden.
export function PasteSpecialDialog({
  source,
  onApply,
  onClose,
}: {
  readonly source: PasteSpecialSource
  /// Resolves to an error message, or null on success.
  readonly onApply: (options: PasteSpecialOptions) => Promise<string | null>
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const modal = useModalDialog(onClose)
  const availability = pasteSpecialAvailability(source)
  const [type, setType] = useState<PasteSpecialType>('all')
  const [operation, setOperation] = useState<PasteSpecialOperation>('none')
  const [skipBlanks, setSkipBlanks] = useState(false)
  const [transpose, setTranspose] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)

  const canPaste = availability.types.has(type) && !busy
  const run = (link: boolean): void => {
    setBusy(true)
    void onApply({ type, operation, skipBlanks, transpose, link }).then((failure) => {
      setBusy(false)
      setError(failure)
      if (failure === null) onClose()
    })
  }

  const radio = (
    name: string,
    value: string,
    checked: boolean,
    enabled: boolean,
    label: string,
    onPick: () => void,
  ): React.JSX.Element => (
    <label key={value} className={`dialog-check${enabled ? '' : ' is-disabled'}`}>
      <input type="radio" name={name} checked={checked} disabled={!enabled} onChange={onPick} />
      {label}
    </label>
  )

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        {...modal}
        className="format-cells-dialog paste-special-dialog"
        role="dialog"
        aria-label={t('dlgPasteSpecialTitle')}
        onClick={(event) => event.stopPropagation()}
        onKeyDown={(event) => {
          if (event.key === 'Enter' && !(event.target instanceof HTMLButtonElement)) {
            event.preventDefault()
            if (canPaste) run(false)
            return
          }
          modal.onKeyDown(event)
        }}
      >
        <header>{t('dlgPasteSpecialTitle')}</header>
        <section className="dialog-body paste-special-body">
          <fieldset className="paste-special-group paste-special-types">
            <legend>{t('dlgPasteSpecialPaste')}</legend>
            {PASTE_SPECIAL_TYPES.map((option) =>
              radio(
                'paste-special-type',
                option,
                type === option,
                availability.types.has(option),
                t(TYPE_LABELS[option]),
                () => setType(option),
              ),
            )}
          </fieldset>
          <fieldset className="paste-special-group paste-special-operations">
            <legend>{t('dlgPasteSpecialOperation')}</legend>
            {PASTE_SPECIAL_OPERATIONS.map((option) =>
              radio(
                'paste-special-operation',
                option,
                operation === option,
                availability.operation || option === 'none',
                t(OPERATION_LABELS[option]),
                () => setOperation(option),
              ),
            )}
          </fieldset>
          <div className="paste-special-flags">
            <label className={`dialog-check${availability.skipBlanks ? '' : ' is-disabled'}`}>
              <input
                type="checkbox"
                checked={skipBlanks}
                disabled={!availability.skipBlanks}
                onChange={(event) => setSkipBlanks(event.target.checked)}
              />
              {t('dlgPasteSpecialSkipBlanks')}
            </label>
            <label className={`dialog-check${availability.transpose ? '' : ' is-disabled'}`}>
              <input
                type="checkbox"
                checked={transpose}
                disabled={!availability.transpose}
                onChange={(event) => setTranspose(event.target.checked)}
              />
              {t('dlgPasteSpecialTranspose')}
            </label>
          </div>
          {source.kind === 'none' && <p className="dialog-note">{t('dlgPasteSpecialNothing')}</p>}
          {source.kind === 'text' && <p className="dialog-note">{t('dlgPasteSpecialExternal')}</p>}
          {source.kind === 'internal' && source.cut && (
            <p className="dialog-note">{t('dlgPasteSpecialCut')}</p>
          )}
          {error && (
            <p className="dialog-note dialog-error" role="alert">
              {error}
            </p>
          )}
        </section>
        <div className="dialog-actions">
          <button
            className="secondary"
            disabled={!availability.link || busy}
            onClick={() => run(true)}
          >
            {t('dlgPasteSpecialLink')}
          </button>
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary-action" disabled={!canPaste} onClick={() => run(false)}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
