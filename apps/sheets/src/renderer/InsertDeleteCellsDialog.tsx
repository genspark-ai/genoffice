import { useState } from 'react'

import { useI18n } from './i18n/locale'
import { CELLS_CHOICES, DEFAULT_CELLS_CHOICE } from './insert-delete-cells'
import type { CellsChoice, CellsMode } from './insert-delete-cells'
import { useModalDialog } from './modal-dialog'

const LABEL_KEYS = {
  insert: {
    'shift-horizontal': 'dlgCellsShiftRight',
    'shift-vertical': 'dlgCellsShiftDown',
    'entire-row': 'dlgCellsEntireRow',
    'entire-column': 'dlgCellsEntireColumn',
  },
  delete: {
    'shift-horizontal': 'dlgCellsShiftLeft',
    'shift-vertical': 'dlgCellsShiftUp',
    'entire-row': 'dlgCellsEntireRow',
    'entire-column': 'dlgCellsEntireColumn',
  },
} as const

export function InsertDeleteCellsDialog({
  mode,
  onApply,
  onClose,
}: {
  readonly mode: CellsMode
  readonly onApply: (choice: CellsChoice) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const modal = useModalDialog(onClose)
  const [choice, setChoice] = useState<CellsChoice>(DEFAULT_CELLS_CHOICE)
  const title = t(mode === 'insert' ? 'dlgInsertCellsTitle' : 'dlgDeleteCellsTitle')

  const apply = (): void => {
    onApply(choice)
    onClose()
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        {...modal}
        className="format-cells-dialog cells-dialog"
        role="dialog"
        aria-label={title}
        onClick={(event) => event.stopPropagation()}
        onKeyDown={(event) => {
          // Buttons keep their own Enter (Cancel must not apply)
          if (event.key === 'Enter' && !(event.target instanceof HTMLButtonElement)) {
            event.preventDefault()
            apply()
            return
          }
          modal.onKeyDown(event)
        }}
      >
        <header>{title}</header>
        <section className="dialog-body cells-dialog-body" role="radiogroup" aria-label={title}>
          {CELLS_CHOICES.map((option) => (
            <label key={option} className="dialog-check">
              <input
                type="radio"
                name="insert-delete-cells"
                autoFocus={option === DEFAULT_CELLS_CHOICE}
                checked={choice === option}
                onChange={() => setChoice(option)}
              />
              {t(LABEL_KEYS[mode][option])}
            </label>
          ))}
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary-action" onClick={apply}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
