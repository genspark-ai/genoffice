import { useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { Command } from '@tiptap/pm/state'
import { useI18n } from '../i18n/locale'
import {
  canShiftCells,
  deleteCells,
  editableOrToast,
  insertRowsOrColumns,
  splitCellsInto,
  type DeleteCellsMode,
} from '../editor/table-ops'
import { useModalKeys } from './modal-keys'
import { LengthInput } from './LengthInput'
import { TWIPS_PER_CM } from '../list-presets'
import { applyTableValue, tableValueOf, type TableValue } from './TablePropertiesDialog'

export type TableDialogKind =
  'properties' | 'splitCells' | 'insertCells' | 'deleteCells' | 'cellMargins'

/** Word's small table dialogs: Split Cells… / Insert Cells… / Delete Cells… */

interface DialogProps {
  editor: Editor
  onClose: () => void
}

function runAndClose(editor: Editor, command: Command, onClose: () => void) {
  if (editableOrToast(editor)) {
    editor.view.focus()
    command(editor.state, editor.view.dispatch)
  }
  onClose()
}

function Frame({
  title,
  onClose,
  onOk,
  okDisabled,
  children,
  className,
}: {
  title: string
  onClose: () => void
  onOk: () => void
  okDisabled?: boolean
  children: React.ReactNode
  className: string
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className={`modal gs-form table-small-dialog ${className}`}
        onSubmit={(e) => {
          e.preventDefault()
          if (!okDisabled) onOk()
        }}
      >
        <h2>{title}</h2>
        {children}
        <div className="modal-actions">
          <button type="button" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button type="submit" className="primary" disabled={okDisabled}>
            {t('ribbonOk')}
          </button>
        </div>
      </form>
    </div>
  )
}

export function SplitCellsDialog({ editor, onClose }: DialogProps) {
  const { t } = useI18n()
  const [cols, setCols] = useState(2)
  const [rows, setRows] = useState(1)
  const [mergeFirst, setMergeFirst] = useState(true)
  const num = (value: number, set: (n: number) => void, label: string) => (
    <label>
      {label}
      <input
        type="number"
        min={1}
        max={50}
        value={value}
        onChange={(e) => set(Math.min(50, Math.max(1, Math.round(Number(e.target.value) || 1))))}
      />
    </label>
  )
  return (
    <Frame
      className="split-cells-dialog"
      title={t('ribbonSplitCellsDialog')}
      onClose={onClose}
      okDisabled={rows === 1 && cols === 1}
      onOk={() => runAndClose(editor, splitCellsInto(rows, cols, mergeFirst), onClose)}
    >
      <div className="table-dialog-grid">
        {num(cols, setCols, t('ribbonTableColsLabel'))}
        {num(rows, setRows, t('ribbonTableRowsLabel'))}
      </div>
      <label className="table-dialog-check">
        <input
          type="checkbox"
          checked={mergeFirst}
          onChange={(e) => setMergeFirst(e.target.checked)}
        />
        {t('ribbonMergeBeforeSplit')}
      </label>
    </Frame>
  )
}

type InsertWhere = 'above' | 'below' | 'left' | 'right'

export function InsertCellsDialog({
  editor,
  onClose,
  sectionContentWidthPx,
}: DialogProps & { sectionContentWidthPx: number }) {
  const { t } = useI18n()
  const [where, setWhere] = useState<InsertWhere>('below')
  const [count, setCount] = useState(1)
  const options: Array<[InsertWhere, string]> = [
    ['above', t('ribbonInsertAbove')],
    ['below', t('ribbonInsertBelow')],
    ['left', t('ribbonInsertLeft')],
    ['right', t('ribbonInsertRight')],
  ]
  return (
    <Frame
      className="insert-cells-dialog"
      title={t('ribbonInsertCells')}
      onClose={onClose}
      onOk={() =>
        runAndClose(editor, insertRowsOrColumns(count, where, sectionContentWidthPx), onClose)
      }
    >
      <div className="table-dialog-radios" role="radiogroup">
        {options.map(([value, label]) => (
          <label key={value}>
            <input
              type="radio"
              name="insert-cells-where"
              checked={where === value}
              onChange={() => setWhere(value)}
            />
            {label}
          </label>
        ))}
      </div>
      <label>
        {t('ribbonInsertCount')}
        <input
          type="number"
          min={1}
          max={100}
          value={count}
          onChange={(e) =>
            setCount(Math.min(100, Math.max(1, Math.round(Number(e.target.value) || 1))))
          }
        />
      </label>
    </Frame>
  )
}

export function DeleteCellsDialog({ editor, onClose }: DialogProps) {
  const { t } = useI18n()
  const shiftAllowed = canShiftCells(editor.state)
  const [mode, setMode] = useState<DeleteCellsMode>(shiftAllowed ? 'shiftLeft' : 'row')
  const options: Array<[DeleteCellsMode, string, boolean]> = [
    ['shiftLeft', t('ribbonDeleteShiftLeft'), shiftAllowed],
    ['shiftUp', t('ribbonDeleteShiftUp'), shiftAllowed],
    ['row', t('ribbonDeleteEntireRow'), true],
    ['column', t('ribbonDeleteEntireColumn'), true],
  ]
  return (
    <Frame
      className="delete-cells-dialog"
      title={t('ribbonDeleteCells')}
      onClose={onClose}
      onOk={() => runAndClose(editor, deleteCells(mode), onClose)}
    >
      <div className="table-dialog-radios" role="radiogroup">
        {options.map(([value, label, enabled]) => (
          <label key={value} className={enabled ? '' : 'disabled'}>
            <input
              type="radio"
              name="delete-cells-mode"
              disabled={!enabled}
              checked={mode === value}
              onChange={() => setMode(value)}
            />
            {label}
          </label>
        ))}
      </div>
      {!shiftAllowed && <p className="modal-desc">{t('ribbonShiftNeedsPlainGrid')}</p>}
    </Frame>
  )
}

type MarginKey = 'marginTopCm' | 'marginBottomCm' | 'marginLeftCm' | 'marginRightCm'

/** Table Layout ▸ Cell Margins: Word's Table Options dialog (default cell margins only) */
export function CellMarginsDialog({
  editor,
  onClose,
  sectionContentWidthPx,
}: DialogProps & { sectionContentWidthPx: number }) {
  const { t } = useI18n()
  const [value, setValue] = useState<TableValue>(() =>
    tableValueOf(editor.getAttributes('docTable')),
  )
  const fields: Array<[MarginKey, string]> = [
    ['marginTopCm', t('ribbonMarginTop')],
    ['marginBottomCm', t('ribbonMarginBottom')],
    ['marginLeftCm', t('ribbonMarginLeft')],
    ['marginRightCm', t('ribbonMarginRight')],
  ]
  return (
    <Frame
      className="cell-margins-dialog"
      title={t('ribbonTableOptions')}
      onClose={onClose}
      onOk={() => {
        if (editableOrToast(editor)) {
          applyTableValue(editor, value, new Set(['margins']), sectionContentWidthPx)
        }
        onClose()
      }}
    >
      <p className="modal-desc">{t('ribbonCellMargins')}</p>
      <div className="table-properties-grid">
        {fields.map(([key, label]) => (
          <label key={key}>
            {label}
            <span>
              <LengthInput
                value={Math.round(value[key] * TWIPS_PER_CM)}
                min={0}
                max={50 * TWIPS_PER_CM}
                ariaLabel={label}
                live
                onCommit={(twips) =>
                  setValue((prev) => ({ ...prev, [key]: (twips ?? 0) / TWIPS_PER_CM }))
                }
              />
            </span>
          </label>
        ))}
      </div>
    </Frame>
  )
}
