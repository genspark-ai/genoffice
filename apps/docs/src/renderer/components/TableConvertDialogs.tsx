import { useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { TableAutoFitMode } from '@genoffice/docx-engine'
import {
  MAX_CONVERT_COLS,
  convertTableToText,
  convertTextToTable,
  convertedRows,
  suggestedColumns,
  textToTableScope,
  type ConvertSeparator,
} from '../editor/table-convert'
import {
  FORMULA_FUNCTIONS,
  FORMULA_NUMBER_FORMATS,
  formulaFieldAt,
  formulaInstr,
  insertTableFormula,
  parseFormulaInstr,
  suggestedFormula,
  type FormulaRefusal,
} from '../editor/table-formula'
import { collectBookmarks, sortBookmarks } from '../editor/cross-reference'
import { editableOrToast } from '../editor/table-ops'
import { useI18n, type StringKey } from '../i18n/locale'
import { FormDialog } from './MenuDialogs'

/** Word's Table > Convert dialogs and its Formula dialog */

const SEPARATORS = [
  { id: 'paragraph', text: null },
  { id: 'tab', text: '\t' },
  { id: 'comma', text: ',' },
  { id: 'other', text: null },
] as const

type SeparatorId = (typeof SEPARATORS)[number]['id']

function separatorOf(id: SeparatorId, other: string): ConvertSeparator {
  if (id === 'paragraph') return { kind: 'paragraph' }
  return { kind: 'text', text: SEPARATORS.find((entry) => entry.id === id)?.text ?? other }
}

function SeparatorChoices({
  caption,
  paragraphLabel,
  value,
  other,
  onPick,
  onOther,
}: {
  caption: string
  paragraphLabel: StringKey
  value: SeparatorId
  other: string
  onPick: (id: SeparatorId) => void
  onOther: (text: string) => void
}) {
  const { t } = useI18n()
  const labels: Record<SeparatorId, StringKey> = {
    paragraph: paragraphLabel,
    tab: 'ribbonSortTabs',
    comma: 'ribbonSortCommas',
    other: 'ribbonSortOther',
  }
  return (
    <fieldset>
      <legend>{caption}</legend>
      <div className="sort-choices">
        {SEPARATORS.map((entry) => (
          <label key={entry.id} className="tco-check">
            <input
              type="radio"
              name="separator"
              value={entry.id}
              checked={value === entry.id}
              onChange={() => onPick(entry.id)}
            />
            {t(labels[entry.id])}
          </label>
        ))}
        <input
          name="other"
          className="sort-other"
          aria-label={t('ribbonSortOther')}
          maxLength={1}
          value={other}
          disabled={value !== 'other'}
          onChange={(e) => onOther(e.target.value)}
        />
      </div>
    </fieldset>
  )
}

interface DialogProps {
  editor: Editor
  onClose: () => void
}

const AUTO_FIT: ReadonlyArray<{ mode: TableAutoFitMode; label: StringKey }> = [
  { mode: 'fixed', label: 'ribbonFixedColumnWidth' },
  { mode: 'contents', label: 'ribbonAutoFitContents' },
  { mode: 'window', label: 'ribbonAutoFitWindow' },
]

/** tabs part the text when it has any, else commas, else each paragraph is a cell: Word's guess */
function guessSeparator(texts: readonly string[]): SeparatorId {
  if (texts.some((text) => text.includes('\t'))) return 'tab'
  if (texts.some((text) => text.includes(','))) return 'comma'
  return 'paragraph'
}

export function ConvertTextToTableDialog({
  editor,
  onClose,
  sectionContentWidthPx,
}: DialogProps & { sectionContentWidthPx: number }) {
  const { t } = useI18n()
  const [texts] = useState(() => textToTableScope(editor.state) ?? [])
  const [separatorId, setSeparatorId] = useState<SeparatorId>(() => guessSeparator(texts))
  const [other, setOther] = useState('-')
  const [cols, setCols] = useState(() =>
    suggestedColumns(texts, separatorOf(guessSeparator(texts), '-')),
  )
  const [autoFit, setAutoFit] = useState<TableAutoFitMode>('fixed')
  const separator = separatorOf(separatorId, other)
  // the column count follows the separator until the user sets one, as in Word
  const reseparate = (id: SeparatorId, text: string) => {
    setSeparatorId(id)
    setOther(text)
    setCols(suggestedColumns(texts, separatorOf(id, text)))
  }
  return (
    <FormDialog
      title={t('ribbonConvertTextToTableTitle')}
      name="sort-dialog convert-text-dialog"
      okDisabled={texts.length === 0}
      onClose={onClose}
      onSubmit={() => {
        onClose()
        if (!editableOrToast(editor)) return
        editor.view.focus()
        convertTextToTable(editor, { cols, separator, autoFit }, sectionContentWidthPx)
      }}
    >
      <fieldset>
        <legend>{t('ribbonConvertTableSize')}</legend>
        <div className="table-dialog-grid">
          <label>
            {t('ribbonTableColsLabel')}
            <input
              type="number"
              name="cols"
              min={1}
              max={MAX_CONVERT_COLS}
              value={cols}
              onChange={(e) =>
                setCols(
                  Math.min(MAX_CONVERT_COLS, Math.max(1, Math.round(Number(e.target.value) || 1))),
                )
              }
            />
          </label>
          <label>
            {t('ribbonTableRowsLabel')}
            <input
              type="number"
              name="rows"
              value={convertedRows(texts, separator, cols)}
              // always-disabled: the row count follows from the text and the columns
              disabled
              readOnly
            />
          </label>
        </div>
      </fieldset>
      <fieldset>
        <legend>{t('ribbonConvertAutoFitBehavior')}</legend>
        <div className="sort-choices">
          {AUTO_FIT.map((entry) => (
            <label key={entry.mode} className="tco-check">
              <input
                type="radio"
                name="autoFit"
                value={entry.mode}
                checked={autoFit === entry.mode}
                onChange={() => setAutoFit(entry.mode)}
              />
              {t(entry.label)}
            </label>
          ))}
        </div>
      </fieldset>
      <SeparatorChoices
        caption={t('ribbonConvertSeparateAt')}
        paragraphLabel="appStatParagraphs"
        value={separatorId}
        other={other}
        onPick={(id) => reseparate(id, other)}
        onOther={(text) => reseparate('other', text)}
      />
    </FormDialog>
  )
}

export function ConvertTableToTextDialog({ editor, onClose }: DialogProps) {
  const { t } = useI18n()
  const [separatorId, setSeparatorId] = useState<SeparatorId>('paragraph')
  const [other, setOther] = useState('-')
  const [nested, setNested] = useState(true)
  return (
    <FormDialog
      title={t('ribbonConvertTableToTextTitle')}
      name="sort-dialog convert-table-dialog"
      onClose={onClose}
      onSubmit={() => {
        onClose()
        if (!editableOrToast(editor)) return
        editor.view.focus()
        convertTableToText(editor, { separator: separatorOf(separatorId, other), nested })
      }}
    >
      <SeparatorChoices
        caption={t('ribbonConvertSeparateWith')}
        paragraphLabel="ribbonConvertParagraphMarks"
        value={separatorId}
        other={other}
        onPick={setSeparatorId}
        onOther={setOther}
      />
      <label className="tco-check">
        <input
          type="checkbox"
          name="nested"
          checked={nested}
          onChange={(e) => setNested(e.target.checked)}
        />
        {t('ribbonConvertNestedTables')}
      </label>
    </FormDialog>
  )
}

const FORMULA_REFUSALS: Record<FormulaRefusal, StringKey> = {
  notInCell: 'ribbonFormulaNotInCell',
  unsupported: 'ribbonFormulaUnsupported',
  outsideTable: 'ribbonFormulaOutsideTable',
}

export function FormulaDialog({ editor, onClose }: DialogProps) {
  const { t } = useI18n()
  const [initial] = useState(() => {
    const existing = formulaFieldAt(editor.state)
    const parsed = existing ? parseFormulaInstr(existing.instr) : null
    return parsed
      ? { formula: `=${parsed.expr}`, format: parsed.format ?? '' }
      : { formula: suggestedFormula(editor.state), format: '' }
  })
  const [formula, setFormula] = useState(initial.formula)
  const [format, setFormat] = useState(initial.format)
  const [refusal, setRefusal] = useState<FormulaRefusal | null>(null)
  const empty = formula.trim().replace(/^=/, '').trim() === ''
  const [bookmarks] = useState(() =>
    sortBookmarks(collectBookmarks(editor.state.doc), 'name').map((entry) => entry.name),
  )
  return (
    <FormDialog
      title={t('ribbonFormulaTitle')}
      name="sort-dialog formula-dialog"
      okDisabled={empty}
      onClose={onClose}
      onSubmit={() => {
        if (empty || !editableOrToast(editor)) return
        const refused = insertTableFormula(editor, formulaInstr(formula, format))
        setRefusal(refused)
        if (refused) return
        onClose()
        editor.view.focus()
      }}
    >
      <label>
        {t('ribbonFormulaTitle')}
        <input
          name="formula"
          autoFocus
          spellCheck={false}
          value={formula}
          onChange={(e) => {
            setFormula(e.target.value)
            setRefusal(null)
          }}
        />
      </label>
      <label>
        {t('ribbonFormulaNumberFormat')}
        <input
          name="format"
          list="formula-number-formats"
          spellCheck={false}
          value={format}
          onChange={(e) => setFormat(e.target.value)}
        />
        <datalist id="formula-number-formats">
          {FORMULA_NUMBER_FORMATS.map((picture) => (
            <option key={picture} value={picture} />
          ))}
        </datalist>
      </label>
      <label>
        {t('ribbonFormulaPasteFunction')}
        <select
          name="pasteFunction"
          value=""
          onChange={(e) => {
            if (!e.target.value) return
            setFormula((prev) => `${prev.trim() || '='}${e.target.value}()`)
            setRefusal(null)
          }}
        >
          <option value="" />
          {FORMULA_FUNCTIONS.map((name) => (
            <option key={name} value={name}>
              {name}
            </option>
          ))}
        </select>
      </label>
      <label>
        {t('ribbonFormulaPasteBookmark')}
        <select
          name="pasteBookmark"
          value=""
          onChange={(e) => {
            if (!e.target.value) return
            setFormula((prev) => `${prev.trim() || '='}${e.target.value}`)
            setRefusal(null)
          }}
        >
          <option value="" />
          {bookmarks.map((name) => (
            <option key={name} value={name}>
              {name}
            </option>
          ))}
        </select>
      </label>
      {refusal && (
        <div className="modal-error" role="alert">
          {t(FORMULA_REFUSALS[refusal])}
        </div>
      )}
    </FormDialog>
  )
}
