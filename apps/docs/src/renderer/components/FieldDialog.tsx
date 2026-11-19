import { useState } from 'react'
import { DATE_TIME_PICTURES } from '../editor/date-picture'
import {
  NUMBER_FORMAT_SWITCHES,
  NUMERIC_PICTURES,
  TEXT_FORMAT_SWITCHES,
  addSwitch,
  dateSwitch,
  fieldNameOf,
  fieldResultKind,
  formatFieldNumber,
  formatFieldText,
  formatSwitch,
  numericSwitch,
} from '../editor/field-switches'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

interface SwitchChoice {
  value: string
  label: string
}

/**
 * Word's Field Options: the switches that shape this field's result, picked
 * from lists and added to the code. Only fields whose result is worked out
 * here get one (the caller keeps Options… greyed for the rest).
 */
function FieldOptionsDialog({
  code,
  onSubmit,
  onClose,
}: {
  code: string
  onSubmit: (code: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [history, setHistory] = useState<string[]>([code])
  const [picked, setPicked] = useState('')
  const current = history[history.length - 1]
  const kind = fieldResultKind(code)
  const lists: Array<{ label: string; name: string; choices: SwitchChoice[] }> = []
  if (kind === 'number') {
    lists.push({
      label: t('appFieldOptFormatting'),
      name: 'format',
      choices: NUMBER_FORMAT_SWITCHES.map((name) => ({
        value: formatSwitch(name),
        label: `${[1, 2, 3].map((n) => formatFieldNumber(`X ${formatSwitch(name)}`, n)).join(', ')}, …`,
      })),
    })
    lists.push({
      label: t('appFieldOptNumeric'),
      name: 'numeric',
      choices: NUMERIC_PICTURES.map((p) => ({ value: numericSwitch(p), label: p })),
    })
  } else if (kind === 'date') {
    lists.push({
      label: t('appFieldOptDateTime'),
      name: 'date',
      choices: DATE_TIME_PICTURES.map((p) => ({ value: dateSwitch(p), label: p })),
    })
  } else if (kind === 'text') {
    lists.push({
      label: t('appFieldOptFormatting'),
      name: 'format',
      choices: TEXT_FORMAT_SWITCHES.map((name) => ({
        value: formatSwitch(name),
        label: formatFieldText(`X ${formatSwitch(name)}`, 'first capital'),
      })),
    })
    if (fieldNameOf(code) === 'FILENAME')
      lists.push({
        label: t('appFieldOptSwitches'),
        name: 'specific',
        choices: [{ value: '\\p', label: `\\p — ${t('appFieldOptPath')}` }],
      })
  }
  const add = (value: string) => {
    if (!value) return
    const next = addSwitch(current, value)
    if (next !== current) setHistory([...history, next])
  }
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={(e) => {
        // Escape closes this dialog only, not the Field dialog under it
        e.stopPropagation()
        keys.onKeyDown(e)
      }}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal tco-dialog field-dialog field-options-dialog">
        <h2>{t('appFieldOptionsTitle')}</h2>
        {lists.map((list) => (
          <label key={list.name}>
            {list.label}
            <select
              size={Math.min(8, Math.max(2, list.choices.length))}
              name={list.name}
              value={list.choices.some((c) => c.value === picked) ? picked : ''}
              onChange={(e) => setPicked(e.target.value)}
              onDoubleClick={(e) => add(e.currentTarget.value)}
            >
              {list.choices.map((c) => (
                <option key={c.value} value={c.value}>
                  {c.label}
                </option>
              ))}
            </select>
          </label>
        ))}
        <div className="field-options-buttons">
          <button className="btn-ghost" name="add" disabled={!picked} onClick={() => add(picked)}>
            {t('appFieldOptAdd')}
          </button>
          <button
            className="btn-ghost"
            name="undo"
            disabled={history.length < 2}
            onClick={() => setHistory(history.slice(0, -1))}
          >
            {t('appFieldOptUndo')}
          </button>
        </div>
        <label>
          {t('appFieldCode')}
          <input name="optionsCode" value={current} readOnly spellCheck={false} />
        </label>
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button className="btn-ghost" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button
              className="btn-primary"
              onClick={() => {
                onSubmit(current)
                onClose()
              }}
            >
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}

const MERGEFORMAT = /\s*\\\*\s*MERGEFORMAT\b/gi
const preserves = (code: string): boolean => code.search(MERGEFORMAT) >= 0

/** Edit Field…: Word's Field dialog reduced to the instruction string (the result is recomputed on OK) */
export function FieldDialog({
  instr,
  pageIsLive,
  onSubmit,
  onClose,
}: {
  instr: string
  /** a header or footer shows PAGE live, numbered by the section's page number format */
  pageIsLive?: boolean
  onSubmit: (instr: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [code, setCode] = useState(instr)
  // Word's switch that keeps the result's formatting when the field updates; on for a new field
  const [preserve, setPreserve] = useState(instr.trim() ? preserves(instr) : true)
  const [options, setOptions] = useState(false)
  const hasOptions = fieldResultKind(code) !== null && !(pageIsLive && fieldNameOf(code) === 'PAGE')
  const submit = () => {
    const typed = code.trim()
    if (!typed) return
    onSubmit(preserve && !preserves(typed) ? `${typed} \\* MERGEFORMAT` : typed)
    onClose()
  }
  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal tco-dialog field-dialog">
        <h2>{t('appFieldDialogTitle')}</h2>
        <label>
          {t('appFieldCode')}
          <input
            autoFocus
            value={code}
            spellCheck={false}
            onChange={(e) => {
              setCode(e.target.value)
              if (preserves(e.target.value)) setPreserve(true)
            }}
            onKeyDown={(e) => e.key === 'Enter' && submit()}
          />
        </label>
        <label className="tco-check">
          <input
            type="checkbox"
            name="preserve"
            checked={preserve}
            onChange={(e) => {
              setPreserve(e.target.checked)
              if (!e.target.checked) setCode(code.replace(MERGEFORMAT, ''))
            }}
          />
          {t('appFieldPreserveFormat')}
        </label>
        <div className="modal-actions">
          <button
            className="btn-ghost"
            name="options"
            disabled={!hasOptions}
            onClick={() => setOptions(true)}
          >
            {t('ribbonHlcOptions')}
          </button>
          <div className="tco-actions-right">
            <button className="btn-ghost" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button className="btn-primary" disabled={!code.trim()} onClick={submit}>
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </div>
      {options && (
        <FieldOptionsDialog
          code={code.trim()}
          onClose={() => setOptions(false)}
          onSubmit={(next) => {
            setCode(next)
            if (preserves(next)) setPreserve(true)
          }}
        />
      )}
    </div>
  )
}
