import { useState } from 'react'

import { Dropdown } from '@genoffice/ui'

import {
  isValidConsolidateReference,
  type ConsolidateConfig,
  type ConsolidateFn,
} from './consolidate'
import { useI18n } from './i18n/locale'

/// Excel's Data → Consolidate: a function, a list of source references, and
/// an optional "use labels in left column" mode. Output lands at the active
/// cell as live formulas over the sources.

export function ConsolidateDialog({
  defaultReference,
  targetLabel,
  onCreate,
  onClose,
}: {
  /// Prefill for the reference input, usually the current selection.
  readonly defaultReference: string
  readonly targetLabel: string
  /// Returns an error message, or null on success.
  readonly onCreate: (config: ConsolidateConfig) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [fn, setFn] = useState<ConsolidateFn>('sum')
  const [references, setReferences] = useState<string[]>([])
  const [input, setInput] = useState(defaultReference)
  const [leftLabels, setLeftLabels] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const addReference = (): string | null => {
    const trimmed = input.trim()
    if (trimmed === '') return null
    if (!isValidConsolidateReference(trimmed)) {
      return t('dlgConsBadRef', { ref: trimmed })
    }
    if (references.includes(trimmed)) {
      setInput('')
      return null
    }
    setReferences([...references, trimmed])
    setInput('')
    return null
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={t('dlgConsTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgConsTitle')}</header>
        <div className="dialog-grid">
          <label>
            {t('dlgConsFunction')}
            {/* max/min were disabled options in label mode; the shared Dropdown
                has no per-option disabling, so they drop out of the list */}
            <Dropdown
              ariaLabel={t('dlgConsFunction')}
              value={fn}
              options={[
                { value: 'sum', label: t('dlgPivotAggSum') },
                { value: 'count', label: t('dlgPivotAggCount') },
                { value: 'average', label: t('dlgPivotAggAverage') },
                ...(leftLabels
                  ? []
                  : [
                      { value: 'max', label: t('dlgPivotAggMax') },
                      { value: 'min', label: t('dlgPivotAggMin') },
                    ]),
              ]}
              onPick={(v) => setFn(v as ConsolidateFn)}
            />
          </label>
          <label className="dialog-span">
            {t('dlgConsReference')}
            <span className="input-row">
              <input
                value={input}
                placeholder="Sheet2!A1:C5"
                onChange={(event) => setInput(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault()
                    setError(addReference())
                  }
                }}
              />
              <button className="secondary" onClick={() => setError(addReference())}>
                {t('dlgConsAdd')}
              </button>
            </span>
          </label>
          {references.length > 0 && (
            <ul className="reference-list" aria-label={t('dlgConsAllRefs')}>
              {references.map((reference) => (
                <li key={reference}>
                  <code>{reference}</code>
                  <button
                    className="secondary"
                    aria-label={t('dlgConsDeleteRef', { ref: reference })}
                    onClick={() => setReferences(references.filter((entry) => entry !== reference))}
                  >
                    {t('dlgConsDelete')}
                  </button>
                </li>
              ))}
            </ul>
          )}
          <label className="dialog-check dialog-span">
            <input
              type="checkbox"
              checked={leftLabels}
              onChange={(event) => {
                setLeftLabels(event.target.checked)
                if (event.target.checked && (fn === 'max' || fn === 'min')) setFn('sum')
              }}
            />
            {t('dlgConsLeftLabels')}
          </label>
          <p className="dialog-note dialog-span">
            {leftLabels
              ? t('dlgConsNoteLabels', { target: targetLabel })
              : t('dlgConsNotePosition', { target: targetLabel })}
          </p>
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
              // A typed-but-not-added reference still counts on OK.
              const pending = input.trim()
              let all = references
              if (pending !== '') {
                if (!isValidConsolidateReference(pending)) {
                  setError(t('dlgConsBadRef', { ref: pending }))
                  return
                }
                if (!references.includes(pending)) all = [...references, pending]
              }
              if (all.length === 0) {
                setError(t('dlgConsNeedOneRef'))
                return
              }
              const failure = onCreate({ fn, references: all, leftLabels })
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
