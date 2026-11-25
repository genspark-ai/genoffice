import { useState } from 'react'
import type { CompareGranularity } from '@genoffice/docx-engine'
import {
  ALL_COMPARE_SETTINGS,
  COMPARE_SETTING_KEYS,
  effectiveSettings,
  originalLabelAvailable,
  settingAvailable,
  targetAvailable,
  type CompareRequest,
  type CompareSettingKey,
  type CompareSide,
  type CompareTarget,
} from '../compare-request'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

type PickedFile = Extract<CompareSide, { kind: 'file' }>

interface Props {
  mode: CompareRequest['mode']
  /** name of the document open in this tab */
  currentName: string
  userName: string
  /** false on a read-only document: changes can only be shown in a new one */
  canWrite: boolean
  onPick: () => Promise<PickedFile | { error: string } | null>
  onRun: (request: CompareRequest) => void
  onClose: () => void
}

const SETTING_LABELS: Record<CompareSettingKey, StringKey> = {
  moves: 'ribbonTcMoves',
  comments: 'ribbonTcComments',
  formatting: 'ribbonTcFormatting',
  caseChanges: 'ribbonCmpDlgCase',
  whiteSpace: 'ribbonCmpDlgWhiteSpace',
  tables: 'ribbonGroupTables',
  headersFooters: 'ribbonCmpDlgHeadersFooters',
  notes: 'ribbonCmpDlgNotes',
  textBoxes: 'ribbonCmpDlgTextBoxes',
  fields: 'ribbonCmpDlgFields',
}

const TARGETS: Array<{ value: CompareTarget; label: StringKey }> = [
  { value: 'original', label: 'ribbonCmpDlgOriginal' },
  { value: 'revised', label: 'ribbonCmpDlgRevised' },
  { value: 'new', label: 'ribbonCmpDlgNewDoc' },
]

/** Word's Tools > Track Changes > Compare Documents and Tools > Combine Documents */
export function CompareDocumentsDialog({
  mode,
  currentName,
  userName,
  canWrite,
  onPick,
  onRun,
  onClose,
}: Props) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [original, setOriginal] = useState<CompareSide | null>({ kind: 'current' })
  const [revised, setRevised] = useState<CompareSide | null>(null)
  /** null follows the revised document's own author */
  const [label, setLabel] = useState<string | null>(null)
  /** null follows the original document's own author */
  const [originalLabel, setOriginalLabel] = useState<string | null>(null)
  const [expanded, setExpanded] = useState(false)
  const [settings, setSettings] = useState(ALL_COMPARE_SETTINGS)
  const [granularity, setGranularity] = useState<CompareGranularity>('word')
  const [wanted, setWanted] = useState<CompareTarget>(mode === 'combine' ? 'original' : 'new')
  const [error, setError] = useState<string | null>(null)

  const target = targetAvailable(wanted, original, revised, canWrite) ? wanted : 'new'
  const author = label ?? (revised?.kind === 'file' ? revised.author : userName)
  const originalAuthor =
    originalLabel ?? (original?.kind === 'file' ? original.author : original ? userName : '')
  const ready =
    !!original && !!revised && !(original.kind === 'current' && revised.kind === 'current')
  const sides = { original, revised }
  const setSide = { original: setOriginal, revised: setRevised }

  const choose = (side: 'original' | 'revised', next: CompareSide) => {
    const other = side === 'original' ? 'revised' : 'original'
    setSide[side](next)
    if (next.kind === 'current' && sides[other]?.kind === 'current') setSide[other](null)
  }

  const browse = async (side: 'original' | 'revised') => {
    const picked = await onPick()
    if (!picked) return
    if ('error' in picked) {
      setError(picked.error)
      return
    }
    setError(null)
    choose(side, picked)
  }

  const originalLabelBox =
    mode === 'compare' ? (
      <>
        <label htmlFor="cmp-label-original">{t('ribbonCmpDlgLabel')}</label>
        {/* always-disabled: Word keeps this box grey in Compare Documents, a comparison takes the revised side's name only */}
        <input
          disabled
          id="cmp-label-original"
          name="compare-label-original"
          type="text"
          title={t('ribbonCmpDlgLabelReason')}
        />
      </>
    ) : (
      <>
        <label htmlFor="cmp-label-original">{t('ribbonCmpDlgLabel')}</label>
        <input
          id="cmp-label-original"
          name="compare-label-original"
          type="text"
          value={originalLabelAvailable(mode, original, target) ? originalAuthor : ''}
          disabled={!originalLabelAvailable(mode, original, target)}
          title={t('ribbonCmpDlgOriginalLabelHint')}
          onChange={(e) => setOriginalLabel(e.target.value)}
        />
      </>
    )

  const sideFields = (side: 'original' | 'revised') => {
    const value = sides[side]
    return (
      <div className="cmp-side">
        <label htmlFor={`cmp-${side}`}>
          {t(side === 'original' ? 'ribbonCmpDlgOriginal' : 'ribbonCmpDlgRevised')}
        </label>
        <div className="cmp-pick">
          <select
            id={`cmp-${side}`}
            name={`compare-${side}`}
            value={value?.kind ?? ''}
            onChange={(e) => e.target.value === 'current' && choose(side, { kind: 'current' })}
          >
            {!value && <option value="">{t('ribbonCmpDlgChoose')}</option>}
            <option value="current">{t('ribbonCmpDlgCurrent', { name: currentName })}</option>
            {value?.kind === 'file' && <option value="file">{value.name}</option>}
          </select>
          <button type="button" name={`browse-${side}`} onClick={() => void browse(side)}>
            {t('ribbonCmpDlgBrowse')}
          </button>
        </div>
        {side === 'revised' ? (
          <>
            <label htmlFor="cmp-label-revised">{t('ribbonCmpDlgLabel')}</label>
            <input
              id="cmp-label-revised"
              name="compare-label-revised"
              type="text"
              value={revised ? author : ''}
              disabled={!revised}
              onChange={(e) => setLabel(e.target.value)}
            />
          </>
        ) : (
          originalLabelBox
        )}
      </div>
    )
  }

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog cmp-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          if (!ready) return
          onRun({
            mode,
            original,
            revised,
            author: author.trim() || userName,
            ...(originalLabelAvailable(mode, original, target)
              ? { originalAuthor: originalAuthor.trim() || userName }
              : {}),
            settings: effectiveSettings(settings, target),
            granularity,
            target,
          })
          onClose()
        }}
      >
        <h2>
          {t(mode === 'combine' ? 'ribbonCombineDocuments' : 'ribbonCompareDocuments').replace(
            /\s*(?:…|\.\.\.)$/,
            '',
          )}
        </h2>
        <div className="cmp-sides">
          {sideFields('original')}
          <button
            type="button"
            className="cmp-swap"
            name="compare-swap"
            title={t('ribbonCmpDlgSwap')}
            aria-label={t('ribbonCmpDlgSwap')}
            disabled={!original || !revised}
            onClick={() => {
              setOriginal(revised)
              setRevised(original)
              setLabel(originalLabel)
              setOriginalLabel(label)
            }}
          >
            {'⇄'}
          </button>
          {sideFields('revised')}
        </div>
        {error && (
          <p className="cmp-error" role="alert">
            {error}
          </p>
        )}
        <button
          type="button"
          className="cmp-more"
          name="compare-options"
          aria-expanded={expanded}
          onClick={() => setExpanded((on) => !on)}
        >
          {t('ribbonGroupHfOptions')} {expanded ? '▴' : '▾'}
        </button>
        {expanded && (
          <div className="cmp-options">
            <fieldset>
              <legend>{t('ribbonCmpDlgSettings')}</legend>
              <div className="cmp-settings">
                {COMPARE_SETTING_KEYS.map((key) => {
                  const available = settingAvailable(key, target)
                  return (
                    <label
                      key={key}
                      className="tco-check"
                      title={available ? undefined : t('ribbonCmpDlgNewOnlyReason')}
                    >
                      <input
                        type="checkbox"
                        name={`compare-${key}`}
                        checked={available && settings[key]}
                        disabled={!available}
                        onChange={(e) => setSettings({ ...settings, [key]: e.target.checked })}
                      />
                      {t(SETTING_LABELS[key])}
                    </label>
                  )
                })}
              </div>
              {target !== 'new' && <p className="cmp-reason">{t('ribbonCmpDlgNewOnlyReason')}</p>}
            </fieldset>
            <fieldset>
              <legend>{t('ribbonCmpDlgShowAt')}</legend>
              {(['char', 'word'] as const).map((value) => (
                <label key={value} className="tco-check">
                  <input
                    type="radio"
                    name="compare-level"
                    value={value}
                    checked={granularity === value}
                    onChange={() => setGranularity(value)}
                  />
                  {t(value === 'char' ? 'ribbonCmpDlgCharLevel' : 'ribbonCmpDlgWordLevel')}
                </label>
              ))}
            </fieldset>
            <fieldset>
              <legend>{t('ribbonCmpDlgShowIn')}</legend>
              {TARGETS.map(({ value, label: text }) => (
                <label key={value} className="tco-check">
                  <input
                    type="radio"
                    name="compare-target"
                    value={value}
                    checked={target === value}
                    disabled={!targetAvailable(value, original, revised, canWrite)}
                    onChange={() => setWanted(value)}
                  />
                  {t(text)}
                </label>
              ))}
              {TARGETS.some(
                ({ value }) => !targetAvailable(value, original, revised, canWrite),
              ) && <p className="cmp-reason">{t('ribbonCmpDlgTargetReason')}</p>}
            </fieldset>
          </div>
        )}
        <div className="modal-actions">
          <button type="button" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button type="submit" className="primary" disabled={!ready}>
            {t('ribbonOk')}
          </button>
        </div>
      </form>
    </div>
  )
}
