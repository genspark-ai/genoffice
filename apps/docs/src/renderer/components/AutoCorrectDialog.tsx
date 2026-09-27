/** Word's Tools ▸ AutoCorrect Options…: one switch per replace-as-you-type rule; applies immediately. */
import { useState } from 'react'
import {
  AUTOCORRECT_RULES,
  autocorrectPrefs,
  resetAutocorrectPrefs,
  setAutocorrectPref,
  type AutocorrectRule,
} from '../autocorrect-pref'
import { defaultPasteMode, setDefaultPasteMode, type PasteMode } from '../editor/paste-options'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

const PASTE_MODES: Array<[PasteMode, StringKey]> = [
  ['source', 'appPasteKeepSource'],
  ['merge', 'appPasteMergeFormat'],
  ['text', 'appPasteTextOnly'],
]

const LABELS: Record<AutocorrectRule, StringKey> = {
  smartQuotes: 'appAcSmartQuotes',
  dashes: 'appAcDashes',
  autoLists: 'appAcAutoLists',
  symbols: 'appAcSymbols',
  ordinals: 'appAcOrdinals',
  capitalize: 'appAcCapitalize',
}

export function AutoCorrectDialog({ onClose }: { onClose: () => void }) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [prefs, setPrefs] = useState(autocorrectPrefs)
  const [pasteMode, setPasteMode] = useState(defaultPasteMode)

  const toggle = (rule: AutocorrectRule, on: boolean) => {
    setAutocorrectPref(rule, on)
    setPrefs(autocorrectPrefs())
  }
  const pickPasteMode = (mode: PasteMode) => {
    setDefaultPasteMode(mode)
    setPasteMode(mode)
  }
  const restore = () => {
    resetAutocorrectPrefs()
    setPrefs(autocorrectPrefs())
    pickPasteMode('source')
  }

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal protect-dialog" role="dialog" aria-label={t('appAcTitle')}>
        <h2>{t('appAcTitle')}</h2>
        <p className="modal-desc">{t('appAcIntro')}</p>
        {AUTOCORRECT_RULES.map((rule) => (
          <label key={rule} className="protect-check">
            <input
              type="checkbox"
              checked={prefs[rule]}
              onChange={(e) => toggle(rule, e.target.checked)}
            />
            <span className="ctl" aria-hidden="true" />
            {t(LABELS[rule])}
          </label>
        ))}
        <p className="fld-hint">{t('appAcUndoHint')}</p>
        <h3 className="protect-section-title">{t('appAcPasteHeading')}</h3>
        <p className="modal-desc">{t('appAcPasteFromOther')}</p>
        {PASTE_MODES.map(([mode, key]) => (
          <label key={mode} className="protect-check">
            <input
              type="radio"
              name="default-paste-mode"
              checked={pasteMode === mode}
              onChange={() => pickPasteMode(mode)}
            />
            <span className="ctl" aria-hidden="true" />
            {t(key)}
          </label>
        ))}
        <div className="modal-actions">
          <button onClick={restore}>{t('appAcRestore')}</button>
          <button className="btn-primary" onClick={onClose}>
            {t('appClose')}
          </button>
        </div>
      </div>
    </div>
  )
}
