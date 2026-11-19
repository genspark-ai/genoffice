/** Word's Tools ▸ AutoCorrect Options…: every switch and list applies as it is changed. */
import { useState, type ReactNode } from 'react'
import {
  autocorrectLists,
  autocorrectPrefs,
  resetAutocorrectPrefs,
  setAutocorrectList,
  setAutocorrectPref,
  upsertEntry,
  type AutocorrectRule,
  type ReplaceEntry,
} from '../autocorrect-pref'
import type { AutoTextEntry } from '../editor/auto-text'
import { defaultPasteMode, setDefaultPasteMode, type PasteMode } from '../editor/paste-options'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'
import { Unavailable } from './Unavailable'

const PASTE_MODES: Array<[PasteMode, StringKey]> = [
  ['source', 'appPasteKeepSource'],
  ['merge', 'appPasteMergeFormat'],
  ['text', 'appPasteTextOnly'],
]

type Tab = 'autocorrect' | 'math' | 'autoformat' | 'autotext'
const TABS: Array<[Tab, StringKey]> = [
  ['autocorrect', 'appAcTabAutoCorrect'],
  ['math', 'appAcTabMath'],
  ['autoformat', 'appAcTabAutoFormat'],
  ['autotext', 'appAcTabAutoText'],
]

export interface AutoCorrectSelection {
  text: string
  html: string
}

export interface AutoCorrectAutoText {
  entries: AutoTextEntry[]
  /** the start of the selection, or '' when nothing is selected to keep */
  suggestedName: string
  onAdd: (name: string) => void
  onInsert: (entry: AutoTextEntry) => void
  onDelete: (name: string) => void
}

/** the Replace / With table of the AutoCorrect and Math AutoCorrect tabs */
function EntryTable({
  entries,
  onChange,
  disabled,
  selection,
}: {
  entries: ReplaceEntry[]
  onChange: (next: ReplaceEntry[]) => void
  disabled: boolean
  /** AutoCorrect tab only: what is selected in the document, offered as a formatted replacement */
  selection?: AutoCorrectSelection | null
}) {
  const { t } = useI18n()
  const [from, setFrom] = useState('')
  const [to, setTo] = useState(selection?.text ?? '')
  const [formatted, setFormatted] = useState(!!selection)
  // where the formatting comes from: the selection, or the picked entry's stored copy
  const [source, setSource] = useState(selection ?? null)
  // it belongs to that text as it stands, not to an edited copy
  const html = formatted && source && to === source.text ? source.html : undefined
  const known = entries.some(([f]) => f === from)
  return (
    <div className="ac-entries">
      <div className="ac-entry-fields">
        <label>
          {t('appAcReplaceLabel')}
          <input
            type="text"
            value={from}
            spellCheck={false}
            disabled={disabled}
            onChange={(e) => setFrom(e.target.value)}
          />
        </label>
        <label>
          {t('appAcWithLabel')}
          <input
            type="text"
            value={to}
            spellCheck={false}
            disabled={disabled}
            onChange={(e) => setTo(e.target.value)}
          />
        </label>
      </div>
      {selection !== undefined && (
        <div className="ac-entry-kind">
          {(['plain', 'formatted'] as const).map((kind) => (
            <label key={kind} className="protect-check">
              <input
                type="radio"
                name="autocorrect-entry-kind"
                checked={formatted === (kind === 'formatted')}
                disabled={disabled || !source}
                onChange={() => setFormatted(kind === 'formatted')}
              />
              <span className="ctl" aria-hidden="true" />
              {t(kind === 'plain' ? 'appAcPlainText' : 'appAcFormattedText')}
            </label>
          ))}
        </div>
      )}
      <div className="ac-table" role="listbox" aria-label={t('appAcReplaceLabel')}>
        {entries.map(([f, w, entryHtml]) => (
          <button
            key={f}
            type="button"
            role="option"
            aria-selected={f === from}
            className={f === from ? 'ac-row active' : 'ac-row'}
            disabled={disabled}
            onClick={() => {
              setFrom(f)
              setTo(w)
              setSource(entryHtml ? { text: w, html: entryHtml } : (selection ?? null))
              setFormatted(!!entryHtml)
            }}
          >
            <span>{f}</span>
            <span>{w}</span>
          </button>
        ))}
      </div>
      <div className="ac-entry-actions">
        <button
          type="button"
          disabled={disabled || from.trim() === '' || to === ''}
          onClick={() => onChange(upsertEntry(entries, from.trim(), to, html))}
        >
          {t(known ? 'appReplace' : 'ribbonAdd')}
        </button>
        <button
          type="button"
          disabled={disabled || !known}
          onClick={() => {
            onChange(entries.filter(([f]) => f !== from))
            setFrom('')
            setTo('')
          }}
        >
          {t('ribbonGroupDelete')}
        </button>
      </div>
    </div>
  )
}

function WordList({
  label,
  words,
  onChange,
}: {
  label: string
  words: string[]
  onChange: (next: string[]) => void
}) {
  const { t } = useI18n()
  const [word, setWord] = useState('')
  const typed = word.trim()
  const known = words.includes(typed)
  return (
    <div className="ac-entries">
      <div className="ac-entry-fields">
        <label>
          {label}
          <input
            type="text"
            value={word}
            spellCheck={false}
            autoFocus
            onChange={(e) => setWord(e.target.value)}
          />
        </label>
      </div>
      <div className="ac-table" role="listbox" aria-label={label}>
        {words.map((w) => (
          <button
            key={w}
            type="button"
            role="option"
            aria-selected={w === typed}
            className={w === typed ? 'ac-row active' : 'ac-row'}
            onClick={() => setWord(w)}
          >
            <span>{w}</span>
          </button>
        ))}
      </div>
      <div className="ac-entry-actions">
        <button
          type="button"
          disabled={typed === '' || known}
          onClick={() => {
            onChange([...words, typed].sort((a, b) => a.localeCompare(b)))
            setWord('')
          }}
        >
          {t('ribbonAdd')}
        </button>
        <button
          type="button"
          disabled={!known}
          onClick={() => {
            onChange(words.filter((w) => w !== typed))
            setWord('')
          }}
        >
          {t('ribbonGroupDelete')}
        </button>
      </div>
    </div>
  )
}

function ExceptionsDialog({
  firstLetter,
  initialCaps,
  onChange,
  onClose,
}: {
  firstLetter: string[]
  initialCaps: string[]
  onChange: (name: 'firstLetter' | 'initialCaps', next: string[]) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [tab, setTab] = useState<'firstLetter' | 'initialCaps'>('firstLetter')
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={(e) => {
        e.stopPropagation()
        keys.onKeyDown(e)
      }}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        className="modal protect-dialog autocorrect-dialog ac-exceptions"
        role="dialog"
        aria-label={t('appAcExceptionsTitle')}
      >
        <h2>{t('appAcExceptionsTitle')}</h2>
        <div className="dlg-tabs" role="tablist">
          {(['firstLetter', 'initialCaps'] as const).map((k) => (
            <button
              key={k}
              type="button"
              role="tab"
              aria-selected={tab === k}
              className={tab === k ? 'active' : ''}
              onClick={() => setTab(k)}
            >
              {t(k === 'firstLetter' ? 'appAcExFirstLetter' : 'appAcExInitialCaps')}
            </button>
          ))}
        </div>
        {tab === 'firstLetter' ? (
          <WordList
            key="firstLetter"
            label={t('appAcExDontCapitalizeAfter')}
            words={firstLetter}
            onChange={(next) => onChange('firstLetter', next)}
          />
        ) : (
          <WordList
            key="initialCaps"
            label={t('appAcExDontCorrect')}
            words={initialCaps}
            onChange={(next) => onChange('initialCaps', next)}
          />
        )}
        <div className="modal-actions">
          <button className="btn-primary" onClick={onClose}>
            {t('appClose')}
          </button>
        </div>
      </div>
    </div>
  )
}

function FunctionsDialog({
  functions,
  onChange,
  onClose,
}: {
  functions: string[]
  onChange: (next: string[]) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={(e) => {
        e.stopPropagation()
        keys.onKeyDown(e)
      }}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        className="modal protect-dialog autocorrect-dialog ac-exceptions"
        role="dialog"
        aria-label={t('appAcFunctionsTitle')}
      >
        <h2>{t('appAcFunctionsTitle')}</h2>
        <WordList label={t('appAcFunctionsHint')} words={functions} onChange={onChange} />
        <div className="modal-actions">
          <button className="btn-primary" onClick={onClose}>
            {t('appClose')}
          </button>
        </div>
      </div>
    </div>
  )
}

function AutoTextTab({
  autoText,
  switchRow,
}: {
  autoText: AutoCorrectAutoText
  switchRow: ReactNode
}) {
  const { t } = useI18n()
  const [name, setName] = useState(autoText.suggestedName)
  const typed = name.trim()
  const entry = autoText.entries.find((e) => e.name === typed)
  return (
    <>
      {switchRow}
      <div className="ac-entries">
        <div className="ac-entry-fields">
          <label>
            {t('appAcAutoTextEntries')}
            <input
              type="text"
              value={name}
              spellCheck={false}
              onChange={(e) => setName(e.target.value)}
            />
          </label>
        </div>
        <div className="ac-table" role="listbox" aria-label={t('appAcAutoTextEntries')}>
          {autoText.entries.map((e) => (
            <button
              key={e.name}
              type="button"
              role="option"
              aria-selected={e.name === typed}
              className={e.name === typed ? 'ac-row active' : 'ac-row'}
              onClick={() => setName(e.name)}
              onDoubleClick={() => autoText.onInsert(e)}
            >
              <span>{e.name}</span>
            </button>
          ))}
        </div>
        <div className="ac-entry-actions">
          <button
            type="button"
            disabled={typed === '' || autoText.suggestedName === ''}
            onClick={() => autoText.onAdd(typed)}
          >
            {t('ribbonAdd')}
          </button>
          <button type="button" disabled={!entry} onClick={() => autoText.onDelete(typed)}>
            {t('ribbonGroupDelete')}
          </button>
          <button type="button" disabled={!entry} onClick={() => entry && autoText.onInsert(entry)}>
            {t('appInsert')}
          </button>
        </div>
        <Unavailable id="autocorrect.autotext-look-in">
          <label className="ac-look-in">
            {t('appAcLookIn')}
            <select disabled>
              <option>{t('appAcAllTemplates')}</option>
            </select>
          </label>
        </Unavailable>
        <h3 className="protect-section-title">{t('ribbonPbPreview')}</h3>
        <p className="fld-hint ac-preview">{entry?.text.slice(0, 240)}</p>
      </div>
    </>
  )
}

export function AutoCorrectDialog({
  onClose,
  autoText,
  selection,
}: {
  onClose: () => void
  autoText: AutoCorrectAutoText
  /** one paragraph or less selected in the document when the dialog opened */
  selection: AutoCorrectSelection | null
}) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [tab, setTab] = useState<Tab>('autocorrect')
  const [prefs, setPrefs] = useState(autocorrectPrefs)
  const [lists, setLists] = useState(autocorrectLists)
  const [pasteMode, setPasteMode] = useState(defaultPasteMode)
  const [exceptions, setExceptions] = useState(false)
  const [functions, setFunctions] = useState(false)

  const toggle = (rule: AutocorrectRule, on: boolean) => {
    setAutocorrectPref(rule, on)
    setPrefs(autocorrectPrefs())
  }
  const setList: typeof setAutocorrectList = (name, value) => {
    setAutocorrectList(name, value)
    setLists(autocorrectLists())
  }
  const pickPasteMode = (mode: PasteMode) => {
    setDefaultPasteMode(mode)
    setPasteMode(mode)
  }
  const restore = () => {
    resetAutocorrectPrefs()
    setPrefs(autocorrectPrefs())
    setLists(autocorrectLists())
    pickPasteMode('source')
  }

  const check = (rule: AutocorrectRule, label: StringKey) => (
    <label key={rule} className="protect-check">
      <input
        type="checkbox"
        checked={prefs[rule]}
        disabled={rule !== 'enabled' && !prefs.enabled}
        onChange={(e) => toggle(rule, e.target.checked)}
      />
      <span className="ctl" aria-hidden="true" />
      {t(label)}
    </label>
  )
  /** a rule Word lists and this app has nothing to switch for */
  const off = (id: string, label: StringKey) => (
    <Unavailable id={id}>
      <label className="protect-check">
        <input type="checkbox" checked={false} disabled readOnly />
        <span className="ctl" aria-hidden="true" />
        {t(label)}
      </label>
    </Unavailable>
  )
  const heading = (label: StringKey) => <h3 className="protect-section-title">{t(label)}</h3>

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        className="modal protect-dialog autocorrect-dialog"
        role="dialog"
        aria-label={t('appAcTitle')}
      >
        <h2>{t('appAcTitle')}</h2>
        {check('enabled', 'appAcMaster')}
        <div className="dlg-tabs" role="tablist">
          {TABS.map(([k, label]) => (
            <button
              key={k}
              type="button"
              role="tab"
              aria-selected={tab === k}
              className={tab === k ? 'active' : ''}
              onClick={() => setTab(k)}
            >
              {t(label)}
            </button>
          ))}
        </div>
        {tab === 'autocorrect' && (
          <>
            {check('smartButton', 'appAcSmartButton')}
            <div className="ac-with-button">
              <div>
                {check('twoInitialCaps', 'appAcTwoCaps')}
                {check('capitalize', 'appAcCapitalize')}
                {check('capitalizeDays', 'appAcDays')}
                {check('capitalizeCells', 'appAcCells')}
              </div>
              <button type="button" onClick={() => setExceptions(true)}>
                {t('appAcExceptions')}
              </button>
            </div>
            {check('replaceText', 'appAcReplaceText')}
            <EntryTable
              entries={lists.replace}
              selection={selection}
              disabled={!prefs.enabled || !prefs.replaceText}
              onChange={(next) => setList('replace', next)}
            />
            {off('autocorrect.spelling-suggestions', 'appAcSpellSuggest')}
            <p className="fld-hint">{t('appAcUndoHint')}</p>
            {heading('appAcPasteHeading')}
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
          </>
        )}
        {tab === 'math' && (
          <>
            {check('mathOutside', 'appAcMathOutside')}
            <p className="fld-hint">{t('appAcMathConflict')}</p>
            {check('mathInside', 'appAcReplaceText')}
            <EntryTable
              entries={lists.math}
              disabled={!prefs.enabled || (!prefs.mathOutside && !prefs.mathInside)}
              onChange={(next) => setList('math', next)}
            />
            <div className="ac-entry-actions">
              <button type="button" onClick={() => setFunctions(true)}>
                {t('appAcRecognizedFunctions')}
              </button>
            </div>
          </>
        )}
        {tab === 'autoformat' && (
          <>
            {heading('appAcReplaceHeading')}
            {check('autoBullets', 'appAcBullets')}
            {check('autoNumbers', 'appAcNumbers')}
            {check('borderLines', 'appAcBorderLines')}
            {check('headingStyles', 'appAcHeadingStyles')}
            {check('tables', 'ribbonGroupTables')}
            {off('autocorrect.date-style', 'appAcDateStyle')}
            {check('indentSpaces', 'appAcIndentSpaces')}
            {heading('appAcAutoHeading')}
            {check('asianSpaces', 'appAcAsianSpaces')}
            {check('boldItalic', 'appAcBoldItalic')}
            {check('fractions', 'appAcFractions')}
            {check('hyperlinks', 'appAcHyperlinks')}
            {check('parens', 'appAcParens')}
            {check('ordinals', 'appAcOrdinals')}
            {check('smartQuotes', 'appAcSmartQuotes')}
            {check('dashes', 'appAcDashes')}
            {heading('appAcAutoHeading')}
            {off('autocorrect.define-styles', 'appAcDefineStyles')}
            {check('tabIndents', 'appAcTabIndents')}
            {check('listStart', 'appAcListStart')}
            {check('closingPhrase', 'appAcClosingPhrase')}
          </>
        )}
        {tab === 'autotext' && (
          <AutoTextTab autoText={autoText} switchRow={check('autoComplete', 'appAcAutoComplete')} />
        )}
        <div className="modal-actions">
          <button onClick={restore}>{t('appAcRestore')}</button>
          <button className="btn-primary" onClick={onClose}>
            {t('appClose')}
          </button>
        </div>
      </div>
      {functions && (
        <FunctionsDialog
          functions={lists.functions}
          onChange={(next) => setList('functions', next)}
          onClose={() => setFunctions(false)}
        />
      )}
      {exceptions && (
        <ExceptionsDialog
          firstLetter={lists.firstLetter}
          initialCaps={lists.initialCaps}
          onChange={setList}
          onClose={() => setExceptions(false)}
        />
      )}
    </div>
  )
}
