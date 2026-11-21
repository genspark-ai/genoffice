import { useState } from 'react'
import type { NoteProps } from '@genoffice/docx-engine'
import { fontSymbolDisplay } from '../editor/font-symbol'
import { useI18n } from '../i18n/locale'
import type { NoteKind } from '../note-format'
import type { NoteConversion } from '../review-actions'
import { useModalKeys } from './modal-keys'
import { SymbolDialog } from './SymbolDialog'
import { Unavailable } from './Unavailable'

export type NoteSettings = Record<NoteKind, NoteProps>
export type NoteApplyScope = 'document' | 'section'

/** Word's number format list, as it labels them */
const NUMBER_FORMATS: Array<[string, string]> = [
  ['decimal', '1, 2, 3, …'],
  ['lowerLetter', 'a, b, c, …'],
  ['upperLetter', 'A, B, C, …'],
  ['lowerRoman', 'i, ii, iii, …'],
  ['upperRoman', 'I, II, III, …'],
  ['chicago', '*, †, ‡, §, …'],
]

/** w:numStart as typed: 0 is a value a file can carry */
const START_AT = /^\d{1,5}$/

const DEFAULTS: Record<NoteKind, Required<NoteProps>> = {
  footnote: { pos: 'pageBottom', numFmt: 'decimal', numStart: 1, numRestart: 'continuous' },
  endnote: { pos: 'docEnd', numFmt: 'lowerRoman', numStart: 1, numRestart: 'continuous' },
}

const filled = (kind: NoteKind, props: NoteProps): Required<NoteProps> => ({
  ...DEFAULTS[kind],
  ...props,
})

const same = (a: NoteProps, b: NoteProps) =>
  a.pos === b.pos &&
  a.numFmt === b.numFmt &&
  a.numStart === b.numStart &&
  a.numRestart === b.numRestart

function ConvertNotesDialog({
  counts,
  onConvert,
  onClose,
}: {
  counts: Record<NoteKind, number>
  onConvert: (mode: NoteConversion) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const available: Record<NoteConversion, boolean> = {
    toEndnotes: counts.footnote > 0,
    toFootnotes: counts.endnote > 0,
    swap: counts.footnote > 0 && counts.endnote > 0,
  }
  const [mode, setMode] = useState<NoteConversion>(
    available.toEndnotes ? 'toEndnotes' : 'toFootnotes',
  )
  const option = (value: NoteConversion, label: string) => (
    <label className="tco-check">
      <input
        type="radio"
        name="convert-notes"
        value={value}
        checked={mode === value}
        disabled={!available[value]}
        autoFocus={mode === value}
        onChange={() => setMode(value)}
      />
      {label}
    </label>
  )
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog convert-notes-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onConvert(mode)
        }}
      >
        <h2>{t('ribbonFnDlgConvertTitle')}</h2>
        {option('toEndnotes', t('ribbonFnDlgConvertToEndnotes'))}
        {option('toFootnotes', t('ribbonFnDlgConvertToFootnotes'))}
        {option('swap', t('ribbonFnDlgConvertSwap'))}
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary">
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}

/** Word's Insert > Footnote…: where notes sit, how they are numbered, and inserting one */
export function FootnoteDialog({
  settings,
  multiSection,
  counts,
  onApply,
  onInsert,
  onConvert,
  onClose,
}: {
  /** the numbering in effect in the caret's section, and where the document places notes */
  settings: NoteSettings
  /** the document has more than one section, so one of them can be numbered on its own */
  multiSection: boolean
  /** notes in the document that Convert can move */
  counts: Record<NoteKind, number>
  onApply: (next: NoteSettings, scope: NoteApplyScope) => void
  onInsert: (kind: NoteKind, customMark: string) => void
  onConvert: (mode: NoteConversion) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const initial: Record<NoteKind, Required<NoteProps>> = {
    footnote: filled('footnote', settings.footnote),
    endnote: filled('endnote', settings.endnote),
  }
  const [kind, setKind] = useState<NoteKind>('footnote')
  const [draft, setDraft] = useState(initial)
  const [startText, setStartText] = useState<Record<NoteKind, string>>({
    footnote: String(initial.footnote.numStart),
    endnote: String(initial.endnote.numStart),
  })
  const [scope, setScope] = useState<NoteApplyScope>('document')
  const [mark, setMark] = useState('')
  const [picking, setPicking] = useState(false)
  const [converting, setConverting] = useState(false)

  const cur = draft[kind]
  const set = (patch: NoteProps) => setDraft((d) => ({ ...d, [kind]: { ...d[kind], ...patch } }))
  // only the field on screen can hold the buttons back; the other kind's was reset when it left
  const startValid = START_AT.test(startText[kind].trim())
  const show = (next: NoteKind) => {
    setStartText((t) => ({ ...t, [kind]: String(draft[kind].numStart) }))
    setKind(next)
  }
  const changed = !same(draft.footnote, initial.footnote) || !same(draft.endnote, initial.endnote)
  const apply = () => {
    if (!changed) return
    // only what the dialog changed goes out: untouched fields keep whatever the file says
    const delta = (k: NoteKind): NoteProps =>
      Object.fromEntries(
        (['pos', 'numFmt', 'numStart', 'numRestart'] as const)
          .filter((f) => scope === 'document' || f !== 'pos')
          .filter((f) => draft[k][f] !== initial[k][f] || settings[k][f] !== undefined)
          .map((f) => [f, draft[k][f]]),
      )
    onApply({ footnote: delta('footnote'), endnote: delta('endnote') }, scope)
  }
  const applyTo = (next: NoteApplyScope) => {
    setScope(next)
    // the page places notes by one document-wide setting: a section cannot take its own
    if (next === 'section')
      setDraft((d) => ({
        footnote: { ...d.footnote, pos: initial.footnote.pos },
        endnote: { ...d.endnote, pos: initial.endnote.pos },
      }))
  }
  const formats = NUMBER_FORMATS.some(([value]) => value === cur.numFmt)
    ? NUMBER_FORMATS
    : [...NUMBER_FORMATS, [cur.numFmt, cur.numFmt] as [string, string]]

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog footnote-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          if (!startValid) return
          apply()
          onInsert(kind, mark.trim())
        }}
      >
        <h2>{t('ribbonFnDlgTitle')}</h2>
        <fieldset>
          <legend>{t('ribbonBookmarkSortLocation')}</legend>
          <div className="tco-grid fn-dlg-grid">
            <label className="tco-check">
              <input
                type="radio"
                name="note-kind"
                value="footnote"
                checked={kind === 'footnote'}
                onChange={() => show('footnote')}
              />
              {t('ribbonGroupFootnotes')}
            </label>
            <select
              name="footnote-pos"
              aria-label={t('ribbonGroupFootnotes')}
              disabled={kind !== 'footnote' || scope === 'section'}
              value={draft.footnote.pos}
              onChange={(e) => set({ pos: e.target.value as NoteProps['pos'] })}
            >
              <option value="pageBottom">{t('ribbonFnDlgBottomOfPage')}</option>
              <option value="beneathText">{t('ribbonFnDlgBelowText')}</option>
            </select>
            <label className="tco-check">
              <input
                type="radio"
                name="note-kind"
                value="endnote"
                checked={kind === 'endnote'}
                onChange={() => show('endnote')}
              />
              {t('ribbonFnDlgEndnotes')}
            </label>
            <select
              name="endnote-pos"
              aria-label={t('ribbonFnDlgEndnotes')}
              disabled={kind !== 'endnote' || scope === 'section'}
              value={draft.endnote.pos}
              onChange={(e) => set({ pos: e.target.value as NoteProps['pos'] })}
            >
              <option value="docEnd">{t('ribbonFnDlgEndOfDocument')}</option>
              <option value="sectEnd">{t('ribbonFnDlgEndOfSection')}</option>
            </select>
          </div>
          <div className="fn-dlg-convert">
            <button
              type="button"
              name="convert"
              disabled={counts.footnote + counts.endnote === 0}
              onClick={() => setConverting(true)}
            >
              {t('ribbonFnDlgConvert')}
            </button>
          </div>
        </fieldset>
        <Unavailable id="footnote.columns">
          <fieldset>
            <legend>{t('ribbonFnDlgLayout')}</legend>
            <div className="tco-grid fn-dlg-grid">
              <span>{t('ribbonColumns')}</span>
              <select name="note-columns" aria-label={t('ribbonColumns')} disabled value="match">
                <option value="match">{t('ribbonFnDlgMatchSection')}</option>
              </select>
            </div>
          </fieldset>
        </Unavailable>
        <fieldset>
          <legend>{t('ribbonFnDlgFormat')}</legend>
          <div className="tco-grid fn-dlg-grid">
            <span>{t('ribbonListFormatText')}</span>
            <select
              name="note-format"
              aria-label={t('ribbonListFormatText')}
              value={cur.numFmt}
              onChange={(e) => set({ numFmt: e.target.value })}
            >
              {formats.map(([value, label]) => (
                <option key={value} value={value}>
                  {label}
                </option>
              ))}
            </select>
            <span>{t('ribbonFnDlgCustomMark')}</span>
            <div className="fn-dlg-mark">
              <input
                type="text"
                name="custom-mark"
                aria-label={t('ribbonFnDlgCustomMark')}
                maxLength={10}
                value={mark}
                onChange={(e) => setMark(e.target.value)}
              />
              <button type="button" name="symbol" onClick={() => setPicking(true)}>
                {t('ribbonBulletSymbol')}…
              </button>
            </div>
            <span>{t('ribbonListStartAt')}</span>
            <input
              type="number"
              name="note-start"
              aria-label={t('ribbonListStartAt')}
              min={0}
              max={99999}
              value={startText[kind]}
              onChange={(e) => {
                const text = e.target.value
                setStartText((s) => ({ ...s, [kind]: text }))
                if (START_AT.test(text.trim())) set({ numStart: parseInt(text, 10) })
              }}
            />
            <span>{t('ribbonLnNumbering')}</span>
            <select
              name="note-restart"
              aria-label={t('ribbonLnNumbering')}
              value={cur.numRestart}
              onChange={(e) => set({ numRestart: e.target.value as NoteProps['numRestart'] })}
            >
              <option value="continuous">{t('ribbonBreakContinuous')}</option>
              <option value="eachSect">{t('ribbonLnRestartSection')}</option>
              {cur.numRestart === 'eachPage' ? (
                // always-disabled: the value the file carries, listed so the box can show it
                <option value="eachPage" disabled>
                  {t('ribbonLnRestartPage')}
                </option>
              ) : (
                kind === 'footnote' && (
                  <Unavailable id="footnote.restart-each-page">
                    <option value="eachPage" disabled>
                      {t('ribbonLnRestartPage')}
                    </option>
                  </Unavailable>
                )
              )}
            </select>
          </div>
        </fieldset>
        <fieldset>
          <legend>{t('ribbonFnDlgApplyChanges')}</legend>
          <div className="tco-grid fn-dlg-grid">
            <span>{t('ribbonFnDlgApplyTo')}</span>
            <select
              name="note-apply-to"
              aria-label={t('ribbonFnDlgApplyTo')}
              value={scope}
              onChange={(e) => applyTo(e.target.value as NoteApplyScope)}
            >
              <option value="document">{t('ribbonPbApplyWhole')}</option>
              <option value="section" disabled={!multiSection}>
                {t('ribbonPbApplySection')}
              </option>
            </select>
          </div>
        </fieldset>
        <div className="modal-actions">
          <button type="button" name="apply" disabled={!changed || !startValid} onClick={apply}>
            {t('ribbonApply')}
          </button>
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary" disabled={!startValid}>
              {t('ribbonInsert')}
            </button>
          </div>
        </div>
      </form>
      {picking && (
        <SymbolDialog
          onInsert={(symbol, font) => {
            // a symbol-font character has no plain-text form unless the font maps to Unicode
            const shown = font ? fontSymbolDisplay(font, symbol) : { text: symbol, ownFont: false }
            if (!shown || shown.ownFont) return
            setMark(shown.text)
            setPicking(false)
          }}
          onClose={() => setPicking(false)}
        />
      )}
      {converting && (
        <ConvertNotesDialog
          counts={counts}
          onConvert={(mode) => {
            setConverting(false)
            onConvert(mode)
          }}
          onClose={() => setConverting(false)}
        />
      )}
    </div>
  )
}
