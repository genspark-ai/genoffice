import { useEffect, useMemo, useRef, useState } from 'react'

import { useI18n } from './i18n/locale'
import {
  collectWorkbookTextCells,
  indexOfWordOf,
  replaceInRichDoc,
  replaceSuggestion,
  runSheetSpellCheck,
  type RichDoc,
  type SpellCellRef,
  type SpellFinding,
  type SpellSnapshot,
} from './spell-check'
import { loadLearnedWords, saveLearnedWords } from './spellcheck-pref'

/**
 * Structural slice of Univer's FWorkbook/FWorksheet the dialog needs; keeps
 * the component off the full facade types (and unit-testable in principle).
 */
interface SpellWorkbookLike {
  getSheets(): { getSheetId(): string; getSheetName(): string }[] | null
  getSnapshot(): unknown
  getSheetBySheetId(sheetId: string): SpellWorksheetLike | null
}

interface SpellWorksheetLike {
  getRange(row: number, column: number): { setValue(value: string): unknown }
}

/** Stable identity for one finding row: the flagged word in one cell. */
const findingKey = (finding: SpellFinding): string =>
  `${finding.cell.sheetId}:${finding.cell.address}:${finding.word}`

/** Long cell texts are clipped for the table; the detail row shows more. */
const excerpt = (text: string, at: number, keep = 90): string => {
  if (text.length <= keep) return text
  const start = Math.max(0, Math.min(at - Math.floor(keep / 3), text.length - keep))
  return `${start > 0 ? '…' : ''}${text.slice(start, start + keep)}${start + keep < text.length ? '…' : ''}`
}

/** The live snapshot cell, as the replacement's starting point. */
function liveCell(
  snapshot: SpellSnapshot,
  sheetId: string,
  row: number,
  column: number,
): SpellCellRef | null {
  const cell = snapshot.sheets[sheetId]?.cellData?.[String(row)]?.[String(column)]
  if (!cell || typeof cell.v !== 'string') return null
  return { sheetId, sheetName: '', address: '', row, column, text: cell.v }
}

/**
 * Review ▸ Spelling / F7 (Excel's Spelling dialog, list form). The whole
 * workbook's string cells are collected locally and checked against the
 * frame's own dictionary — the same engine docs' squiggles run on; nothing
 * leaves the machine. Findings land in a clickable list that jumps to the
 * cell. Ignore (one cell), Ignore All (word, this workbook session), Add to
 * Dictionary (word, persisted) and Replace All mirror the docs spelling
 * menu's semantics — see spellcheck-pref.ts.
 */
export function SpellingDialog({
  onClose,
  getWorkbook,
  isWithheld,
  onJumpTo,
  lazyNote,
  sessionIgnoredRef,
  setMessage,
}: {
  readonly onClose: () => void
  /** active workbook getter, refreshed per run so a reopened file is picked up */
  readonly getWorkbook: () => SpellWorkbookLike | null
  /** withheld-cell gate: a redacted cell is never checked */
  readonly isWithheld: (sheetId: string, row: number, column: number) => boolean
  readonly onJumpTo: (sheetId: string, row: number, column: number) => void
  /** '' or the loaded-region-only note for streamed (lazy) workbooks */
  readonly lazyNote: string
  /** words dismissed with Ignore All this workbook session (lower-cased) */
  readonly sessionIgnoredRef: { readonly current: Set<string> }
  readonly setMessage: (message: string) => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [phase, setPhase] = useState<'done' | 'unavailable'>('done')
  const [findings, setFindings] = useState<readonly SpellFinding[]>([])
  const [selectedKey, setSelectedKey] = useState<string | null>(null)
  const [truncated, setTruncated] = useState(false)
  const learnedRef = useRef<Set<string>>(new Set())

  /** One synchronous pass over the workbook; bumped by Recheck and re-open. */
  const startRun = () => {
    setSelectedKey(null)
    setTruncated(false)
    setFindings([])
    const engine = window.nativeSpell
    const workbook = getWorkbook()
    const sheets = workbook?.getSheets()
    if (!engine || !workbook || !sheets || sheets.length === 0) {
      setPhase('unavailable')
      return
    }
    setPhase('done')
    learnedRef.current = new Set(loadLearnedWords().map((word) => word.toLowerCase()))
    // the runtime snapshot is a structural superset of SpellSnapshot; the
    // cast only erases Univer's branded matrix types, the shape is the same
    const { cells, truncated: capped } = collectWorkbookTextCells(
      workbook.getSnapshot() as SpellSnapshot,
      sheets.map((sheet) => ({
        sheetId: sheet.getSheetId(),
        sheetName: sheet.getSheetName(),
      })),
      isWithheld,
    )
    setTruncated(capped)
    setFindings(runSheetSpellCheck(cells, engine, learnedRef.current, sessionIgnoredRef.current))
  }

  useEffect(() => {
    startRun()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const dropFinding = (key: string): void => {
    setFindings((previous) => previous.filter((finding) => findingKey(finding) !== key))
    setSelectedKey((current) => (current === key ? null : current))
  }

  const ignoreAll = (word: string): void => {
    sessionIgnoredRef.current.add(word.toLowerCase())
    setFindings((previous) => previous.filter((finding) => finding.word !== word))
    setSelectedKey(null)
    setMessage(t('appSpellIgnoredAll', { word }))
  }

  const learnWord = (word: string): void => {
    const lowered = word.toLowerCase()
    learnedRef.current.add(lowered)
    saveLearnedWords([...learnedRef.current])
    setFindings((previous) => previous.filter((finding) => finding.word !== word))
    setSelectedKey(null)
    setMessage(t('appSpellLearned', { word }))
  }

  /**
   * Writes one replacement into a cell, against its CURRENT value — the user
   * may have edited it since the scan — through Univer so undo, journal and
   * dirty state all see it. A rich-text cell (cell.p) keeps its runs: the
   * replacement rewrites the document's dataStream and shifts the runs, the
   * shape Excel's own in-cell corrector leaves behind.
   */
  const applyToCell = (finding: SpellFinding, suggestion: string): boolean => {
    const workbook = getWorkbook()
    const worksheet = workbook?.getSheetBySheetId(finding.cell.sheetId)
    if (!workbook || !worksheet) return false
    const snapshot = workbook.getSnapshot() as SpellSnapshot
    const live = liveCell(snapshot, finding.cell.sheetId, finding.cell.row, finding.cell.column)
    if (live === null) return false
    const rich = (
      snapshot.sheets[finding.cell.sheetId]?.cellData?.[String(finding.cell.row)]?.[
        String(finding.cell.column)
      ] as { p?: RichDoc } | undefined
    )?.p
    if (rich !== undefined && rich !== null) {
      const doc = replaceInRichDoc(rich, finding.word, suggestion)
      if (doc === null) return false
      // setValues takes the full cell object (ICellData), which is what
      // carries the document; the cast only erases the facade's branded types
      const range = worksheet.getRange(finding.cell.row, finding.cell.column) as unknown as {
        setValues(values: { v: string; p: RichDoc }[][]): unknown
      }
      range.setValues([[{ v: replaceSuggestion(live.text, finding.word, suggestion), p: doc }]])
      return true
    }
    const next = replaceSuggestion(live.text, finding.word, suggestion)
    if (next === live.text) return false
    worksheet.getRange(finding.cell.row, finding.cell.column).setValue(next)
    return true
  }

  const applySuggestion = (finding: SpellFinding, suggestion: string): void => {
    if (applyToCell(finding, suggestion)) {
      dropFinding(findingKey(finding))
      setMessage(t('appSpellChanged', { address: finding.cell.address }))
    } else {
      setMessage(t('appSpellChangeFailed', { address: finding.cell.address }))
    }
  }

  const replaceAll = (finding: SpellFinding, suggestion: string): void => {
    const same = findings.filter((entry) => entry.word === finding.word)
    let changed = 0
    for (const entry of same) {
      if (applyToCell(entry, suggestion)) changed += 1
    }
    if (changed > 0) {
      setFindings((previous) => previous.filter((entry) => entry.word !== finding.word))
      setSelectedKey(null)
      setMessage(t('appSpellChangedAll', { word: finding.word, count: String(changed) }))
    } else {
      setMessage(t('appSpellChangeFailed', { address: finding.cell.address }))
    }
  }

  const selected = useMemo(
    () => findings.find((finding) => findingKey(finding) === selectedKey) ?? null,
    [findings, selectedKey],
  )

  const statusLine = () => {
    if (phase === 'unavailable') return t('appSpellUnavailable')
    if (findings.length === 0) return t('appSpellNone')
    return t('appSpellFindings', { count: String(findings.length) })
  }

  return (
    <div
      className="find-replace-panel"
      role="dialog"
      aria-label={t('appSpellTitle')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') {
          event.stopPropagation()
          onClose()
        }
      }}
    >
      <header>
        {t('appSpellTitle')}
        <button type="button" className="fr-close" aria-label={t('dlgClose')} onClick={onClose}>
          ×
        </button>
      </header>
      <div className="fr-status" role="status">
        {statusLine()}
        {lazyNote ? ` ${lazyNote}` : ''}
        {truncated ? ` ${t('appSpellTruncated')}` : ''}
      </div>
      {findings.length > 0 && (
        <div className="fr-results">
          <table>
            <thead>
              <tr>
                <th>{t('dlgFrResultCell')}</th>
                <th>{t('dlgFrResultValue')}</th>
                <th>{t('appSpellColWord')}</th>
              </tr>
            </thead>
            <tbody>
              {findings.map((finding) => {
                const key = findingKey(finding)
                return (
                  <tr
                    key={key}
                    className={key === selectedKey ? 'spell-selected' : undefined}
                    onClick={() => {
                      setSelectedKey(key)
                      onJumpTo(finding.cell.sheetId, finding.cell.row, finding.cell.column)
                    }}
                  >
                    <td>{`${finding.cell.sheetName}!${finding.cell.address}`}</td>
                    <td>
                      {excerpt(finding.cell.text, indexOfWordOf(finding.cell.text, finding.word))}
                    </td>
                    <td>{finding.word}</td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      )}
      {selected && (
        <div className="spell-detail">
          <div className="spell-detail-text">
            <strong>{`${selected.cell.sheetName}!${selected.cell.address}`}</strong>
            <span>
              {excerpt(selected.cell.text, indexOfWordOf(selected.cell.text, selected.word), 160)}
            </span>
          </div>
          <div className="spell-suggestions">
            {selected.suggestions.length === 0 && (
              <span className="spell-no-suggestion">{t('appSpellNoSuggestions')}</span>
            )}
            {selected.suggestions.map((suggestion) => (
              <button
                key={suggestion}
                type="button"
                className="spell-suggestion"
                onClick={() => applySuggestion(selected, suggestion)}
              >
                {suggestion}
              </button>
            ))}
          </div>
          <div className="spell-actions">
            <button type="button" onClick={() => dropFinding(findingKey(selected))}>
              {t('appSpellIgnore')}
            </button>
            <button type="button" onClick={() => ignoreAll(selected.word)}>
              {t('appSpellIgnoreAll')}
            </button>
            <button type="button" onClick={() => learnWord(selected.word)}>
              {t('appSpellLearn')}
            </button>
            {selected.suggestions[0] !== undefined && (
              <button type="button" onClick={() => replaceAll(selected, selected.suggestions[0]!)}>
                {t('appSpellReplaceAll', { suggestion: selected.suggestions[0] })}
              </button>
            )}
          </div>
        </div>
      )}
      <div className="dialog-actions fr-actions">
        <button type="button" onClick={startRun}>
          {t('appSpellRecheck')}
        </button>
        <button type="button" onClick={onClose}>
          {t('dlgClose')}
        </button>
      </div>
    </div>
  )
}
