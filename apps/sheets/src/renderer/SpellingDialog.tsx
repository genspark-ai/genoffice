import { useEffect, useMemo, useRef, useState } from 'react'
import type { AiSettings } from '@genoffice/ai-provider'

import { useI18n } from './i18n/locale'
import {
  collectSheetTextCells,
  replaceSuggestion,
  runSheetSpellCheck,
  type SpellFinding,
  type SpellSnapshot,
} from './spell-check'
import { loadLearnedWords, saveLearnedWords } from './spellcheck-pref'

/**
 * Structural slice of Univer's FWorkbook/FWorksheet the dialog needs; keeps
 * the component off the full facade types (and unit-testable in principle).
 */
interface SpellWorkbookLike {
  getActiveSheet(): { getSheetId(): string; getSheetName(): string } | null
  getSnapshot(): unknown
  getSheetBySheetId(
    sheetId: string,
  ): { getRange(row: number, column: number): { setValue(value: string): unknown } } | null
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

/**
 * Review ▸ Spelling / F7 (Excel's Spelling dialog, list form). The active
 * worksheet's string cells are collected locally, deduplicated into batches
 * and checked through the AI provider; findings land in a clickable list that
 * jumps to the cell. Ignore (one cell), Ignore All (word, this workbook
 * session) and Add to Dictionary (word, persisted) mirror the docs spelling
 * menu's semantics — see spellcheck-pref.ts for why the native Chromium
 * checker cannot drive any of this in the canvas-rendered Univer editor.
 */
export function SpellingDialog({
  onClose,
  getWorkbook,
  isWithheld,
  aiConfigured,
  getSettings,
  onJumpTo,
  lazyNote,
  sessionIgnoredRef,
  setMessage,
}: {
  readonly onClose: () => void
  /** active workbook getter, refreshed per run so a reopened file is picked up */
  readonly getWorkbook: () => SpellWorkbookLike | null
  /** withheld-cell gate: a redacted cell never leaves the app for checking */
  readonly isWithheld: (sheetId: string, row: number, column: number) => boolean
  readonly aiConfigured: () => boolean
  readonly getSettings: () => AiSettings | null
  readonly onJumpTo: (sheetId: string, row: number, column: number) => void
  /** '' or the loaded-region-only note for streamed (lazy) workbooks */
  readonly lazyNote: string
  /** words dismissed with Ignore All this workbook session (lower-cased) */
  readonly sessionIgnoredRef: { readonly current: Set<string> }
  readonly setMessage: (message: string) => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [phase, setPhase] = useState<'scanning' | 'done' | 'canceled' | 'error' | 'unconfigured'>(
    'scanning',
  )
  const [findings, setFindings] = useState<readonly SpellFinding[]>([])
  const [progress, setProgress] = useState({ done: 0, total: 0 })
  const [selectedKey, setSelectedKey] = useState<string | null>(null)
  const [truncated, setTruncated] = useState(false)
  const [errorText, setErrorText] = useState('')
  const [runId, setRunId] = useState(0)
  const runIdRef = useRef(0)
  const learnedRef = useRef<Set<string>>(new Set())
  runIdRef.current = runId

  /** One check pass over the active sheet; bumped by Recheck and re-open. */
  const startRun = () => {
    const id = runIdRef.current + 1
    setRunId(id)
    setFindings([])
    setSelectedKey(null)
    setProgress({ done: 0, total: 0 })
    setTruncated(false)
    setErrorText('')
    if (!aiConfigured()) {
      setPhase('unconfigured')
      return
    }
    const settings = getSettings()
    const workbook = getWorkbook()
    const worksheet = workbook?.getActiveSheet()
    if (!workbook || !worksheet || !settings) {
      setPhase('done')
      return
    }
    learnedRef.current = new Set(loadLearnedWords().map((word) => word.toLowerCase()))
    // the runtime snapshot is a structural superset of SpellSnapshot; the
    // cast only erases Univer's branded matrix types, the shape is the same
    const cells = collectSheetTextCells(
      workbook.getSnapshot() as SpellSnapshot,
      worksheet.getSheetId(),
      worksheet.getSheetName(),
      isWithheld,
    )
    setPhase('scanning')
    void runSheetSpellCheck({
      cells,
      learnedWords: learnedRef.current,
      ignoredWords: sessionIgnoredRef.current,
      requestChat: (system, user) =>
        window.desktopApi.aiChat({ settings, system, user }).then((reply) => ({
          ok: reply.ok,
          ...(reply.content !== undefined ? { content: reply.content } : {}),
          ...(reply.error !== undefined ? { error: reply.error } : {}),
        })),
      onProgress: (done, total) => {
        if (runIdRef.current === id) setProgress({ done, total })
      },
      onFindings: (found) => {
        if (runIdRef.current === id) setFindings((previous) => [...previous, ...found])
      },
      shouldContinue: () => runIdRef.current === id,
    }).then((result) => {
      if (runIdRef.current !== id) return
      if (result.status === 'error') {
        setErrorText(result.error)
        setPhase('error')
      } else {
        if (result.status === 'done' && result.truncated) setTruncated(true)
        setPhase(result.status === 'canceled' ? 'canceled' : 'done')
      }
    })
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

  /** Applies one suggestion through Univer so undo, journal and dirty state all see it. */
  const applySuggestion = (finding: SpellFinding, suggestion: string): void => {
    const workbook = getWorkbook()
    const worksheet = workbook?.getSheetBySheetId(finding.cell.sheetId)
    if (!worksheet) return
    const next = replaceSuggestion(finding.cell.text, finding.word, suggestion)
    try {
      worksheet.getRange(finding.cell.row, finding.cell.column).setValue(next)
    } catch {
      setMessage(t('appSpellChangeFailed', { address: finding.cell.address }))
      return
    }
    dropFinding(findingKey(finding))
    setMessage(t('appSpellChanged', { address: finding.cell.address }))
  }

  const selected = useMemo(
    () => findings.find((finding) => findingKey(finding) === selectedKey) ?? null,
    [findings, selectedKey],
  )

  const statusLine = () => {
    if (phase === 'scanning') {
      return progress.total > 0
        ? t('appSpellProgress', { done: String(progress.done), total: String(progress.total) })
        : t('appSpellCollecting')
    }
    if (phase === 'unconfigured') return t('appSpellUnconfigured')
    if (phase === 'error') return t('appSpellError', { error: errorText })
    if (phase === 'canceled') return t('appSpellCanceled')
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
      <div className="spell-note">{t('appSpellAiNote')}</div>
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
                    <td>{finding.cell.address}</td>
                    <td>{excerpt(finding.cell.text, finding.cell.text.indexOf(finding.word))}</td>
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
            <strong>{selected.cell.address}</strong>
            <span>
              {excerpt(selected.cell.text, selected.cell.text.indexOf(selected.word), 160)}
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
          </div>
        </div>
      )}
      <div className="dialog-actions fr-actions">
        {phase === 'scanning' ? (
          <button type="button" onClick={() => setRunId(runIdRef.current + 1)}>
            {t('appSpellStop')}
          </button>
        ) : (
          <button type="button" onClick={startRun}>
            {t('appSpellRecheck')}
          </button>
        )}
        <button type="button" onClick={onClose}>
          {t('dlgClose')}
        </button>
      </div>
    </div>
  )
}
