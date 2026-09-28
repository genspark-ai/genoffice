import { useEffect, useMemo, useRef, useState } from 'react'
import type { Observable } from 'rxjs'
import type { Workbook } from '@univerjs/core'
import {
  createInitFindReplaceState,
  FindBy,
  FindDirection,
  FindScope,
  type IFindReplaceService,
  type IFindReplaceState,
  type IReplaceAllResult,
} from '@univerjs/find-replace'
import type { RangeBounds } from '@genoffice/xlsx-gateway/domain/cell-address'

import { useI18n } from './i18n/locale'
import {
  buildFindAllRows,
  FIND_HISTORY_KEY,
  getSessionMatches,
  hasUncommittedQuery,
  hideReplace,
  loadHistory,
  pushHistory,
  REPLACE_HISTORY_KEY,
} from './find-replace-session'

/// Excel caps its own list around the tens of thousands; 500 keeps the DOM
/// light while the footer still reports the real total.
const RESULTS_DISPLAY_LIMIT = 500

function useObservableValue<T>(observable: Observable<T>, initial: T): T {
  const [value, setValue] = useState<T>(initial)
  useEffect(() => {
    const subscription = observable.subscribe(setValue)
    return () => subscription.unsubscribe()
  }, [observable])
  return value
}

/**
 * Session state with trustworthy `revealed` flags. The state observable's
 * INITIAL snapshot claims `revealed: true` before any session starts (Univer's
 * `createInitFindReplaceState`), so the panel reads the service getters —
 * which track the real `_revealed` fields — on every emission.
 */
function useFindReplaceState(service: IFindReplaceService): IFindReplaceState {
  const [value, setValue] = useState<IFindReplaceState>(() => ({
    ...createInitFindReplaceState(),
    revealed: service.revealed,
    replaceRevealed: service.replaceRevealed,
  }))
  useEffect(() => {
    const subscription = service.state$.subscribe((state) =>
      setValue({
        ...state,
        revealed: service.revealed,
        replaceRevealed: service.replaceRevealed,
      }),
    )
    return () => subscription.unsubscribe()
  }, [service])
  return value
}

/**
 * Excel's Find and Replace dialog (Ctrl+F / Ctrl+H): a modeless, draggable
 * panel with Find/Replace tabs, the full option set Univer's service
 * supports, a Find All results table and input history. The stock Univer
 * dialog is suppressed; sessions and matching stay on IFindReplaceService,
 * so streamed-workbook coverage and replace semantics are unchanged.
 */
export function FindReplacePanel({
  service,
  getWorkbook,
  onJumpTo,
  onClose,
  registerContainer,
}: {
  readonly service: IFindReplaceService
  readonly getWorkbook: () => Workbook | null
  readonly onJumpTo: (sheetId: string, bounds: RangeBounds) => void
  readonly onClose: () => void
  /**
   * Registers the panel root with Univer's layout service so focusing its
   * inputs does not count as leaving the workbook — otherwise the find
   * controller closes the session on the first keystroke.
   */
  readonly registerContainer: (element: HTMLElement) => () => void
}): React.JSX.Element | null {
  const { t } = useI18n()
  const state = useFindReplaceState(service)
  const currentMatch = useObservableValue(service.currentMatch$, null)
  const replaceables = useObservableValue(service.replaceables$, [])

  const [showResults, setShowResults] = useState(false)
  const [replaceReport, setReplaceReport] = useState<IReplaceAllResult | null>(null)
  const [findHistory, setFindHistory] = useState<string[]>(() =>
    loadHistory(globalThis.localStorage, FIND_HISTORY_KEY),
  )
  const [replaceHistory, setReplaceHistory] = useState<string[]>(() =>
    loadHistory(globalThis.localStorage, REPLACE_HISTORY_KEY),
  )
  const [position, setPosition] = useState<{ x: number; y: number } | null>(null)

  const findInputRef = useRef<HTMLInputElement>(null)
  const panelRef = useRef<HTMLDivElement>(null)
  const composingRef = useRef(false)
  const liveSearchTimer = useRef<number | null>(null)

  const revealed = state.revealed
  const replaceTab = state.replaceRevealed

  useEffect(() => {
    if (!revealed) return
    const element = panelRef.current
    if (!element) return
    return registerContainer(element)
  }, [revealed, registerContainer])

  // (Re)focus the query input whenever the panel opens or the tab changes.
  useEffect(() => {
    if (revealed) findInputRef.current?.focus()
  }, [revealed, replaceTab])

  const cancelLiveSearch = (): void => {
    if (liveSearchTimer.current === null) return
    window.clearTimeout(liveSearchTimer.current)
    liveSearchTimer.current = null
  }

  // A new panel session starts clean; a pending Find-tab search must not
  // fire into a closed session.
  useEffect(() => {
    if (!revealed) {
      if (liveSearchTimer.current !== null) window.clearTimeout(liveSearchTimer.current)
      liveSearchTimer.current = null
      setShowResults(false)
      setReplaceReport(null)
      setPosition(null)
    }
  }, [revealed])

  useEffect(
    () => () => {
      if (liveSearchTimer.current !== null) window.clearTimeout(liveSearchTimer.current)
    },
    [],
  )

  const onQueryChange = (value: string): void => {
    setReplaceReport(null)
    service.changeInputtingFindString(value)
    // Always drop the previous timer: a Find-tab search left pending across a
    // tab switch would otherwise commit an older value on top of this one.
    cancelLiveSearch()
    if (replaceTab) return // the replace-autosearch shim handles live search
    // Find tab: live-search like the stock dialog, but not mid-IME-composition.
    liveSearchTimer.current = window.setTimeout(() => {
      liveSearchTimer.current = null
      if (!composingRef.current) service.changeFindString(value)
    }, 400)
  }

  const showFindTab = (): void => {
    cancelLiveSearch()
    hideReplace(service)
  }

  /// Switching to Replace keeps the typed query: Univer's `revealReplace`
  /// resets the box to the searched string, so text still waiting on the
  /// live-search timer is typed back afterwards and the replace-autosearch
  /// shim commits and searches it (past Univer's 200ms research throttle,
  /// which a direct changeFindString + find here could trip over).
  const showReplaceTab = (): void => {
    cancelLiveSearch()
    const typed = state.inputtingFindString
    service.revealReplace()
    if (typed !== service.getFindString()) service.changeInputtingFindString(typed)
  }

  const rememberQueries = (): void => {
    setFindHistory(
      pushHistory(globalThis.localStorage, FIND_HISTORY_KEY, state.inputtingFindString),
    )
    if (replaceTab && (state.replaceString ?? '') !== '') {
      setReplaceHistory(
        pushHistory(globalThis.localStorage, REPLACE_HISTORY_KEY, state.replaceString ?? ''),
      )
    }
  }

  /// Mirrors the stock dialog's Find button: commit a pending query first,
  /// then step the session.
  const findStep = (direction: 'next' | 'previous'): void => {
    if (state.inputtingFindString.trim() === '') return
    rememberQueries()
    if (state.findString === state.inputtingFindString) {
      if (direction === 'next') service.moveToNextMatch()
      else service.moveToPreviousMatch()
    } else {
      service.changeFindString(state.inputtingFindString)
      service.find()
    }
  }

  const onFindAll = (): void => {
    if (state.inputtingFindString.trim() === '') return
    rememberQueries()
    if (state.findString !== state.inputtingFindString) {
      service.changeFindString(state.inputtingFindString)
      service.find()
    }
    setShowResults(true)
  }

  /// Only a searched query's matches get rewritten. The buttons are disabled
  /// while the Find what box differs from the searched string (the
  /// replace-autosearch shim searches it within 300ms); this guard covers
  /// the moment between an edit and the re-render.
  const onReplace = (): void => {
    if (hasUncommittedQuery(state)) return
    rememberQueries()
    void service.replace()
  }

  const onReplaceAll = (): void => {
    if (hasUncommittedQuery(state)) return
    rememberQueries()
    void service.replaceAll().then(setReplaceReport)
  }

  const onHeaderPointerDown = (event: React.PointerEvent<HTMLElement>): void => {
    if (event.button !== 0 || (event.target as HTMLElement).closest('button')) return
    const panel = (event.currentTarget as HTMLElement).parentElement
    if (!panel) return
    const rect = panel.getBoundingClientRect()
    const offsetX = event.clientX - rect.left
    const offsetY = event.clientY - rect.top
    const onMove = (move: PointerEvent): void => {
      setPosition({
        x: Math.min(Math.max(move.clientX - offsetX, 8), window.innerWidth - 120),
        y: Math.min(Math.max(move.clientY - offsetY, 8), window.innerHeight - 60),
      })
    }
    const onUp = (): void => {
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', onUp)
    }
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', onUp)
  }

  const results = useMemo(() => {
    if (!showResults || !revealed) return null
    const workbook = getWorkbook()
    if (!workbook) return null
    return buildFindAllRows(getSessionMatches(service), workbook, RESULTS_DISPLAY_LIMIT)
    // replaceables/matchesCount change whenever the session's matches do,
    // so they stand in for the (non-observable) match list itself.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    showResults,
    revealed,
    getWorkbook,
    service,
    state.matchesCount,
    state.findCompleted,
    replaceables,
  ])

  if (!revealed) return null

  const findDisabled = state.inputtingFindString.trim() === ''
  const queryPending = hasUncommittedQuery(state)
  const replaceDisabled =
    queryPending || state.matchesCount === 0 || currentMatch?.replaceable !== true
  const replaceAllDisabled = queryPending || replaceables.length === 0
  const matchesLabel =
    state.matchesCount > 0
      ? t('dlgFrMatches', {
          position: state.matchesPosition > 0 ? String(state.matchesPosition) : '–',
          count: String(state.matchesCount),
        })
      : state.findCompleted && state.findString !== ''
        ? t('dlgFrNoMatch')
        : ''

  return (
    <div
      ref={panelRef}
      className="find-replace-panel"
      role="dialog"
      aria-label={t('dlgFrTitle')}
      style={position ? { left: position.x, top: position.y, right: 'auto' } : undefined}
      onKeyDown={(event) => {
        if (event.key === 'Escape') {
          event.stopPropagation()
          onClose()
        }
      }}
    >
      <header onPointerDown={onHeaderPointerDown}>
        {t('dlgFrTitle')}
        <button type="button" className="fr-close" aria-label={t('dlgClose')} onClick={onClose}>
          ×
        </button>
      </header>
      <div className="dialog-tabs" role="tablist">
        <button
          type="button"
          role="tab"
          aria-selected={!replaceTab}
          className={replaceTab ? '' : 'active'}
          onClick={showFindTab}
        >
          {t('dlgFrTabFind')}
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={replaceTab}
          className={replaceTab ? 'active' : ''}
          onClick={showReplaceTab}
        >
          {t('dlgFrTabReplace')}
        </button>
      </div>
      <div className="fr-grid">
        <label>
          {t('dlgFrFindWhat')}
          <input
            ref={findInputRef}
            list="fr-find-history"
            value={state.inputtingFindString}
            onChange={(event) => onQueryChange(event.target.value)}
            onCompositionStart={() => {
              composingRef.current = true
            }}
            onCompositionEnd={(event) => {
              composingRef.current = false
              onQueryChange((event.target as HTMLInputElement).value)
            }}
            onKeyDown={(event) => {
              if (event.key === 'Enter') findStep(event.shiftKey ? 'previous' : 'next')
            }}
          />
        </label>
        <datalist id="fr-find-history">
          {findHistory.map((entry) => (
            <option key={entry} value={entry} />
          ))}
        </datalist>
        {replaceTab && (
          <>
            <label>
              {t('dlgFrReplaceWith')}
              <input
                list="fr-replace-history"
                value={state.replaceString ?? ''}
                onChange={(event) => {
                  setReplaceReport(null)
                  service.changeReplaceString(event.target.value)
                }}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') findStep(event.shiftKey ? 'previous' : 'next')
                }}
              />
            </label>
            <datalist id="fr-replace-history">
              {replaceHistory.map((entry) => (
                <option key={entry} value={entry} />
              ))}
            </datalist>
          </>
        )}
        <div className="fr-options">
          <label>
            {t('dlgFrScope')}
            <select
              value={state.findScope}
              onChange={(event) => service.changeFindScope(event.target.value as FindScope)}
            >
              <option value={FindScope.SUBUNIT}>{t('dlgFrScopeSheet')}</option>
              <option value={FindScope.UNIT}>{t('dlgFrScopeWorkbook')}</option>
            </select>
          </label>
          <label>
            {t('dlgFrOrder')}
            <select
              value={state.findDirection}
              onChange={(event) => service.changeFindDirection(event.target.value as FindDirection)}
            >
              <option value={FindDirection.ROW}>{t('dlgFrOrderRow')}</option>
              <option value={FindDirection.COLUMN}>{t('dlgFrOrderColumn')}</option>
            </select>
          </label>
          <label>
            {t('dlgFrLookIn')}
            <select
              value={state.findBy}
              onChange={(event) => service.changeFindBy(event.target.value as FindBy)}
            >
              <option value={FindBy.VALUE}>{t('dlgFrLookInValue')}</option>
              <option value={FindBy.FORMULA}>{t('dlgFrLookInFormula')}</option>
            </select>
          </label>
        </div>
        <div className="fr-checks">
          <label className="dialog-check">
            <input
              type="checkbox"
              checked={state.caseSensitive}
              onChange={(event) => service.changeCaseSensitive(event.target.checked)}
            />
            {t('dlgFrMatchCase')}
          </label>
          <label className="dialog-check">
            <input
              type="checkbox"
              checked={state.matchesTheWholeCell}
              onChange={(event) => service.changeMatchesTheWholeCell(event.target.checked)}
            />
            {t('dlgFrMatchCell')}
          </label>
        </div>
      </div>
      <div className="fr-status" role="status">
        {replaceReport
          ? t('dlgFrReplacedAll', {
              count: String(replaceReport.success),
            }) +
            (replaceReport.failure > 0
              ? ` (${t('dlgFrReplaceFailed', { count: String(replaceReport.failure) })})`
              : '')
          : matchesLabel}
      </div>
      <div className="dialog-actions fr-actions">
        <button type="button" disabled={findDisabled} onClick={onFindAll}>
          {t('dlgFrFindAll')}
        </button>
        <button type="button" disabled={findDisabled} onClick={() => findStep('previous')}>
          {t('dlgFrFindPrev')}
        </button>
        <button type="button" disabled={findDisabled} onClick={() => findStep('next')}>
          {t('dlgFrFindNext')}
        </button>
        {replaceTab && (
          <>
            <button type="button" disabled={replaceDisabled} onClick={onReplace}>
              {t('dlgFrReplace')}
            </button>
            <button type="button" disabled={replaceAllDisabled} onClick={onReplaceAll}>
              {t('dlgFrReplaceAll')}
            </button>
          </>
        )}
        <button type="button" onClick={onClose}>
          {t('dlgClose')}
        </button>
      </div>
      {showResults && results && (
        <div className="fr-results">
          <table>
            <thead>
              <tr>
                <th>{t('dlgFrResultSheet')}</th>
                <th>{t('dlgFrResultCell')}</th>
                <th>{t('dlgFrResultValue')}</th>
              </tr>
            </thead>
            <tbody>
              {results.rows.map((row) => (
                <tr
                  key={`${row.sheetId}:${row.address}`}
                  onClick={() => onJumpTo(row.sheetId, row.bounds)}
                >
                  <td>{row.sheetName}</td>
                  <td>{row.address}</td>
                  <td>{row.content}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <div className="fr-results-count">
            {results.total > results.rows.length
              ? t('dlgFrResultsTruncated', {
                  shown: String(results.rows.length),
                  count: String(results.total),
                })
              : t('dlgFrResultsCount', { count: String(results.total) })}
          </div>
        </div>
      )}
    </div>
  )
}
