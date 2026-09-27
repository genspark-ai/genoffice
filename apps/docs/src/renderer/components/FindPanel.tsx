import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { TextSelection, type Transaction } from '@tiptap/pm/state'
import { useI18n, type StringKey } from '../i18n/locale'
import { searchPluginKey } from '../editor/extensions'
import {
  HIGHLIGHT_CAP,
  RESULTS_LIST_CAP,
  SCAN_THRESHOLD,
  compileFind,
  expandReplacement,
  findInDoc,
  findInDocAsync,
  matchContext,
  type FindMatch,
  type FindMode,
  type FindOptions,
} from '../editor/find'
import {
  GOTO_TARGETS,
  blockIndexAtPos,
  lineCounts,
  lineIndexAtPos,
  linePosition,
  listGoToItems,
  pageOfBlock,
  pagePositions,
  resolveGoToIndex,
  sectionPositions,
  type GoToTarget,
  type PageLayout,
} from '../editor/goto'

export { foldCase } from '../editor/find'

interface Range {
  from: number
  to: number
}

/** literal search over the editor (kept for callers that only need ranges) */
export function findMatches(
  editor: Editor,
  query: string,
  opts: Pick<FindOptions, 'matchCase' | 'wholeWord'>,
): Range[] {
  const compiled = compileFind(query, { ...opts, mode: 'literal', matchWidth: true })
  if (!compiled.ok) return []
  return findInDoc(editor.state.doc, compiled.query, { ...opts, matchWidth: true })
}

export type FindTab = 'find' | 'replace' | 'goto'

interface FindPanelProps {
  editor: Editor
  onClose: () => void
  /** Ctrl+F / menu Find bump this to put focus back in the find field while the panel is already open */
  focusFindNonce?: number
  /** Ctrl+H bumps this to land focus on the replace field (falls back to find when read-only) */
  focusReplaceNonce?: number
  /** Go To shortcut / menu: switch to the Go To tab */
  focusGoToNonce?: number
  /** last pagination pass, for page / section / line targets */
  layout?: () => PageLayout | null
}

const SCAN_DEBOUNCE_MS = 150

const ERROR_KEYS: Record<string, StringKey> = {
  'invalid-pattern': 'appFindErrInvalid',
  'trailing-backslash': 'appFindErrInvalid',
  'nothing-to-repeat': 'appFindErrInvalid',
  'bad-repeat': 'appFindErrInvalid',
  'unbalanced-paren': 'appFindErrParen',
  'unbalanced-bracket': 'appFindErrBracket',
  'empty-set': 'appFindErrBracket',
  'bad-caret-code': 'appFindErrCaret',
  'paragraph-mark-position': 'appFindErrParaMark',
}

const TARGET_KEYS: Record<GoToTarget, StringKey> = {
  page: 'appGoToPage',
  section: 'appGoToSection',
  line: 'appGoToLine',
  bookmark: 'appGoToBookmark',
  comment: 'appGoToComment',
  footnote: 'appGoToFootnote',
  endnote: 'appGoToEndnote',
  table: 'appGoToTable',
  heading: 'appGoToHeading',
}

/** index of the last item at or before the caret, -1 when none */
function lastAtOrBefore(positions: Array<number | null>, caret: number): number {
  let idx = -1
  positions.forEach((p, i) => {
    if (p !== null && p <= caret) idx = i
  })
  return idx
}

/** place the caret at `pos` (nearest valid text position) and bring it on screen */
function jumpTo(editor: Editor, pos: number): void {
  const { state, view } = editor
  const clamped = Math.max(0, Math.min(pos, state.doc.content.size))
  const sel = TextSelection.near(state.doc.resolve(clamped), 1)
  view.dispatch(state.tr.setSelection(sel).scrollIntoView())
  const { node } = view.domAtPos(sel.from)
  const el = node instanceof HTMLElement ? node : node.parentElement
  el?.scrollIntoView({ block: 'center' })
}

export function FindPanel({
  editor,
  onClose,
  focusFindNonce,
  focusReplaceNonce,
  focusGoToNonce,
  layout,
}: FindPanelProps) {
  const { t } = useI18n()
  const canEdit = editor.isEditable
  const [tab, setTab] = useState<FindTab>(() =>
    focusGoToNonce ? 'goto' : focusReplaceNonce && canEdit ? 'replace' : 'find',
  )
  const [query, setQuery] = useState('')
  const [replacement, setReplacement] = useState('')
  const [matches, setMatches] = useState<FindMatch[]>([])
  const [index, setIndex] = useState(0)
  const [matchCase, setMatchCase] = useState(false)
  const [wholeWord, setWholeWord] = useState(false)
  const [mode, setMode] = useState<FindMode>('literal')
  const [matchWidth, setMatchWidth] = useState(false)
  const [error, setError] = useState<StringKey | null>(null)
  const [showResults, setShowResults] = useState(false)
  const [scanning, setScanning] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)
  const replaceInputRef = useRef<HTMLInputElement>(null)
  const gotoInputRef = useRef<HTMLInputElement>(null)
  const indexRef = useRef(0)
  const matchesRef = useRef<FindMatch[]>([])
  const scanTokenRef = useRef<{ aborted: boolean } | null>(null)

  const options = useMemo<FindOptions>(
    () => ({ matchCase, wholeWord, mode, matchWidth }),
    [matchCase, wholeWord, mode, matchWidth],
  )

  const highlight = useCallback(
    (ranges: FindMatch[], activeIndex: number) => {
      const shown =
        ranges.length > HIGHLIGHT_CAP && ranges[activeIndex]
          ? { ranges: [ranges[activeIndex]], activeIndex: 0 }
          : { ranges, activeIndex }
      editor.view.dispatch(editor.state.tr.setMeta(searchPluginKey, shown))
    },
    [editor],
  )

  const scrollTo = useCallback(
    (range: Range) => {
      const { node } = editor.view.domAtPos(range.from)
      const el = node instanceof HTMLElement ? node : node.parentElement
      el?.scrollIntoView({ block: 'center' })
    },
    [editor],
  )

  const commit = useCallback(
    (ranges: FindMatch[], keepIndex: number) => {
      const active = ranges.length === 0 ? 0 : Math.min(keepIndex, ranges.length - 1)
      setMatches(ranges)
      matchesRef.current = ranges
      setIndex(active)
      indexRef.current = active
      highlight(ranges, active)
      return ranges
    },
    [highlight],
  )

  // rescan only updates matches/highlight; scrolling happens on explicit navigation.
  // Large documents scan in time slices; a newer scan aborts the one in flight.
  const refresh = useCallback(
    (q: string, keepIndex = 0, opts: FindOptions = options): FindMatch[] | null => {
      if (scanTokenRef.current) scanTokenRef.current.aborted = true
      scanTokenRef.current = null
      setScanning(false)
      // hits of a superseded query must never drive Next / Replace
      matchesRef.current = []
      if (!q) {
        setError(null)
        return commit([], 0)
      }
      const compiled = compileFind(q, opts)
      if (!compiled.ok) {
        setError(ERROR_KEYS[compiled.error] ?? 'appFindErrInvalid')
        return commit([], 0)
      }
      setError(null)
      const doc = editor.state.doc
      if (doc.nodeSize <= SCAN_THRESHOLD.nodeSize)
        return commit(findInDoc(doc, compiled.query, opts), keepIndex)
      const token = { aborted: false }
      scanTokenRef.current = token
      setScanning(true)
      setMatches([])
      void findInDocAsync(doc, compiled.query, opts, token).then((found) => {
        if (!found || scanTokenRef.current !== token) return
        scanTokenRef.current = null
        setScanning(false)
        commit(found, keepIndex)
      })
      return null
    },
    [editor, options, commit],
  )

  const refreshRef = useRef(refresh)
  refreshRef.current = refresh
  const timerRef = useRef<number | null>(null)
  const pendingKeepRef = useRef<'reset' | 'current'>('reset')
  const scheduleRefresh = useCallback((q: string, keep: 'reset' | 'current' = 'reset') => {
    if (timerRef.current !== null) window.clearTimeout(timerRef.current)
    pendingKeepRef.current = keep
    timerRef.current = window.setTimeout(() => {
      timerRef.current = null
      refreshRef.current(q, keep === 'current' ? indexRef.current : 0)
    }, SCAN_DEBOUNCE_MS)
  }, [])
  /** run a pending debounced scan now (Enter right after typing must see fresh matches) */
  const flushPending = useCallback(
    (q: string) => {
      if (timerRef.current === null) return null
      window.clearTimeout(timerRef.current)
      timerRef.current = null
      const keep = pendingKeepRef.current
      const ranges = refresh(q, keep === 'current' ? indexRef.current : 0)
      return ranges ? { ranges, queryChanged: keep === 'reset' } : null
    },
    [refresh],
  )

  // the target field may not be mounted yet when the tab is switching: retry next frame
  const focusTabInput = useCallback(
    (which: FindTab) => {
      const pick = (late: boolean) =>
        which === 'goto'
          ? gotoInputRef.current
          : which === 'replace' && canEdit
            ? (replaceInputRef.current ?? (late ? inputRef.current : null))
            : inputRef.current
      const focus = (el: HTMLElement | null) => {
        el?.focus()
        if (el instanceof HTMLInputElement) el.select()
        return !!el
      }
      if (!focus(pick(false))) requestAnimationFrame(() => focus(pick(true)))
    },
    [canEdit],
  )

  useEffect(() => {
    focusTabInput(tab)
    return () => {
      if (timerRef.current !== null) window.clearTimeout(timerRef.current)
      if (scanTokenRef.current) scanTokenRef.current.aborted = true
    }
    // mount only: later focus moves come from the nonces
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    if (!focusFindNonce) return
    setTab((cur) => (cur === 'goto' ? 'find' : cur))
    focusTabInput('find')
  }, [focusFindNonce, focusTabInput])

  // declared after the mount effect so opening straight into replace wins the focus
  useEffect(() => {
    if (!focusReplaceNonce) return
    setTab(canEdit ? 'replace' : 'find')
    focusTabInput('replace')
  }, [focusReplaceNonce, canEdit, focusTabInput])

  useEffect(() => {
    if (!focusGoToNonce) return
    setTab('goto')
    focusTabInput('goto')
  }, [focusGoToNonce, focusTabInput])

  // stay in sync while the document changes underneath (typing, AI edits).
  // The listener reads the query through a ref: between a keystroke in the find
  // box and the next render, the effect closure still holds the previous query
  // and would overwrite the pending scan for the new needle with stale text.
  const queryRef = useRef(query)
  queryRef.current = query
  useEffect(() => {
    const onUpdate = () => {
      if (queryRef.current) scheduleRefresh(queryRef.current, 'current')
    }
    editor.on('update', onUpdate)
    return () => {
      editor.off('update', onUpdate)
    }
  }, [editor, scheduleRefresh])

  const close = useCallback(() => {
    if (timerRef.current !== null) window.clearTimeout(timerRef.current)
    timerRef.current = null
    if (scanTokenRef.current) scanTokenRef.current.aborted = true
    highlight([], 0)
    onClose()
  }, [highlight, onClose])

  const setOption = (next: Partial<FindOptions>) => {
    const merged = { ...options, ...next }
    if (next.matchCase !== undefined) setMatchCase(next.matchCase)
    if (next.wholeWord !== undefined) setWholeWord(next.wholeWord)
    if (next.mode !== undefined) setMode(next.mode)
    if (next.matchWidth !== undefined) setMatchWidth(next.matchWidth)
    refresh(query, index, merged)
  }

  const goTo = useCallback(
    (ranges: FindMatch[], next: number) => {
      setIndex(next)
      indexRef.current = next
      highlight(ranges, next)
      scrollTo(ranges[next])
    },
    [highlight, scrollTo],
  )

  const step = useCallback(
    (dir: 1 | -1) => {
      const fresh = flushPending(query)
      if (scanTokenRef.current) return
      const ranges = fresh ? fresh.ranges : matchesRef.current
      if (ranges.length === 0) return
      // A flushed scan for a new query already landed on the first match — Enter should
      // visit it, not skip past it. A keep-current refresh (document changed underneath)
      // must still move in the requested direction from the refreshed position.
      const next = fresh?.queryChanged
        ? indexRef.current
        : (indexRef.current + dir + ranges.length) % ranges.length
      goTo(ranges, next)
    },
    [flushPending, query, goTo],
  )

  /** insert the replacement, splitting the block wherever the template had a paragraph mark */
  const applyReplacement = useCallback(
    (tr: Transaction, m: FindMatch) => {
      const segs = expandReplacement(replacement, m, mode)
      const joined = segs.join('\u0000')
      tr.insertText(joined, m.from, m.to)
      const splitsAt: number[] = []
      let offset = 0
      for (let k = 0; k < segs.length - 1; k++) {
        offset += segs[k].length
        splitsAt.push(m.from + offset)
        offset += 1
      }
      for (const p of splitsAt.reverse()) {
        tr.delete(p, p + 1)
        tr.split(p)
      }
    },
    [replacement, mode],
  )

  const replaceOne = useCallback(() => {
    if (!editor.isEditable) return
    // a pending debounced rescan means `matches` may describe the previous
    // query or pre-edit positions — replacing those would edit the wrong text
    const fresh = flushPending(query)
    if (scanTokenRef.current) return
    const ranges = fresh ? fresh.ranges : matchesRef.current
    const at = indexRef.current
    const m = ranges[at]
    if (!m) return
    editor.commands.command(({ tr }) => {
      applyReplacement(tr, m)
      return true
    })
    const after = refresh(query, at)
    if (after && after.length > 0) scrollTo(after[Math.min(at, after.length - 1)])
  }, [editor, flushPending, query, applyReplacement, refresh, scrollTo])

  const replaceAll = useCallback(() => {
    if (!editor.isEditable) return
    const fresh = flushPending(query)
    if (scanTokenRef.current) return
    const ranges = fresh ? fresh.ranges : matchesRef.current
    if (ranges.length === 0) return
    editor.commands.command(({ tr }) => {
      for (const m of [...ranges].reverse()) applyReplacement(tr, m)
      return true
    })
    refresh(query)
  }, [editor, flushPending, query, applyReplacement, refresh])

  // ---- Go To ----
  const [target, setTarget] = useState<GoToTarget>('page')
  const [gotoInput, setGotoInput] = useState('')
  const [gotoMiss, setGotoMiss] = useState(false)
  const doc = editor.state.doc
  const gotoItems = useMemo(
    () =>
      tab === 'goto' && !['page', 'section', 'line'].includes(target)
        ? listGoToItems(doc, target)
        : [],
    [tab, target, doc],
  )

  const runGoTo = (dir: 1 | -1, explicit = false) => {
    const lay = layout?.() ?? null
    const caret = editor.state.selection.from
    const view = editor.view
    let positions: Array<number | null>
    let current: number
    if (target === 'page' || target === 'section') {
      if (!lay) return setGotoMiss(true)
      positions = target === 'page' ? pagePositions(view, lay) : sectionPositions(view, lay)
      const bi = blockIndexAtPos(view, lay, caret)
      current =
        target === 'page'
          ? pageOfBlock(lay, Math.max(bi, 0)) - 1
          : Math.max(0, lastAtOrBefore(positions, caret))
    } else if (target === 'line') {
      if (!lay) return setGotoMiss(true)
      const total = lineCounts(lay).reduce((a, b) => a + b, 0)
      current = lineIndexAtPos(view, lay, caret)
      const idx = resolveGoToIndex(explicit ? gotoInput : '', current, total, dir)
      if (idx === null) return setGotoMiss(true)
      const pos = linePosition(view, lay, idx + 1)
      if (pos === null) return setGotoMiss(true)
      setGotoMiss(false)
      jumpTo(editor, pos)
      return
    } else {
      positions = gotoItems.map((i) => i.pos)
      current = lastAtOrBefore(positions, caret)
    }
    let idx: number | null
    if (target === 'bookmark' && explicit && gotoInput) {
      idx = gotoItems.findIndex((i) => i.label === gotoInput)
      if (idx < 0) idx = null
    } else idx = resolveGoToIndex(explicit ? gotoInput : '', current, positions.length, dir)
    const pos = idx === null ? null : positions[idx]
    if (pos === null || pos === undefined) return setGotoMiss(true)
    setGotoMiss(false)
    jumpTo(editor, pos)
  }

  const chip = (on: boolean, label: string, tip: StringKey, onClick: () => void) => (
    <button
      type="button"
      className={`find-opt ${on ? 'on' : ''}`}
      data-tip={t(tip)}
      aria-label={t(tip)}
      aria-pressed={on}
      onClick={onClick}
    >
      {label}
    </button>
  )

  const count = query
    ? scanning
      ? '…'
      : matches.length === 0
        ? t('appNoResults')
        : `${index + 1}/${matches.length}`
    : ''

  return (
    <div className="find-panel" role="dialog" aria-label={t('appFindPlaceholder')}>
      <div className="find-tabs" role="tablist">
        {(['find', ...(canEdit ? (['replace'] as const) : []), 'goto'] as FindTab[]).map((k) => (
          <button
            key={k}
            type="button"
            role="tab"
            aria-selected={tab === k}
            className={`find-tab ${tab === k ? 'on' : ''}`}
            onClick={() => {
              setTab(k)
              focusTabInput(k)
            }}
          >
            {t(
              k === 'find' ? 'appFindPlaceholder' : k === 'replace' ? 'appReplace' : 'appGoToTitle',
            )}
          </button>
        ))}
        <span className="find-spacer" />
        <button
          type="button"
          className="find-btn find-close"
          data-tip={t('appCloseEsc')}
          aria-label={t('appCloseEsc')}
          onClick={close}
        >
          ✕
        </button>
      </div>

      {tab !== 'goto' && (
        <>
          <div className="find-row">
            <input
              ref={inputRef}
              className="find-input"
              placeholder={t('appFindPlaceholder')}
              aria-label={t('appFindPlaceholder')}
              value={query}
              aria-invalid={!!error}
              onChange={(e) => {
                setQuery(e.target.value)
                queryRef.current = e.target.value
                scheduleRefresh(e.target.value)
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter') step(e.shiftKey ? -1 : 1)
                if (e.key === 'Escape') close()
              }}
            />
            <span className="find-count" aria-live="polite">
              {count}
            </span>
            <button
              type="button"
              className="find-btn"
              data-tip={t('appPrevMatch')}
              aria-label={t('appPrevMatch')}
              onClick={() => step(-1)}
              disabled={matches.length === 0 || scanning}
            >
              ‹
            </button>
            <button
              type="button"
              className="find-btn"
              data-tip={t('appNextMatch')}
              aria-label={t('appNextMatch')}
              onClick={() => step(1)}
              disabled={matches.length === 0 || scanning}
            >
              ›
            </button>
          </div>
          <div className="find-row find-options">
            {chip(matchCase, 'Aa', 'appMatchCase', () => setOption({ matchCase: !matchCase }))}
            {chip(wholeWord, 'W', 'appWholeWord', () => setOption({ wholeWord: !wholeWord }))}
            {chip(mode === 'wildcard', '*?', 'appUseWildcards', () =>
              setOption({ mode: mode === 'wildcard' ? 'literal' : 'wildcard' }),
            )}
            {chip(mode === 'regex', '.*', 'appUseRegex', () =>
              setOption({ mode: mode === 'regex' ? 'literal' : 'regex' }),
            )}
            {chip(matchWidth, 'Ａa', 'appMatchWidth', () => setOption({ matchWidth: !matchWidth }))}
            <span className="find-spacer" />
            {query && matches.length > 0 && (
              <button
                type="button"
                className={`find-opt find-results-toggle ${showResults ? 'on' : ''}`}
                aria-expanded={showResults}
                onClick={() => setShowResults((v) => !v)}
              >
                {t('appFindResults', { count: matches.length })}
              </button>
            )}
          </div>
          {error && (
            <div className="find-error" role="alert">
              {t(error)}
            </div>
          )}
          {tab === 'replace' && canEdit && (
            <div className="find-row">
              <input
                ref={replaceInputRef}
                className="find-input"
                placeholder={t('appReplacePlaceholder')}
                aria-label={t('appReplacePlaceholder')}
                value={replacement}
                onChange={(e) => setReplacement(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') replaceOne()
                  if (e.key === 'Escape') close()
                }}
              />
              <button
                type="button"
                className="find-action"
                onClick={replaceOne}
                disabled={matches.length === 0 || scanning}
              >
                {t('appReplace')}
              </button>
              <button
                type="button"
                className="find-action"
                onClick={replaceAll}
                disabled={matches.length === 0 || scanning}
              >
                {t('appReplaceAll')}
              </button>
            </div>
          )}
          {showResults && matches.length > 0 && (
            <div
              className="find-results"
              role="listbox"
              aria-label={t('appFindResults', { count: matches.length })}
            >
              {matches.slice(0, RESULTS_LIST_CAP).map((m, i) => {
                const ctx = matchContext(m)
                return (
                  <button
                    type="button"
                    key={`${m.from}-${i}`}
                    role="option"
                    aria-selected={i === index}
                    className={`find-result ${i === index ? 'on' : ''}`}
                    onClick={() => goTo(matches, i)}
                  >
                    <span className="find-result-ctx">{ctx.before}</span>
                    <mark>{ctx.hit}</mark>
                    <span className="find-result-ctx">{ctx.after}</span>
                  </button>
                )
              })}
              {matches.length > RESULTS_LIST_CAP && (
                <div className="find-results-note">
                  {t('appFindResultsCapped', { shown: RESULTS_LIST_CAP, total: matches.length })}
                </div>
              )}
            </div>
          )}
        </>
      )}

      {tab === 'goto' && (
        <>
          <div className="find-row">
            <select
              className="find-select"
              aria-label={t('appGoToWhat')}
              value={target}
              onChange={(e) => {
                setTarget(e.target.value as GoToTarget)
                setGotoInput('')
                setGotoMiss(false)
                focusTabInput('goto')
              }}
            >
              {GOTO_TARGETS.map((k) => (
                <option key={k} value={k}>
                  {t(TARGET_KEYS[k])}
                </option>
              ))}
            </select>
            {target === 'bookmark' ? (
              <select
                className="find-select find-goto-input"
                aria-label={t('appGoToBookmark')}
                value={gotoInput}
                onChange={(e) => {
                  setGotoInput(e.target.value)
                  setGotoMiss(false)
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') runGoTo(1, true)
                  if (e.key === 'Escape') close()
                }}
              >
                <option value="">{t('appGoToPickBookmark')}</option>
                {gotoItems.map((i) => (
                  <option key={`${i.pos}-${i.label}`} value={i.label}>
                    {i.label}
                  </option>
                ))}
              </select>
            ) : (
              <input
                ref={gotoInputRef}
                className="find-input find-goto-input"
                placeholder={t('appGoToNumberHint')}
                value={gotoInput}
                aria-invalid={gotoMiss}
                onChange={(e) => {
                  setGotoInput(e.target.value)
                  setGotoMiss(false)
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') runGoTo(e.shiftKey ? -1 : 1, true)
                  if (e.key === 'Escape') close()
                }}
              />
            )}
          </div>
          <div className="find-row">
            <span className={`find-hint ${gotoMiss ? 'find-hint-miss' : ''}`}>
              {gotoMiss ? t('appGoToNotFound') : t('appGoToRelativeHint')}
            </span>
            <span className="find-spacer" />
            <button type="button" className="find-action" onClick={() => runGoTo(-1)}>
              {t('appGoToPrevious')}
            </button>
            <button type="button" className="find-action" onClick={() => runGoTo(1)}>
              {t('appGoToNext')}
            </button>
            <button
              type="button"
              className="find-action find-action-primary"
              onClick={() => runGoTo(1, true)}
            >
              {t('appGoToTitle')}
            </button>
          </div>
        </>
      )}
    </div>
  )
}
