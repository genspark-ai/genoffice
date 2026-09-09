import { useCallback, useEffect, useRef, useState } from 'react'
import type { FindOptions } from './find-text'

export interface FindPanelStrings {
  findPlaceholder: string
  replacePlaceholder: string
  matchCase: string
  wholeWord: string
  noResults: string
  prevMatch: string
  nextMatch: string
  closeEsc: string
  replace: string
  replaceAll: string
}

/**
 * Editor adapter the panel drives. The target owns the match list of the last
 * `search`; indexes passed to the other methods refer to that list.
 */
export interface FindTarget {
  readonly editable: boolean
  /** rescan and repaint the highlights; returns the match count */
  search(query: string, opts: FindOptions, activeIndex: number): number
  /** move the active highlight to `index` and scroll it into view */
  activate(index: number): void
  replaceOne(index: number, replacement: string): void
  replaceAll(replacement: string): void
  /** drop every highlight */
  clear(): void
  /** fires after any document change so the panel can rescan */
  onDocChanged(listener: () => void): () => void
}

/** bumped by the host on every Ctrl+F / Ctrl+H so an already-open panel refocuses its field */
export interface FindFocusRequest {
  field: 'find' | 'replace'
  nonce: number
}

interface FindPanelProps {
  target: FindTarget
  strings: FindPanelStrings
  onClose: () => void
  focusRequest?: FindFocusRequest
}

const SCAN_DEBOUNCE_MS = 150

export function FindPanel({ target, strings, onClose, focusRequest }: FindPanelProps) {
  const [query, setQuery] = useState('')
  const [replacement, setReplacement] = useState('')
  const [count, setCount] = useState(0)
  const [index, setIndex] = useState(0)
  const [matchCase, setMatchCase] = useState(false)
  const [wholeWord, setWholeWord] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)
  const replaceInputRef = useRef<HTMLInputElement>(null)
  const indexRef = useRef(0)
  const canEdit = target.editable

  // rescan only updates the count/highlight; scrolling happens on explicit navigation
  const refresh = useCallback(
    (q: string, keepIndex = 0, opts?: Partial<FindOptions>) => {
      const n = target.search(q, { matchCase, wholeWord, ...opts }, keepIndex)
      const active = n === 0 ? 0 : Math.min(keepIndex, n - 1)
      setCount(n)
      setIndex(active)
      indexRef.current = active
      return n
    },
    [target, matchCase, wholeWord],
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
      return {
        count: refresh(q, keep === 'current' ? indexRef.current : 0),
        queryChanged: keep === 'reset',
      }
    },
    [refresh],
  )

  const targetRef = useRef(target)
  targetRef.current = target
  useEffect(() => {
    inputRef.current?.focus()
    inputRef.current?.select()
    return () => {
      if (timerRef.current !== null) window.clearTimeout(timerRef.current)
      // unmounting without Esc (host hides the panel) must not leave hits painted
      targetRef.current.clear()
    }
  }, [])

  // declared after the mount effect so opening straight into replace wins the focus
  // (replace falls back to find when read-only)
  useEffect(() => {
    if (!focusRequest?.nonce) return
    const el =
      (focusRequest.field === 'replace' ? replaceInputRef.current : null) ?? inputRef.current
    el?.focus()
    el?.select()
  }, [focusRequest])

  // stay in sync while the document changes underneath (typing, AI edits); the
  // query is read through a ref so a keystroke in the find box is never
  // overwritten by a rescan scheduled for the previous needle
  const queryRef = useRef(query)
  queryRef.current = query
  useEffect(
    () =>
      target.onDocChanged(() => {
        if (queryRef.current) scheduleRefresh(queryRef.current, 'current')
      }),
    [target, scheduleRefresh],
  )

  const close = useCallback(() => {
    if (timerRef.current !== null) window.clearTimeout(timerRef.current)
    timerRef.current = null
    target.clear()
    onClose()
  }, [target, onClose])

  const step = useCallback(
    (dir: 1 | -1) => {
      const fresh = flushPending(query)
      const n = fresh ? fresh.count : count
      if (n === 0) return
      // A flushed scan for a new query already landed on the first match — Enter should
      // visit it, not skip past it. A keep-current refresh (document changed underneath)
      // must still move in the requested direction from the refreshed position.
      const next = fresh?.queryChanged
        ? indexRef.current
        : ((fresh ? indexRef.current : index) + dir + n) % n
      setIndex(next)
      indexRef.current = next
      target.activate(next)
    },
    [flushPending, query, count, index, target],
  )

  const replaceOne = useCallback(() => {
    if (!canEdit) return
    // a pending debounced rescan means the target's list may describe the previous
    // query or pre-edit positions — replacing those would edit the wrong text
    const fresh = flushPending(query)
    const n = fresh ? fresh.count : count
    const at = fresh ? indexRef.current : index
    if (at >= n) return
    target.replaceOne(at, replacement)
    const after = refresh(query, at)
    if (after > 0) target.activate(Math.min(at, after - 1))
  }, [canEdit, flushPending, count, index, replacement, query, refresh, target])

  const replaceAll = useCallback(() => {
    if (!canEdit) return
    const fresh = flushPending(query)
    const n = fresh ? fresh.count : count
    if (n === 0) return
    target.replaceAll(replacement)
    refresh(query)
  }, [canEdit, flushPending, count, replacement, query, refresh, target])

  return (
    <div className="find-panel">
      <div className="find-row">
        <input
          ref={inputRef}
          className="find-input"
          placeholder={strings.findPlaceholder}
          value={query}
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
        <button
          type="button"
          className={`find-opt ${matchCase ? 'on' : ''}`}
          data-tip={strings.matchCase}
          aria-label={strings.matchCase}
          aria-pressed={matchCase}
          onClick={() => {
            setMatchCase(!matchCase)
            refresh(query, index, { matchCase: !matchCase })
          }}
        >
          Aa
        </button>
        <button
          type="button"
          className={`find-opt ${wholeWord ? 'on' : ''}`}
          data-tip={strings.wholeWord}
          aria-label={strings.wholeWord}
          aria-pressed={wholeWord}
          onClick={() => {
            setWholeWord(!wholeWord)
            refresh(query, index, { wholeWord: !wholeWord })
          }}
        >
          W
        </button>
        <span className="find-count">
          {query ? (count === 0 ? strings.noResults : `${index + 1}/${count}`) : ''}
        </span>
        <button
          type="button"
          className="find-btn"
          data-tip={strings.prevMatch}
          aria-label={strings.prevMatch}
          onClick={() => step(-1)}
          disabled={count === 0}
        >
          ‹
        </button>
        <button
          type="button"
          className="find-btn"
          data-tip={strings.nextMatch}
          aria-label={strings.nextMatch}
          onClick={() => step(1)}
          disabled={count === 0}
        >
          ›
        </button>
        <button
          type="button"
          className="find-btn find-close"
          data-tip={strings.closeEsc}
          aria-label={strings.closeEsc}
          onClick={close}
        >
          ✕
        </button>
      </div>
      {canEdit && (
        <div className="find-row">
          <input
            ref={replaceInputRef}
            className="find-input"
            placeholder={strings.replacePlaceholder}
            value={replacement}
            onChange={(e) => setReplacement(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') replaceOne()
              if (e.key === 'Escape') close()
            }}
          />
          <button type="button" className="find-action" onClick={replaceOne} disabled={count === 0}>
            {strings.replace}
          </button>
          <button type="button" className="find-action" onClick={replaceAll} disabled={count === 0}>
            {strings.replaceAll}
          </button>
        </div>
      )}
    </div>
  )
}
