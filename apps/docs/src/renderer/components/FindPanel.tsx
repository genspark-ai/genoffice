import { useCallback, useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n } from '../i18n/locale'
import { searchPluginKey } from '../editor/extensions'

interface Range {
  from: number
  to: number
}

interface FindOptions {
  matchCase: boolean
  wholeWord: boolean
}

const isWordChar = (ch: string | undefined) => !!ch && /[\p{L}\p{N}_]/u.test(ch)

/** length-preserving lowercase: chars whose lowercase grows ('İ' → 'i̇') stay as-is so match offsets never shift */
export function foldCase(s: string): string {
  let out = ''
  for (const ch of s) {
    const lower = ch.toLowerCase()
    out += lower.length === ch.length ? lower : ch
  }
  return out
}

/** collect matches inside editable textblocks (protected blocks excluded) */
export function findMatches(editor: Editor, query: string, opts: FindOptions): Range[] {
  const found: Range[] = []
  if (!query) return found
  const needle = opts.matchCase ? query : foldCase(query)
  editor.state.doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    // flatten the block's inline content so matches spanning marks are found
    let text = ''
    const posAt: number[] = []
    node.forEach((child, offset) => {
      if (child.isText && child.text) {
        for (let k = 0; k < child.text.length; k++) posAt.push(pos + 1 + offset + k)
        text += child.text
      } else {
        posAt.push(pos + 1 + offset)
        text += '\u0000' // leaf placeholder (hard break) never matches
      }
    })
    const haystack = opts.matchCase ? text : foldCase(text)
    let i = 0
    while ((i = haystack.indexOf(needle, i)) !== -1) {
      const isWhole =
        !opts.wholeWord || (!isWordChar(text[i - 1]) && !isWordChar(text[i + query.length]))
      if (isWhole) {
        found.push({ from: posAt[i], to: posAt[i + query.length - 1] + 1 })
        i += query.length
      } else {
        i += 1
      }
    }
    return false
  })
  return found
}

interface FindPanelProps {
  editor: Editor
  onClose: () => void
}

const SCAN_DEBOUNCE_MS = 150

export function FindPanel({ editor, onClose }: FindPanelProps) {
  const { t } = useI18n()
  const [query, setQuery] = useState('')
  const [replacement, setReplacement] = useState('')
  const [matches, setMatches] = useState<Range[]>([])
  const [index, setIndex] = useState(0)
  const [matchCase, setMatchCase] = useState(false)
  const [wholeWord, setWholeWord] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)
  const indexRef = useRef(0)
  const canEdit = editor.isEditable

  const highlight = useCallback(
    (ranges: Range[], activeIndex: number) => {
      editor.view.dispatch(editor.state.tr.setMeta(searchPluginKey, { ranges, activeIndex }))
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

  // rescan only updates matches/highlight; scrolling happens on explicit navigation
  const refresh = useCallback(
    (q: string, keepIndex = 0, opts?: Partial<FindOptions>) => {
      const ranges = findMatches(editor, q, { matchCase, wholeWord, ...opts })
      const active = ranges.length === 0 ? 0 : Math.min(keepIndex, ranges.length - 1)
      setMatches(ranges)
      setIndex(active)
      indexRef.current = active
      highlight(ranges, active)
      return ranges
    },
    [editor, highlight, matchCase, wholeWord],
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
        ranges: refresh(q, keep === 'current' ? indexRef.current : 0),
        queryChanged: keep === 'reset',
      }
    },
    [refresh],
  )

  useEffect(() => {
    inputRef.current?.focus()
    inputRef.current?.select()
    return () => {
      if (timerRef.current !== null) window.clearTimeout(timerRef.current)
    }
  }, [])

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
    highlight([], 0)
    onClose()
  }, [highlight, onClose])

  const step = useCallback(
    (dir: 1 | -1) => {
      const fresh = flushPending(query)
      const ranges = fresh ? fresh.ranges : matches
      if (ranges.length === 0) return
      // A flushed scan for a new query already landed on the first match — Enter should
      // visit it, not skip past it. A keep-current refresh (document changed underneath)
      // must still move in the requested direction from the refreshed position.
      const next = fresh?.queryChanged
        ? indexRef.current
        : ((fresh ? indexRef.current : index) + dir + ranges.length) % ranges.length
      setIndex(next)
      indexRef.current = next
      highlight(ranges, next)
      scrollTo(ranges[next])
    },
    [flushPending, query, matches, index, highlight, scrollTo],
  )

  const replaceOne = useCallback(() => {
    if (!editor.isEditable) return
    // a pending debounced rescan means `matches` may describe the previous
    // query or pre-edit positions — replacing those would edit the wrong text
    const fresh = flushPending(query)
    const ranges = fresh ? fresh.ranges : matches
    const at = fresh ? indexRef.current : index
    const m = ranges[at]
    if (!m) return
    editor.commands.command(({ tr }) => {
      tr.insertText(replacement, m.from, m.to)
      return true
    })
    const after = refresh(query, at)
    if (after.length > 0) scrollTo(after[Math.min(at, after.length - 1)])
  }, [editor, flushPending, matches, index, replacement, query, refresh, scrollTo])

  const replaceAll = useCallback(() => {
    if (!editor.isEditable) return
    const fresh = flushPending(query)
    const ranges = fresh ? fresh.ranges : matches
    if (ranges.length === 0) return
    editor.commands.command(({ tr }) => {
      for (const m of [...ranges].reverse()) tr.insertText(replacement, m.from, m.to)
      return true
    })
    refresh(query)
  }, [editor, flushPending, matches, replacement, query, refresh])

  return (
    <div className="find-panel">
      <div className="find-row">
        <input
          ref={inputRef}
          className="find-input"
          placeholder={t('appFindPlaceholder')}
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
          className={`find-opt ${matchCase ? 'on' : ''}`}
          data-tip={t('appMatchCase')}
          onClick={() => {
            setMatchCase(!matchCase)
            refresh(query, index, { matchCase: !matchCase })
          }}
        >
          Aa
        </button>
        <button
          className={`find-opt ${wholeWord ? 'on' : ''}`}
          data-tip={t('appWholeWord')}
          onClick={() => {
            setWholeWord(!wholeWord)
            refresh(query, index, { wholeWord: !wholeWord })
          }}
        >
          W
        </button>
        <span className="find-count">
          {query
            ? matches.length === 0
              ? t('appNoResults')
              : `${index + 1}/${matches.length}`
            : ''}
        </span>
        <button
          className="find-btn"
          data-tip={t('appPrevMatch')}
          aria-label={t('appPrevMatch')}
          onClick={() => step(-1)}
          disabled={matches.length === 0}
        >
          ‹
        </button>
        <button
          className="find-btn"
          data-tip={t('appNextMatch')}
          aria-label={t('appNextMatch')}
          onClick={() => step(1)}
          disabled={matches.length === 0}
        >
          ›
        </button>
        <button
          className="find-btn find-close"
          data-tip={t('appCloseEsc')}
          aria-label={t('appCloseEsc')}
          onClick={close}
        >
          ✕
        </button>
      </div>
      {canEdit && (
        <div className="find-row">
          <input
            className="find-input"
            placeholder={t('appReplacePlaceholder')}
            value={replacement}
            onChange={(e) => setReplacement(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') replaceOne()
              if (e.key === 'Escape') close()
            }}
          />
          <button className="find-action" onClick={replaceOne} disabled={matches.length === 0}>
            {t('appReplace')}
          </button>
          <button className="find-action" onClick={replaceAll} disabled={matches.length === 0}>
            {t('appReplaceAll')}
          </button>
        </div>
      )}
    </div>
  )
}
