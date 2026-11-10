/**
 * Anchored "Ask AI" popover opened from the selection markup bar: captures one
 * instruction about the selected passage and sends it immediately as a normal
 * selection-scoped run (pdf deliberately has no edit queue — pending notes are
 * the durable way to mark work for later).
 *
 * Mounted at the app root with fixed positioning; the anchor rect is captured
 * when the popover opens (the native selection collapses once the input takes
 * focus, so nothing here re-reads the live selection).
 */
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import { useI18n, type StringKey } from './i18n/locale'

export interface AskAnchorRect {
  left: number
  top: number
  right: number
  bottom: number
}

const WIDTH = 340
const GAP = 8
const EDGE = 8
const EST_HEIGHT = 150
const INSTRUCTION_MAX = 500

const READ_CHIPS: StringKey[] = ['aiChipExplain', 'aiChipTranslate', 'aiChipSummarizePassage']

export function AiAskPopover({
  rect,
  excerpt,
  readOnly,
  onSend,
  onClose,
}: {
  rect: AskAnchorRect
  excerpt: string
  /** encrypted documents hide the edit-oriented chip */
  readOnly: boolean
  onSend: (text: string) => void
  onClose: () => void
}): ReactElement {
  const { t } = useI18n()
  const [text, setText] = useState('')
  const boxRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  // real height measured before first paint — placing by the estimate and
  // re-reading offsetHeight on later renders would make the box jump/flip
  const [measuredH, setMeasuredH] = useState<number | null>(null)
  useLayoutEffect(() => {
    setMeasuredH(boxRef.current?.offsetHeight ?? null)
  }, [])

  useEffect(() => {
    inputRef.current?.focus()
  }, [])

  const onCloseRef = useRef(onClose)
  onCloseRef.current = onClose
  useEffect(() => {
    const onPointerDown = (e: PointerEvent) => {
      if (boxRef.current?.contains(e.target as Node)) return
      onCloseRef.current()
    }
    document.addEventListener('pointerdown', onPointerDown, true)
    return () => document.removeEventListener('pointerdown', onPointerDown, true)
  }, [])

  const canSubmit = text.trim().length > 0
  const submit = useCallback(() => {
    const value = text.trim()
    if (value) onSend(value)
  }, [text, onSend])

  const height = measuredH ?? EST_HEIGHT
  const below = rect.bottom + GAP
  const above = rect.top - GAP - height
  const top =
    below + height <= window.innerHeight - EDGE
      ? below
      : above >= EDGE
        ? above
        : Math.max(EDGE, window.innerHeight - EDGE - height)
  const left = Math.min(
    Math.max(EDGE, (rect.left + rect.right) / 2 - WIDTH / 2),
    Math.max(EDGE, window.innerWidth - WIDTH - EDGE),
  )
  const chips = readOnly ? READ_CHIPS : [...READ_CHIPS, 'aiChipRewritePassage' as StringKey]
  const shortExcerpt = excerpt.replace(/\s+/g, ' ').trim()

  return (
    <div
      ref={boxRef}
      className="ai-ask-pop"
      style={{ left, top, width: WIDTH }}
      role="dialog"
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.preventDefault()
          e.stopPropagation()
          onClose()
        }
      }}
    >
      <button
        type="button"
        className="ai-ask-pop-close"
        data-tip={t('noteClose')}
        aria-label={t('noteClose')}
        onClick={onClose}
      >
        <svg viewBox="0 0 24 24" width="12" height="12" fill="none" aria-hidden>
          <path
            d="M6 6l12 12M18 6L6 18"
            stroke="currentColor"
            strokeWidth="2"
            strokeLinecap="round"
          />
        </svg>
      </button>
      <div className="ai-ask-pop-title">{t('aiAskTitle')}</div>
      {shortExcerpt && (
        <div className="ai-ask-pop-sub">
          {shortExcerpt.length > 60 ? `${shortExcerpt.slice(0, 60)}…` : shortExcerpt}
        </div>
      )}
      <input
        ref={inputRef}
        className="ai-ask-pop-input"
        value={text}
        maxLength={INSTRUCTION_MAX}
        placeholder={t('aiAskPlaceholder')}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter' && !e.nativeEvent.isComposing) {
            e.preventDefault()
            submit()
          }
        }}
      />
      <div className="ai-ask-pop-chips">
        {chips.map((key) => (
          <button
            key={key}
            className="ai-ask-chip"
            // filling rather than submitting: the shorthand almost always wants a qualifier
            onClick={() => {
              setText(t(key))
              inputRef.current?.focus()
            }}
          >
            {t(key)}
          </button>
        ))}
      </div>
      <div className="ai-ask-pop-foot">
        <button className="ai-ask-confirm" disabled={!canSubmit} onClick={submit}>
          {t('aiSend')}
        </button>
      </div>
    </div>
  )
}
