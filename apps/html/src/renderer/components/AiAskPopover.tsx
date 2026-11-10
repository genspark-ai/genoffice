/**
 * Anchored input for annotating the selected preview element with one AI edit
 * (slides/docs parity). Mounted at the app root in viewport coordinates: the
 * preview stage zooms its host, so the popover must not live inside it.
 */
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import { useI18n, type StringKey } from '../i18n/locale'
import { EDIT_INSTRUCTION_MAX, EDIT_QUEUE_MAX, truncate } from '../ai/edit-queue'

/** viewport rect of the anchor (the selected element) */
export interface AnchorRect {
  left: number
  top: number
  right: number
  bottom: number
  /** band the popover must stay inside — the preview stage, so it never covers the ribbon */
  viewTop: number
  viewBottom: number
}

export interface AskTarget {
  tag: string
  excerpt: string
}

export type AskMode = { kind: 'new' } | { kind: 'edit'; qid: string }

interface Props {
  target: AskTarget
  mode: AskMode
  initialText?: string
  /** null hides the popover (element gone or scrolled out) */
  getAnchorRect: () => AnchorRect | null
  /** new: add to queue; edit: update the queued instruction */
  onSubmit: (text: string) => void
  onCancel: () => void
  onSendNow: (text: string) => void
  onRemove: () => void
  /** disables "Add to queue" only; sending now never touches the queue */
  queueFull: boolean
}

const WIDTH = 340
const GAP = 10
const EDGE = 8
/** height assumed before the first measurement, used only to pick a side */
const EST_HEIGHT = 150

const TEXT_TAGS = new Set([
  'p',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'li',
  'td',
  'th',
  'blockquote',
  'span',
  'a',
  'strong',
  'em',
  'label',
  'figcaption',
  'caption',
  'dd',
  'dt',
  'summary',
  'button',
])

function chipKeys(target: AskTarget): StringKey[] {
  if (target.tag === 'img' || target.tag === 'picture' || target.tag === 'svg')
    return ['aiChipReplaceImage', 'aiChipRecolor']
  if (target.tag === 'table') return ['aiChipTableEdit', 'aiChipRecolor']
  if (TEXT_TAGS.has(target.tag) && !target.excerpt.startsWith('<'))
    return ['aiChipPolish', 'aiChipShorten', 'aiChipExpand', 'aiChipFixGrammar']
  return ['aiChipUnify', 'aiChipRecolor']
}

export function AiAskPopover({
  target,
  mode,
  initialText,
  getAnchorRect,
  onSubmit,
  onCancel,
  onSendNow,
  onRemove,
  queueFull,
}: Props): ReactElement | null {
  const { t } = useI18n()
  const [text, setText] = useState(initialText ?? '')
  const [rect, setRect] = useState<AnchorRect | null>(() => getAnchorRect())
  const boxRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  const textRef = useRef(text)
  textRef.current = text

  // the anchor can arrive after mount (a pin click re-selects first, so the rect is briefly null): focus once the input exists
  const focusedRef = useRef(false)
  useEffect(() => {
    if (focusedRef.current || !inputRef.current) return
    focusedRef.current = true
    inputRef.current.focus()
    inputRef.current.select()
  })

  useLayoutEffect(() => {
    const update = () => setRect(getAnchorRect())
    update()
    window.addEventListener('resize', update)
    return () => window.removeEventListener('resize', update)
  }, [getAnchorRect])

  // a full queue blocks queueing only for new asks; edits update in place
  const queueBlocked = queueFull && mode.kind === 'new'

  /** clicking away keeps whatever was typed: it lands in the queue, where it can still be edited or removed */
  const commitOrCancel = useCallback(() => {
    const value = textRef.current.trim()
    if (value && !queueBlocked) onSubmit(value)
    else onCancel()
  }, [onSubmit, onCancel, queueBlocked])

  useEffect(() => {
    const onPointerDown = (e: PointerEvent) => {
      if (boxRef.current?.contains(e.target as Node)) return
      commitOrCancel()
    }
    // capture phase: the preview iframe swallows pointer events, but a click on it still blurs the window
    document.addEventListener('pointerdown', onPointerDown, true)
    window.addEventListener('blur', commitOrCancel)
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true)
      window.removeEventListener('blur', commitOrCancel)
    }
  }, [commitOrCancel])

  if (!rect) return null

  const height = boxRef.current?.offsetHeight ?? EST_HEIGHT
  const bandTop = Math.max(rect.viewTop, EDGE)
  const bandBottom = Math.min(rect.viewBottom, window.innerHeight - EDGE)
  const below = rect.bottom + GAP
  const above = rect.top - GAP - height
  const top =
    below + height <= bandBottom
      ? below
      : above >= bandTop
        ? above
        : Math.min(Math.max(rect.top + GAP, bandTop), Math.max(bandTop, bandBottom - height))
  const left = Math.min(
    Math.max(EDGE, (rect.left + rect.right) / 2 - WIDTH / 2),
    Math.max(EDGE, window.innerWidth - WIDTH - EDGE),
  )
  const canSubmit = text.trim().length > 0

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
          onCancel()
        } else if (e.key === 'Tab') {
          const focusable = boxRef.current?.querySelectorAll<HTMLElement>('input, button')
          if (!focusable || focusable.length === 0) return
          const first = focusable[0]!
          const last = focusable[focusable.length - 1]!
          if (e.shiftKey && document.activeElement === first) {
            e.preventDefault()
            last.focus()
          } else if (!e.shiftKey && document.activeElement === last) {
            e.preventDefault()
            first.focus()
          }
        }
      }}
    >
      <button className="ai-ask-pop-close" aria-label={t('aiAskCancel')} onClick={onCancel}>
        <svg viewBox="0 0 24 24" width="12" height="12" fill="none" aria-hidden>
          <path
            d="M6 6l12 12M18 6L6 18"
            stroke="currentColor"
            strokeWidth="2"
            strokeLinecap="round"
          />
        </svg>
      </button>
      <div className="ai-ask-pop-title">
        {t(mode.kind === 'edit' ? 'aiAskEditTitle' : 'aiAskTitle')}
      </div>
      <div className="ai-ask-pop-sub">
        {`<${target.tag}> `}
        {truncate(target.excerpt, 60)}
      </div>
      <input
        ref={inputRef}
        className="ai-ask-pop-input"
        value={text}
        maxLength={EDIT_INSTRUCTION_MAX}
        placeholder={t('aiAskPlaceholder')}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter' && !e.nativeEvent.isComposing) {
            e.preventDefault()
            if (!canSubmit) return
            if (queueBlocked) onSendNow(text.trim())
            else onSubmit(text.trim())
          }
        }}
      />
      {mode.kind === 'new' && (
        <div className="ai-ask-pop-chips">
          {chipKeys(target).map((key) => (
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
      )}
      <div className="ai-ask-pop-foot">
        {mode.kind === 'edit' ? (
          <>
            <button className="ai-ask-cancel" onClick={onRemove}>
              {t('elementDelete')}
            </button>
            <button
              className="ai-ask-confirm"
              disabled={!canSubmit}
              onClick={() => canSubmit && onSubmit(text.trim())}
            >
              {t('aiAskUpdate')}
            </button>
          </>
        ) : (
          <>
            <button
              className="ai-ask-cancel"
              disabled={!canSubmit}
              onClick={() => canSubmit && onSendNow(text.trim())}
            >
              {t('aiAskSendNow')}
            </button>
            <button
              className="ai-ask-confirm"
              disabled={!canSubmit || queueFull}
              data-tip={queueFull ? t('aiQueueFullNotice', { max: EDIT_QUEUE_MAX }) : undefined}
              onClick={() => canSubmit && !queueFull && onSubmit(text.trim())}
            >
              {t('aiAskQueue')}
            </button>
          </>
        )}
      </div>
    </div>
  )
}
