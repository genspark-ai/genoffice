/**
 * Anchored input for annotating the selected element(s) with one AI edit.
 *
 * Mounted at the app root, never inside `.stage-scale`: that subtree carries
 * `transform: scale(zoom)`, which would scale the popover with the canvas
 * (illegible at 50%, oversized at 200%). Position is measured in viewport
 * coordinates and refreshed while the canvas scrolls or zooms.
 */
import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import { useI18n, type StringKey } from '../i18n/locale'
import {
  EDIT_INSTRUCTION_MAX,
  NODE_NOUN_KEY,
  truncate,
  type NodeDescriptor,
} from '../ai/edit-queue'

/** Viewport rect of the anchor (the selection bounding box) */
export interface AnchorRect {
  left: number
  top: number
  right: number
  bottom: number
  /** Band the popover must stay inside — the canvas viewport, so it never covers the ribbon */
  viewTop: number
  viewBottom: number
}

export interface AskTarget {
  /** Anchor id handed to the model (durable when the element has one) */
  id: string
  sourceId: string
  desc: NodeDescriptor
}

interface Props {
  targets: AskTarget[]
  /** Pre-filled when re-opening on an element that already carries a queued edit */
  initialText?: string
  /** Re-measured on scroll/zoom; null hides the popover (element scrolled out of the deck) */
  getAnchorRect: () => AnchorRect | null
  onSubmit: (text: string) => void
  onCancel: () => void
}

const WIDTH = 340
/** Gap between the anchor box and the popover */
const GAP = 10
/** Keep-away margin from the viewport edges */
const EDGE = 8
/** Height assumed before the first measurement, used only to pick a side */
const EST_HEIGHT = 150

/** At most four suggestions, keyed by what is selected; each fills the input rather than submitting */
function chipKeys(targets: AskTarget[]): StringKey[] {
  if (targets.length !== 1) return ['aiChipUnify', 'aiChipRecolor']
  const only = targets[0]!.desc
  switch (only.type) {
    case 'picture':
      return ['aiChipReplaceImage', 'aiChipRemoveBg', 'aiChipRecolor']
    case 'table':
      return ['aiChipTableEdit', 'aiChipRecolor', 'aiChipPolish']
    case 'chart':
      return ['aiChipChartType', 'aiChipRecolor']
    case 'group':
      return ['aiChipUnify', 'aiChipRecolor']
    default:
      return only.text
        ? ['aiChipPolish', 'aiChipShorten', 'aiChipExpand', 'aiChipVerify']
        : ['aiChipRecolor']
  }
}

export function AiAskPopover({
  targets,
  initialText,
  getAnchorRect,
  onSubmit,
  onCancel,
}: Props): React.JSX.Element | null {
  const { t } = useI18n()
  const [text, setText] = useState(initialText ?? '')
  const [rect, setRect] = useState<AnchorRect | null>(() => getAnchorRect())
  const boxRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  // Latest text without re-subscribing the outside-click listener on every keystroke
  const textRef = useRef(text)
  textRef.current = text

  useEffect(() => {
    inputRef.current?.focus()
    inputRef.current?.select()
  }, [])

  useLayoutEffect(() => {
    const update = () => setRect(getAnchorRect())
    update()
    const scroller = document.querySelector('.stage-wrap')
    scroller?.addEventListener('scroll', update, { passive: true })
    window.addEventListener('resize', update)
    return () => {
      scroller?.removeEventListener('scroll', update)
      window.removeEventListener('resize', update)
    }
  }, [getAnchorRect])

  /** Clicking away keeps whatever was typed: it lands in the queue, where it can still be edited or removed */
  const commitOrCancel = useCallback(() => {
    const value = textRef.current.trim()
    if (value) onSubmit(value)
    else onCancel()
  }, [onSubmit, onCancel])

  useEffect(() => {
    const onPointerDown = (e: PointerEvent) => {
      if (boxRef.current?.contains(e.target as Node)) return
      commitOrCancel()
    }
    // Capture phase: the canvas swallows pointer events on the Konva stage
    document.addEventListener('pointerdown', onPointerDown, true)
    return () => document.removeEventListener('pointerdown', onPointerDown, true)
  }, [commitOrCancel])

  if (!rect) return null

  const height = boxRef.current?.offsetHeight ?? EST_HEIGHT
  const bandTop = Math.max(rect.viewTop, EDGE)
  const bandBottom = Math.min(rect.viewBottom, window.innerHeight - EDGE)
  const below = rect.bottom + GAP
  const above = rect.top - GAP - height
  // Below, else above, else tucked just inside the anchor's top edge — the last
  // case covers elements taller than the free space (full-bleed pictures)
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

  const noun = t(NODE_NOUN_KEY[targets[0]?.desc.type ?? 'shape'])
  const preview = targets.length === 1 ? targets[0]!.desc.text : undefined
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
          // Focus trap: the canvas behind must never take focus while this is open
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
      <div className="ai-ask-pop-title">
        {targets.length === 1
          ? t('aiAskTitle', { label: noun })
          : t('aiAskTitleMulti', { count: targets.length })}
      </div>
      {preview && <div className="ai-ask-pop-sub">{truncate(preview, 60)}</div>}
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
            if (canSubmit) onSubmit(text.trim())
          }
        }}
      />
      <div className="ai-ask-pop-chips">
        {chipKeys(targets).map((key) => (
          <button
            key={key}
            className="ai-ask-chip"
            // Filling rather than submitting: the shorthand almost always wants a qualifier
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
        <button className="ai-ask-cancel" onClick={onCancel}>
          {t('paneCancel')}
        </button>
        <button
          className="ai-ask-confirm"
          disabled={!canSubmit}
          onClick={() => canSubmit && onSubmit(text.trim())}
        >
          {t('aiAskConfirm')}
        </button>
      </div>
    </div>
  )
}
