import React, { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { useI18n } from '../i18n/locale'
import {
  selectionAskPosition,
  type SelectionAskAnchor,
  type SelectionAskPosition,
} from './selection-ask'

interface Props {
  anchor: SelectionAskAnchor
  range: string
  onSend: (instruction: string) => void
  onDismiss: () => void
}

const WIDTH = 340
const GAP = 8
const EDGE = 8
const EST_HEIGHT = 166
const MAX_INSTRUCTION = 2000

/** Grid-selection Ask AI trigger and its send-now popover. */
export function AiSelectionAsk({ anchor, range, onSend, onDismiss }: Props): React.JSX.Element {
  const { t } = useI18n()
  const [open, setOpen] = useState(false)
  const [text, setText] = useState('')
  const [triggerPlacement, setTriggerPlacement] = useState<{
    position: SelectionAskPosition
    width: number
    height: number
  } | null>(null)
  const triggerRef = useRef<HTMLButtonElement>(null)
  const boxRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  const askLabel = t('aiAskBtn')

  useLayoutEffect(() => {
    if (open || !triggerRef.current) return
    const width = triggerRef.current.offsetWidth
    const height = triggerRef.current.offsetHeight
    setTriggerPlacement({
      position: selectionAskPosition(anchor.pointer, anchor.bounds, width, height),
      width,
      height,
    })
  }, [anchor, askLabel, open])

  useEffect(() => {
    if (!open) return
    inputRef.current?.focus()
    const onPointerDown = (event: PointerEvent) => {
      if (!boxRef.current?.contains(event.target as Node)) onDismiss()
    }
    document.addEventListener('pointerdown', onPointerDown, true)
    return () => document.removeEventListener('pointerdown', onPointerDown, true)
  }, [onDismiss, open])

  if (!open) {
    return (
      <button
        ref={triggerRef}
        type="button"
        className="ai-ask-trigger"
        style={{
          left: triggerPlacement?.position.left ?? anchor.pointer.x,
          top: triggerPlacement?.position.top ?? anchor.pointer.y,
          visibility: triggerPlacement ? 'visible' : 'hidden',
        }}
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => setOpen(true)}
      >
        <svg viewBox="0 0 24 24" width="15" height="15" fill="none" aria-hidden>
          <path
            d="M12 3l1.7 4.6L18 9.3l-4.3 1.7L12 15.6l-1.7-4.6L6 9.3l4.3-1.7L12 3zM19 15l.85 2.3L22 18.15l-2.15.85L19 21.3l-.85-2.3-2.15-.85 2.15-.85L19 15z"
            fill="currentColor"
          />
        </svg>
        {askLabel}
      </button>
    )
  }

  const height = boxRef.current?.offsetHeight ?? EST_HEIGHT
  const triggerLeft = triggerPlacement?.position.left ?? anchor.pointer.x
  const triggerTop = triggerPlacement?.position.top ?? anchor.pointer.y
  const triggerWidth = triggerPlacement?.width ?? 0
  const triggerHeight = triggerPlacement?.height ?? 0
  const below = triggerTop + triggerHeight + GAP
  const above = triggerTop - GAP - height
  const top =
    below + height <= window.innerHeight - EDGE
      ? below
      : above >= EDGE
        ? above
        : Math.max(EDGE, window.innerHeight - EDGE - height)
  const left = Math.min(
    Math.max(EDGE, triggerLeft + triggerWidth / 2 - WIDTH / 2),
    Math.max(EDGE, window.innerWidth - WIDTH - EDGE),
  )
  const canSubmit = text.trim().length > 0
  const submit = () => {
    if (!canSubmit) return
    onSend(text.trim())
    onDismiss()
  }

  return (
    <div
      ref={boxRef}
      className="ai-ask-pop"
      style={{ left, top, width: WIDTH }}
      role="dialog"
      aria-label={askLabel}
      onKeyDown={(event) => {
        if (event.key === 'Escape') {
          event.preventDefault()
          onDismiss()
        }
      }}
    >
      <div className="ai-ask-pop-title">{askLabel}</div>
      <div className="ai-ask-pop-sub">{range}</div>
      <input
        ref={inputRef}
        className="ai-ask-pop-input"
        value={text}
        maxLength={MAX_INSTRUCTION}
        placeholder={t('aiComposerPlaceholder')}
        onChange={(event) => setText(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === 'Enter' && !event.nativeEvent.isComposing) {
            event.preventDefault()
            submit()
          }
        }}
      />
      <div className="ai-ask-pop-chips">
        <button
          className="ai-ask-chip"
          onClick={() => {
            setText(t('aiAnalyzePrompt'))
            inputRef.current?.focus()
          }}
        >
          {t('aiAnalyzeBtn')}
        </button>
        <button
          className="ai-ask-chip"
          onClick={() => {
            setText(t('aiCheckPrompt'))
            inputRef.current?.focus()
          }}
        >
          {t('aiCheckBtn')}
        </button>
      </div>
      <div className="ai-ask-pop-foot">
        <button className="ai-ask-cancel" onClick={onDismiss}>
          {t('aiCancel')}
        </button>
        <button className="ai-ask-confirm" disabled={!canSubmit} onClick={submit}>
          {t('aiSend')}
        </button>
      </div>
    </div>
  )
}
