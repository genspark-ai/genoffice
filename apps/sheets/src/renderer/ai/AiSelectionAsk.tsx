import React, { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { useI18n } from '../i18n/locale'
import {
  AiInlineLauncher,
  type AiInlineAction,
  type AiInlineLauncherStrings,
} from '@genoffice/ui'
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
  /** Optional one-shot translate trigger: receives the user-typed instruction
   * and returns the translated text (or null on failure). */
  onTranslate?: (instruction: string) => Promise<string | null>
}

const WIDTH = 340
const GAP = 8
const EDGE = 8
const EST_HEIGHT = 166
const MAX_INSTRUCTION = 2000

const INLINE_LAUNCHER_STRINGS: AiInlineLauncherStrings = {
  title: 'Ask AI about selection',
  polish: 'Polish',
  expand: 'Expand',
  shorten: 'Shorten',
  summarize: 'Summarize',
  translate: 'Translate',
}

/** Grid-selection Ask AI trigger and its send-now popover. */
export function AiSelectionAsk({
  anchor,
  range,
  onSend,
  onDismiss,
  onTranslate,
}: Props): React.JSX.Element {
  const { t } = useI18n()
  const [open, setOpen] = useState(false)
  const [text, setText] = useState('')
  const [triggerPlacement, setTriggerPlacement] = useState<{
    position: SelectionAskPosition
    width: number
    height: number
  } | null>(null)
  const [launcherRevision, setLauncherRevision] = useState(0)
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
    setLauncherRevision((revision) => revision + 1)
  }, [anchor])

  useEffect(() => {
    if (!open) return
    inputRef.current?.focus()
    const onPointerDown = (event: PointerEvent) => {
      if (!boxRef.current?.contains(event.target as Node)) onDismiss()
    }
    document.addEventListener('pointerdown', onPointerDown, true)
    return () => document.removeEventListener('pointerdown', onPointerDown, true)
  }, [onDismiss, open])

  const getInlineAnchorRect = (): {
    left: number
    top: number
    right: number
    bottom: number
    viewTop: number
    viewBottom: number
  } | null => {
    if (!anchor.bounds) return null
    return {
      left: anchor.bounds.left,
      top: anchor.bounds.top,
      right: anchor.bounds.right,
      bottom: anchor.bounds.bottom,
      viewTop: anchor.bounds.top,
      viewBottom: anchor.bounds.bottom,
    }
  }

  const handleInlinePick = (action: AiInlineAction) => {
    if (action === 'translate') {
      if (onTranslate) {
        const instruction = range
        void onTranslate(instruction)
      }
      return
    }
    let prompt = t('aiAnalyzePrompt')
    if (action === 'polish') prompt = t('aiCheckPrompt')
    setText(prompt)
    if (!open) setOpen(true)
    inputRef.current?.focus()
  }

  if (!open) {
    return (
      <>
        <AiInlineLauncher
          getAnchorRect={getInlineAnchorRect}
          strings={INLINE_LAUNCHER_STRINGS}
          onPick={handleInlinePick}
          revision={launcherRevision}
        />
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
      </>
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
  const handleTranslate = async () => {
    const instruction = text.trim() || range
    if (!onTranslate) return
    const translated = await onTranslate(instruction)
    if (translated !== null) {
      setText(translated)
      inputRef.current?.focus()
    }
  }
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
        <button
          className="ai-ask-chip"
          disabled={!onTranslate}
          onClick={handleTranslate}
        >
          {t('aiTranslateBtn')}
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
