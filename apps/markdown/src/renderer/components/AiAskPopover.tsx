/**
 * Floating "Ask AI" entry on a text selection, plus the anchored popover that
 * captures one instruction — sent immediately (a normal selection-scoped run)
 * or queued as an anchored edit. Clicking a queued anchor in the document
 * re-opens the same popover in edit mode. (Docs parity; markdown selection
 * kinds add code blocks and task lists.)
 *
 * Mounted at the app root with fixed positioning; ProseMirror's coordsAtPos
 * already speaks viewport coordinates. Positions refresh while the document
 * scrolls; outside clicks commit non-empty text into the queue (it stays
 * editable there) instead of discarding it.
 */
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import type { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import { useI18n, type StringKey } from '../i18n/locale'
import { AI_QUEUE_ANCHOR_CLICK, queueAnchorRange } from '../editor/aiQueueAnchors'
import {
  EDIT_INSTRUCTION_MAX,
  EDIT_QUEUE_MAX,
  truncate,
  type EditQueueItem,
} from '../ai/edit-queue'

interface AnchorRect {
  left: number
  top: number
  right: number
  bottom: number
}

interface Props {
  editor: Editor
  queueFull: boolean
  getItem: (qid: string) => EditQueueItem | undefined
  /** run the instruction right away as a normal selection-scoped request */
  onSendNow: (text: string) => void
  /** anchor the current selection and append to the queue */
  onQueueAdd: (text: string) => void
  onQueueUpdate: (qid: string, text: string) => void
  onQueueRemove: (qid: string) => void
}

type OpenState = { mode: 'new' } | { mode: 'edit'; qid: string }

const WIDTH = 340
const GAP = 8
const EDGE = 8
const EST_HEIGHT = 150

type SelectionKind = 'image' | 'table' | 'code' | 'task' | 'text'

/** Suggestions match what the AI can actually do to the selected element */
const CHIP_KEYS: Record<SelectionKind, StringKey[]> = {
  image: ['aiChipReplaceImage', 'aiChipRegenImage', 'aiChipImageCaption'],
  table: ['aiChipTableEdit', 'aiChipPolish', 'aiChipFixGrammar'],
  code: ['aiChipExplainCode', 'aiChipFixCode', 'aiChipCommentCode'],
  task: ['aiChipSplitTasks', 'aiChipPrioritizeTasks', 'aiChipPolish'],
  text: ['aiChipPolish', 'aiChipShorten', 'aiChipExpand', 'aiChipFixGrammar'],
}

function selectionKind(editor: Editor): SelectionKind {
  const sel = editor.state.selection
  if (sel instanceof CellSelection) return 'table'
  if (sel instanceof NodeSelection) {
    if (sel.node.type.name === 'image') return 'image'
    if (sel.node.type.name === 'table') return 'table'
  }
  const { $from } = sel
  for (let depth = $from.depth; depth > 0; depth--) {
    const name = $from.node(depth).type.name
    if (name === 'codeBlock') return 'code'
    if (name === 'taskItem') return 'task'
    if (name === 'table') return 'table'
  }
  return 'text'
}

/** viewport rect of a PM position range */
function rangeRect(editor: Editor, from: number, to: number): AnchorRect | null {
  try {
    const a = editor.view.coordsAtPos(from)
    const b = editor.view.coordsAtPos(to)
    return {
      left: Math.min(a.left, b.left),
      top: Math.min(a.top, b.top),
      right: Math.max(a.right, b.right),
      bottom: Math.max(a.bottom, b.bottom),
    }
  } catch {
    return null // positions can be transiently unmappable mid-transaction
  }
}

export function AiAskPopover({
  editor,
  queueFull,
  getItem,
  onSendNow,
  onQueueAdd,
  onQueueUpdate,
  onQueueRemove,
}: Props): ReactElement | null {
  const { t } = useI18n()
  const [trigger, setTrigger] = useState<{ left: number; top: number } | null>(null)
  const [open, setOpen] = useState<OpenState | null>(null)
  const [kind, setKind] = useState<SelectionKind>('text')
  const [text, setText] = useState('')
  const [rect, setRect] = useState<AnchorRect | null>(null)
  const boxRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  const openRef = useRef(open)
  openRef.current = open
  const textRef = useRef(text)
  textRef.current = text
  const queueFullRef = useRef(queueFull)
  queueFullRef.current = queueFull

  const anchorRect = useCallback((): AnchorRect | null => {
    const state = openRef.current
    if (!state) return null
    if (state.mode === 'edit') {
      const range = queueAnchorRange(editor.state, state.qid)
      return range ? rangeRect(editor, range.from, range.to) : null
    }
    const { from, to, empty } = editor.state.selection
    return empty ? null : rangeRect(editor, from, to)
  }, [editor])

  // floating trigger follows the live selection while the editor is focused
  useEffect(() => {
    const update = () => {
      if (openRef.current) {
        setTrigger(null)
        return
      }
      const { empty, to } = editor.state.selection
      if (empty || !editor.isFocused) {
        setTrigger(null)
        return
      }
      const c = rangeRect(editor, to, to)
      if (!c) {
        setTrigger(null)
        return
      }
      setTrigger({ left: c.right + 6, top: c.bottom + 4 })
    }
    const hide = () => setTrigger(null)
    editor.on('selectionUpdate', update)
    editor.on('focus', update)
    editor.on('blur', hide)
    const scroller = document.querySelector('.editor-scroll')
    scroller?.addEventListener('scroll', update, { passive: true })
    window.addEventListener('resize', update)
    return () => {
      editor.off('selectionUpdate', update)
      editor.off('focus', update)
      editor.off('blur', hide)
      scroller?.removeEventListener('scroll', update)
      window.removeEventListener('resize', update)
    }
  }, [editor])

  // clicking a queued anchor in the document re-opens the popover on that item
  useEffect(() => {
    const onAnchorClick = (e: Event) => {
      const qid = (e as CustomEvent<{ qid: string }>).detail?.qid
      const item = qid ? getItem(qid) : undefined
      if (!item) return
      setText(item.instruction)
      setOpen({ mode: 'edit', qid: item.qid })
      setTrigger(null)
    }
    window.addEventListener(AI_QUEUE_ANCHOR_CLICK, onAnchorClick)
    return () => window.removeEventListener(AI_QUEUE_ANCHOR_CLICK, onAnchorClick)
  }, [getItem])

  // popover position follows scroll/resize; a vanished anchor closes it
  useLayoutEffect(() => {
    if (!open) return
    const update = () => setRect(anchorRect())
    update()
    const scroller = document.querySelector('.editor-scroll')
    scroller?.addEventListener('scroll', update, { passive: true })
    window.addEventListener('resize', update)
    return () => {
      scroller?.removeEventListener('scroll', update)
      window.removeEventListener('resize', update)
    }
  }, [open, anchorRect])

  useEffect(() => {
    if (open) {
      inputRef.current?.focus()
      inputRef.current?.select()
    }
  }, [open])

  const close = useCallback(() => {
    setOpen(null)
    setText('')
  }, [])

  /** outside click keeps whatever was typed: a new note lands in the queue, an edit is saved */
  const commitOrClose = useCallback(() => {
    const state = openRef.current
    const value = textRef.current.trim()
    if (state && value) {
      if (state.mode === 'edit') onQueueUpdate(state.qid, value)
      else if (!queueFullRef.current) onQueueAdd(value)
    }
    close()
  }, [onQueueAdd, onQueueUpdate, close])

  useEffect(() => {
    if (!open) return
    const onPointerDown = (e: PointerEvent) => {
      if (boxRef.current?.contains(e.target as Node)) return
      commitOrClose()
    }
    document.addEventListener('pointerdown', onPointerDown, true)
    return () => document.removeEventListener('pointerdown', onPointerDown, true)
  }, [open, commitOrClose])

  const openFromSelection = () => {
    setText('')
    setKind(selectionKind(editor))
    setOpen({ mode: 'new' })
    setTrigger(null)
  }

  const canSubmit = text.trim().length > 0
  const editItem = open?.mode === 'edit' ? getItem(open.qid) : undefined

  const excerpt = (() => {
    if (!open) return ''
    if (open.mode === 'edit') {
      const range = queueAnchorRange(editor.state, open.qid)
      return range
        ? editor.state.doc.textBetween(range.from, range.to, ' ', ' ')
        : (editItem?.capturedText ?? '')
    }
    const { from, to } = editor.state.selection
    return editor.state.doc.textBetween(from, to, ' ', ' ')
  })()
    .replace(/\s+/g, ' ')
    .trim()

  let box: ReactElement | null = null
  if (open && rect) {
    const height = boxRef.current?.offsetHeight ?? EST_HEIGHT
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
    box = (
      <div
        ref={boxRef}
        className="ai-ask-pop"
        style={{ left, top, width: WIDTH }}
        role="dialog"
        onKeyDown={(e) => {
          if (e.key === 'Escape') {
            e.preventDefault()
            e.stopPropagation()
            close()
          }
        }}
      >
        <button
          type="button"
          className="ai-ask-pop-close"
          data-tip={t('aiCancel')}
          aria-label={t('aiCancel')}
          onClick={close}
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
        <div className="ai-ask-pop-title">
          {t(open.mode === 'edit' ? 'aiAskEditTitle' : 'aiAskTitle')}
        </div>
        {excerpt && <div className="ai-ask-pop-sub">{truncate(excerpt, 60)}</div>}
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
              if (open.mode === 'edit') {
                onQueueUpdate(open.qid, text.trim())
                close()
              } else if (!queueFull) {
                onQueueAdd(text.trim())
                close()
              }
            }
          }}
        />
        {open.mode === 'new' && (
          <div className="ai-ask-pop-chips">
            {CHIP_KEYS[kind].map((key) => (
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
          {open.mode === 'edit' ? (
            <>
              <button
                className="ai-ask-cancel"
                onClick={() => {
                  onQueueRemove(open.qid)
                  close()
                }}
              >
                {t('blockDelete')}
              </button>
              <button
                className="ai-ask-confirm"
                disabled={!canSubmit}
                onClick={() => {
                  onQueueUpdate(open.qid, text.trim())
                  close()
                }}
              >
                {t('aiAskUpdate')}
              </button>
            </>
          ) : (
            <>
              <button
                className="ai-ask-cancel"
                disabled={!canSubmit}
                onClick={() => {
                  onSendNow(text.trim())
                  close()
                }}
              >
                {t('aiAskSendNow')}
              </button>
              <button
                className="ai-ask-confirm"
                disabled={!canSubmit || queueFull}
                data-tip={queueFull ? t('aiQueueFullNotice', { max: EDIT_QUEUE_MAX }) : undefined}
                onClick={() => {
                  onQueueAdd(text.trim())
                  close()
                }}
              >
                {t('aiAskQueue')}
              </button>
            </>
          )}
        </div>
      </div>
    )
  }

  return (
    <>
      {trigger && !open && (
        <button
          className="ai-ask-trigger"
          style={{ left: trigger.left, top: trigger.top }}
          // keep the editor focused and the selection alive
          onMouseDown={(e) => e.preventDefault()}
          onClick={openFromSelection}
        >
          <svg viewBox="0 0 24 24" width="15" height="15" fill="none" aria-hidden>
            <path
              d="M12 3l1.7 4.6L18 9.3l-4.3 1.7L12 15.6l-1.7-4.6L6 9.3l4.3-1.7L12 3zM19 15l.85 2.3L22 18.15l-2.15.85L19 21.3l-.85-2.3-2.15-.85 2.15-.85L19 15z"
              fill="currentColor"
            />
          </svg>
          {t('aiAskBtn')}
        </button>
      )}
      {box}
    </>
  )
}
