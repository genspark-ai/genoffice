/**
 * Pending selection-scoped edits, shown above the composer. Rows resolve
 * their anchors against the live document on every render, so the excerpt
 * follows the current content instead of a snapshot taken at annotation time.
 */
import React, { useState } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n } from '../i18n/locale'
import {
  EDIT_INSTRUCTION_MAX,
  EDIT_QUEUE_MAX,
  resolveQueueItem,
  truncate,
  type DocsEditQueueItem,
} from './edit-queue'

interface Props {
  items: DocsEditQueueItem[]
  editor: Editor
  busy: boolean
  onEditInstruction: (qid: string, instruction: string) => void
  onRemove: (qid: string) => void
  onDiscardAll: () => void
  onSend: () => void
  /** scroll to the anchored passage and select it */
  onFocus: (qid: string) => void
}

/** beyond this the list starts collapsed so it never swallows the transcript */
const AUTO_COLLAPSE_FROM = 4

export function EditQueueCard({
  items,
  editor,
  busy,
  onEditInstruction,
  onRemove,
  onDiscardAll,
  onSend,
  onFocus,
}: Props): React.JSX.Element | null {
  const { t } = useI18n()
  /** null = follow the length-based default; set once the user clicks the chevron */
  const [manualFold, setManualFold] = useState<boolean | null>(null)
  const [editingQid, setEditingQid] = useState<string | null>(null)
  const [draft, setDraft] = useState('')
  const [confirmDiscard, setConfirmDiscard] = useState(false)

  if (items.length === 0) return null
  const folded = manualFold ?? items.length >= AUTO_COLLAPSE_FROM
  const liveCount = items.filter((item) => resolveQueueItem(editor, item).target !== null).length

  return (
    <div className="ai-queue">
      <div className="ai-queue-head">
        <span className="ai-queue-title">{t('aiQueueTitle')}</span>
        <span className="ai-queue-count">
          {items.length}/{EDIT_QUEUE_MAX}
        </span>
        <button
          className={`ai-queue-fold${folded ? ' folded' : ''}`}
          onClick={() => setManualFold(!folded)}
          aria-expanded={!folded}
        >
          <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden>
            <path
              d="M4 6l4 4 4-4"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.5"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </svg>
        </button>
      </div>
      {!folded && (
        <>
          <div className="ai-queue-hint">{t('aiQueueHint')}</div>
          <ol className="ai-queue-list">
            {items.map((item, i) => {
              const resolved = resolveQueueItem(editor, item)
              const excerpt = resolved.target?.excerpt ?? item.capturedText
              return (
                <li
                  key={item.qid}
                  className={`ai-queue-row${resolved.target ? '' : ' stale'}`}
                  data-tip={resolved.target ? undefined : t('aiQueueOrphan')}
                  onClick={() => {
                    if (editingQid !== item.qid && resolved.target) onFocus(item.qid)
                  }}
                >
                  <span className="ai-queue-ord">{i + 1}</span>
                  {editingQid === item.qid ? (
                    <input
                      className="ai-queue-edit-input"
                      value={draft}
                      autoFocus
                      maxLength={EDIT_INSTRUCTION_MAX}
                      onChange={(e) => setDraft(e.target.value)}
                      onBlur={() => {
                        const next = draft.trim()
                        if (next) onEditInstruction(item.qid, next)
                        setEditingQid(null)
                      }}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' && !e.nativeEvent.isComposing) {
                          e.preventDefault()
                          e.currentTarget.blur()
                        } else if (e.key === 'Escape') {
                          e.preventDefault()
                          setEditingQid(null)
                        }
                      }}
                    />
                  ) : (
                    <span className="ai-queue-text" title={item.instruction}>
                      {excerpt && <span className="ai-queue-target">{truncate(excerpt, 24)}</span>}
                      {item.instruction}
                    </span>
                  )}
                  <span className="ai-queue-row-actions" onClick={(e) => e.stopPropagation()}>
                    <button
                      className="ai-queue-row-btn"
                      data-tip={t('aiQueueRowEdit')}
                      aria-label={t('aiQueueRowEdit')}
                      disabled={busy}
                      onClick={() => {
                        setDraft(item.instruction)
                        setEditingQid(item.qid)
                      }}
                    >
                      <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden>
                        <path
                          d="M11.2 2.8l2 2L5.6 12.4l-2.6.6.6-2.6z"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="1.3"
                          strokeLinejoin="round"
                        />
                      </svg>
                    </button>
                    <button
                      className="ai-queue-row-btn"
                      data-tip={t('ribbonGroupDelete')}
                      aria-label={t('ribbonGroupDelete')}
                      disabled={busy}
                      onClick={() => onRemove(item.qid)}
                    >
                      <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden>
                        <path
                          d="M4 4l8 8M12 4l-8 8"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="1.3"
                          strokeLinecap="round"
                        />
                      </svg>
                    </button>
                  </span>
                </li>
              )
            })}
          </ol>
        </>
      )}
      <div className="ai-queue-foot">
        {confirmDiscard ? (
          <>
            <span className="ai-queue-confirm">
              {t('aiQueueDiscardConfirm', { count: items.length })}
            </span>
            <button className="ai-queue-discard" onClick={() => setConfirmDiscard(false)}>
              {t('appCancel')}
            </button>
            <button
              className="ai-queue-send"
              onClick={() => {
                setConfirmDiscard(false)
                onDiscardAll()
              }}
            >
              {t('aiQueueDiscard')}
            </button>
          </>
        ) : (
          <>
            <button
              className="ai-queue-discard"
              disabled={busy}
              onClick={() => setConfirmDiscard(true)}
            >
              {t('aiQueueDiscard')}
            </button>
            <button className="ai-queue-send" disabled={busy || liveCount === 0} onClick={onSend}>
              {t('aiQueueSend', { count: liveCount })}
            </button>
          </>
        )}
      </div>
    </div>
  )
}
