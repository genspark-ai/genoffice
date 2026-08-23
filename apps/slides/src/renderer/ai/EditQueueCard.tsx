/**
 * Pending element-scoped edits, shown above the composer. Rows are resolved
 * against the live deck on every render, so the element summary follows the
 * current content instead of a snapshot taken at annotation time.
 */
import React, { useState } from 'react'
import type { RenderSlide } from '@genoffice/pptx-render'
import { useI18n } from '../i18n/locale'
import {
  describeNode,
  EDIT_INSTRUCTION_MAX,
  NODE_NOUN_KEY,
  resolveQueueItem,
  truncate,
  type EditQueueItem,
} from './edit-queue'

interface Props {
  items: EditQueueItem[]
  slides: RenderSlide[]
  busy: boolean
  onEditInstruction: (key: string, instruction: string) => void
  onRemove: (key: string) => void
  onDiscardAll: () => void
  onSend: () => void
  /** Jump to the page carrying this item and select its elements */
  onFocus: (key: string) => void
}

/** Beyond this the list starts collapsed so it never swallows the transcript */
const AUTO_COLLAPSE_FROM = 4

export function EditQueueCard({
  items,
  slides,
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
  const [editingKey, setEditingKey] = useState<string | null>(null)
  const [draft, setDraft] = useState('')
  const [confirmDiscard, setConfirmDiscard] = useState(false)

  if (items.length === 0) return null
  const folded = manualFold ?? items.length >= AUTO_COLLAPSE_FROM

  return (
    <div className="ai-queue">
      <div className="ai-queue-head">
        <span className="ai-queue-title">{t('aiQueueTitle')}</span>
        <span className="ai-queue-count">{t('aiQueueCount', { count: items.length })}</span>
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
              const resolved = resolveQueueItem(slides, item)
              const page = (resolved.ok ? resolved.slideIndex : item.slideIndex) + 1
              const summary = resolved.ok
                ? resolved.nodes
                    .map((n) => {
                      const desc = describeNode(n)
                      return desc.text ? truncate(desc.text, 24) : t(NODE_NOUN_KEY[desc.type])
                    })
                    .join(' / ')
                : ''
              return (
                <li
                  key={item.key}
                  className={`ai-queue-row${resolved.ok ? '' : ' stale'}`}
                  onClick={() => {
                    if (editingKey !== item.key && resolved.ok) onFocus(item.key)
                  }}
                >
                  <span className="ai-queue-ord">{i + 1}</span>
                  <span className="ai-queue-page">{t('aiScopeSlide', { n: page })}</span>
                  {editingKey === item.key ? (
                    <input
                      className="ai-queue-edit-input"
                      value={draft}
                      autoFocus
                      maxLength={EDIT_INSTRUCTION_MAX}
                      onChange={(e) => setDraft(e.target.value)}
                      onBlur={() => {
                        const next = draft.trim()
                        if (next) onEditInstruction(item.key, next)
                        setEditingKey(null)
                      }}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' && !e.nativeEvent.isComposing) {
                          e.preventDefault()
                          e.currentTarget.blur()
                        } else if (e.key === 'Escape') {
                          e.preventDefault()
                          setEditingKey(null)
                        }
                      }}
                    />
                  ) : (
                    <span className="ai-queue-text" title={item.instruction}>
                      {summary && <span className="ai-queue-target">{summary}</span>}
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
                        setEditingKey(item.key)
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
                      data-tip={t('appCtxDelete')}
                      aria-label={t('appCtxDelete')}
                      disabled={busy}
                      onClick={() => onRemove(item.key)}
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
              {t('paneCancel')}
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
            <button className="ai-queue-send" disabled={busy} onClick={onSend}>
              {t('aiQueueSend', { count: items.length })}
            </button>
          </>
        )}
      </div>
    </div>
  )
}
