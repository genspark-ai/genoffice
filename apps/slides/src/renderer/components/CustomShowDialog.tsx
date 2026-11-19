/**
 * Custom show management dialog — modeled on PowerPoint "Slide Show → Custom Shows".
 *
 * Left column: show list (new/delete/play); right side: the selected show's name, included-page
 * checkboxes and order adjustment. Data is owned by App (persisted per document to localStorage);
 * this only sends back the whole new list.
 */
import React, { useState } from 'react'
import { useI18n } from '../i18n/locale'
import type { CustomShow } from '../slideshow-utils'

/** Generate a show id (timestamp + random suffix, avoiding collisions when created in the same millisecond) */
function newShowId(): string {
  return `cs-${Date.now().toString(36)}-${Math.floor(Math.random() * 1e6).toString(36)}`
}

/** A show entry index is displayable only inside the live page range. */
export function isShowSlideIndex(i: unknown, slideCount: number): i is number {
  return typeof i === 'number' && Number.isInteger(i) && i >= 0 && i < slideCount
}

export function CustomShowDialog({
  shows,
  slideCount,
  onChange,
  onPlay,
  onClose,
}: {
  shows: CustomShow[]
  /** Document page count (checkbox list; stale out-of-range indexes filtered at render/show time) */
  slideCount: number
  onChange: (shows: CustomShow[]) => void
  onPlay: (show: CustomShow) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [selId, setSelId] = useState<string | null>(shows[0]?.id ?? null)
  const sel = shows.find((s) => s.id === selId) ?? null
  // Out-of-range indexes left over after page deletion aren't shown (the show entry filters them too)
  const selIndices = sel ? sel.slideIndices.filter((i) => isShowSlideIndex(i, slideCount)) : []

  const patchSel = (patch: Partial<CustomShow>) => {
    if (sel) onChange(shows.map((s) => (s.id === sel.id ? { ...s, ...patch } : s)))
  }

  const addShow = () => {
    const show: CustomShow = {
      id: newShowId(),
      name: t('paneCsdDefaultName', { n: shows.length + 1 }),
      slideIndices: [],
    }
    onChange([...shows, show])
    setSelId(show.id)
  }

  const removeShow = () => {
    if (!sel) return
    const rest = shows.filter((s) => s.id !== sel.id)
    onChange(rest)
    setSelId(rest[0]?.id ?? null)
  }

  /** Check/uncheck a page: checking appends to the end of the playback order */
  const togglePage = (i: number) => {
    if (!sel) return
    patchSel({
      slideIndices: selIndices.includes(i) ? selIndices.filter((x) => x !== i) : [...selIndices, i],
    })
  }

  /** Move an item up/down in the playback order */
  const movePage = (posInShow: number, dir: -1 | 1) => {
    const next = posInShow + dir
    if (!sel || next < 0 || next >= selIndices.length) return
    const arr = selIndices.slice()
    ;[arr[posInShow], arr[next]] = [arr[next]!, arr[posInShow]!]
    patchSel({ slideIndices: arr })
  }

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div
        className="modal csd"
        role="dialog"
        aria-modal="true"
        aria-labelledby="csd-title"
        onClick={(e) => e.stopPropagation()}
      >
        <h2 id="csd-title">{t('paneCsdTitle')}</h2>
        <div className="csd-body">
          <div className="csd-shows">
            <div
              className="csd-list"
              role="listbox"
              aria-label={t('paneCsdTitle')}
              onKeyDown={(e) => {
                // Arrow-key selection for the show list (buttons stay natively activatable)
                if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return
                e.preventDefault()
                const ids = shows.map((s) => s.id)
                const at = ids.indexOf(selId ?? '')
                const next = e.key === 'ArrowDown' ? at + 1 : at - 1
                if (next < 0 || next >= ids.length) return
                setSelId(ids[next]!)
                const options = e.currentTarget.querySelectorAll<HTMLElement>('[role="option"]')
                options[next]?.focus()
              }}
            >
              {shows.length === 0 && <div className="csd-empty">{t('paneCsdEmpty')}</div>}
              {shows.map((s) => (
                <button
                  key={s.id}
                  role="option"
                  aria-selected={s.id === selId}
                  className={`csd-item${s.id === selId ? ' active' : ''}`}
                  onClick={() => setSelId(s.id)}
                >
                  {s.name}
                  <span className="csd-item-count">
                    {t('paneCsdPageCount', {
                      n: s.slideIndices.filter((i) => isShowSlideIndex(i, slideCount)).length,
                    })}
                  </span>
                </button>
              ))}
            </div>
            <div className="csd-list-actions">
              <button onClick={addShow}>{t('paneCsdNew')}</button>
              <button disabled={!sel} onClick={removeShow}>
                {t('paneCsdDelete')}
              </button>
            </div>
          </div>
          {sel ? (
            <div className="csd-detail">
              <label>
                {t('paneCsdName')}
                <input
                  type="text"
                  value={sel.name}
                  onChange={(e) => patchSel({ name: e.target.value })}
                  placeholder={t('paneCsdNamePlaceholder')}
                />
              </label>
              <div className="csd-columns">
                <div className="csd-col">
                  <div className="csd-col-title">{t('paneCsdAllSlides')}</div>
                  <div className="csd-list csd-pages">
                    {Array.from({ length: slideCount }, (_, i) => (
                      <label key={i} className="csd-page">
                        <input
                          type="checkbox"
                          checked={selIndices.includes(i)}
                          onChange={() => togglePage(i)}
                        />
                        {t('paneCsdSlideN', { n: i + 1 })}
                      </label>
                    ))}
                  </div>
                </div>
                <div className="csd-col">
                  <div className="csd-col-title">{t('paneCsdOrder')}</div>
                  <div className="csd-list csd-pages">
                    {selIndices.length === 0 && (
                      <div className="csd-empty">{t('paneCsdOrderEmpty')}</div>
                    )}
                    {selIndices.map((slideIdx, posInShow) => (
                      <div key={`${slideIdx}-${posInShow}`} className="csd-order-item">
                        <span>
                          {posInShow + 1}. {t('paneCsdSlideN', { n: slideIdx + 1 })}
                        </span>
                        <span className="csd-order-btns">
                          <button
                            disabled={posInShow === 0}
                            onClick={() => movePage(posInShow, -1)}
                            data-tip={t('paneMoveUp')}
                            aria-label={t('paneMoveUp')}
                          >
                            ↑
                          </button>
                          <button
                            disabled={posInShow === selIndices.length - 1}
                            onClick={() => movePage(posInShow, 1)}
                            data-tip={t('paneMoveDown')}
                            aria-label={t('paneMoveDown')}
                          >
                            ↓
                          </button>
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          ) : (
            <div className="csd-detail csd-empty">{t('paneCsdCreateHint')}</div>
          )}
        </div>
        <div className="modal-actions">
          <button onClick={onClose}>{t('paneCsdClose')}</button>
          <button
            className="primary"
            disabled={!sel || selIndices.length === 0}
            onClick={() => sel && onPlay(sel)}
          >
            {t('paneCsdPlay')}
          </button>
        </div>
      </div>
    </div>
  )
}
