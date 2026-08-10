/**
 * Find/replace floating panel (⌘F) — modeled on PowerPoint "Home → Find/Replace".
 * Find works on render-tree text (element-granularity hits, page jump + select); replace goes
 * through the main-process model layer (in-run matching, byte-faithful patches), and on success
 * the whole RenderSlide set refreshes.
 */
import React, { useEffect, useRef, useState } from 'react'
import { useI18n } from '../i18n/locale'
import type {
  GroupRenderNode,
  RenderNode,
  RenderSlide,
  RenderTextLayout,
  ShapeRenderNode,
  TableRenderNode,
} from '@genoffice/pptx-render'

export interface FindMatch {
  slideIndex: number
  sourceId: string
  count: number
}

/** Laid-out text → plain text (skip bullet glyphs; restore spaces swallowed by wrapping; newline between paragraphs). */
function layoutText(text?: RenderTextLayout): string {
  if (!text) return ''
  let out = ''
  text.lines.forEach((l, i) => {
    if (i > 0 && l.paraStart !== false) out += '\n'
    out += l.runs
      .filter((r) => !r.isBullet)
      .map((r) => r.text)
      .join('')
    if (l.trailingSpace) out += ' '
  })
  return out
}

/** Searchable text aggregated from a node (including group children/table cells). */
function nodeText(n: RenderNode): string {
  if (n.type === 'text' || n.type === 'shape') return layoutText((n as ShapeRenderNode).text)
  if (n.type === 'table')
    return (n as TableRenderNode).cells.map((c) => layoutText(c.text)).join('\n')
  if (n.type === 'group') return (n as GroupRenderNode).children.map(nodeText).join('\n')
  return ''
}

function countHits(hay: string, needle: string, matchCase: boolean): number {
  if (!needle) return 0
  const h = matchCase ? hay : hay.toLowerCase()
  const q = matchCase ? needle : needle.toLowerCase()
  let n = 0
  for (let i = h.indexOf(q); i >= 0; i = h.indexOf(q, i + q.length)) n++
  return n
}

export function buildMatches(
  slides: RenderSlide[],
  query: string,
  matchCase: boolean,
): FindMatch[] {
  const out: FindMatch[] = []
  if (!query) return out
  slides.forEach((sl, si) => {
    for (const n of sl.nodes) {
      if (n.decoration) continue
      const count = countHits(nodeText(n), query, matchCase)
      if (count > 0) out.push({ slideIndex: si, sourceId: n.sourceId, count })
    }
  })
  return out
}

export function FindReplaceDialog({
  slides,
  onNavigate,
  onReplaced,
  onClose,
}: {
  slides: RenderSlide[]
  onNavigate: (slideIndex: number, sourceId: string) => void
  onReplaced: (slides: RenderSlide[]) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [query, setQuery] = useState('')
  const [replaceText, setReplaceText] = useState('')
  const [matchCase, setMatchCase] = useState(false)
  const [cursor, setCursor] = useState(-1)
  const [status, setStatus] = useState('')
  const findRef = useRef<HTMLInputElement>(null)

  useEffect(() => findRef.current?.focus(), [])

  // Recompute the hit list when slides change (refresh after replace) or conditions change; the cursor is invalidated
  const matches = buildMatches(slides, query, matchCase)
  const total = matches.reduce((a, m) => a + m.count, 0)

  const findNext = () => {
    if (!matches.length) {
      setStatus(t('paneFrNotFound'))
      return
    }
    const next = (cursor + 1) % matches.length
    setCursor(next)
    const m = matches[next]!
    onNavigate(m.slideIndex, m.sourceId)
    setStatus(t('paneFrMatchPos', { i: String(next + 1), n: String(matches.length) }))
  }

  const doReplace = async (all: boolean) => {
    if (!query) return
    const cur = !all && cursor >= 0 ? matches[cursor] : undefined
    if (!all && !cur) {
      findNext()
      return
    }
    const r = await window.slidesApi.findReplace({
      find: query,
      replace: replaceText,
      matchCase,
      ...(all ? {} : { firstOnly: true, slideIndex: cur!.slideIndex, elementId: cur!.sourceId }),
    })
    if (!r || !r.count || !r.slides) {
      setStatus(t('paneFrNotFound'))
      return
    }
    onReplaced(r.slides)
    setStatus(t('paneFrReplaced', { n: String(r.count) }))
    if (!all) setCursor(cursor - 1) // The current item is consumed; the next findNext lands on the following one
  }

  return (
    <div className="find-panel" onKeyDown={(e) => e.key === 'Escape' && onClose()}>
      <div className="find-panel-head">
        <span>{t('paneFrTitle')}</span>
        <button className="find-panel-close" onClick={onClose} data-tip="Esc" aria-label="Esc">
          ×
        </button>
      </div>
      <div className="find-panel-row">
        <input
          ref={findRef}
          placeholder={t('paneFrFind')}
          value={query}
          onChange={(e) => {
            setQuery(e.target.value)
            setCursor(-1)
            setStatus('')
          }}
          onKeyDown={(e) => e.key === 'Enter' && findNext()}
        />
      </div>
      <div className="find-panel-row">
        <input
          placeholder={t('paneFrReplaceWith')}
          value={replaceText}
          onChange={(e) => setReplaceText(e.target.value)}
        />
      </div>
      <div className="find-panel-row find-panel-opts">
        <label>
          <input
            type="checkbox"
            checked={matchCase}
            onChange={(e) => setMatchCase(e.target.checked)}
          />
          {t('paneFrMatchCase')}
        </label>
        <span className="find-panel-status">
          {status ||
            (query
              ? t('paneFrMatchPos', {
                  i: String(Math.max(cursor + 1, 0)),
                  n: String(matches.length),
                })
              : '')}
        </span>
      </div>
      <div className="find-panel-row find-panel-actions">
        <button onClick={findNext} disabled={!query}>
          {t('paneFrFindNext')}
        </button>
        <button onClick={() => void doReplace(false)} disabled={!query}>
          {t('paneFrReplace')}
        </button>
        <button onClick={() => void doReplace(true)} disabled={!query || !total}>
          {t('paneFrReplaceAll')}
        </button>
      </div>
    </div>
  )
}
