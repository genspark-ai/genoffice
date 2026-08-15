/**
 * Animation pane (right side) — the current page's animation order list: select/reorder/delete/change trigger.
 * The source of truth lives in the main process (the pptx <p:timing>); the list is owned by App
 * and written back wholesale (setAnimations overwrite-style); this only displays and sends back intents.
 */
import React from 'react'
import type { AnimEffectKind, AnimationItem } from '../../shared/ipc'
import { animClassOf } from '../animation-play'
import { useI18n, type StringKey } from '../i18n/locale'
import { IconSidebarCollapse } from './icons'

const EFFECT_KEY: Record<AnimEffectKind, StringKey> = {
  appear: 'paneAnimEffAppear',
  fade: 'paneAnimEffFade',
  flyIn: 'paneAnimEffFlyIn',
  wipe: 'paneAnimEffWipe',
  wipeDown: 'paneAnimEffWipeDown',
  splitIn: 'paneAnimEffSplitIn',
  bounce: 'paneAnimEffBounce',
  flipIn: 'paneAnimEffFlipIn',
  zoom: 'paneAnimEffZoom',
  pulse: 'paneAnimEffPulse',
  spin: 'paneAnimEffSpin',
  grow: 'paneAnimEffGrow',
  teeter: 'paneAnimEffTeeter',
  disappear: 'paneAnimEffDisappear',
  fadeOut: 'paneAnimEffFadeOut',
  flyOut: 'paneAnimEffFlyOut',
  wipeOut: 'paneAnimEffWipeOut',
  shrink: 'paneAnimEffShrink',
  zoomOut: 'paneAnimEffZoomOut',
  motionPath: 'paneAnimEffMotionPath',
}

const TRIGGER_GLYPH = { onClick: '🖱', withPrev: '⇉', afterPrev: '⏱' } as const
const TRIGGER_KEY = {
  onClick: 'paneAnimTrigOnClick',
  withPrev: 'paneAnimTrigWithPrev',
  afterPrev: 'paneAnimTrigAfterPrev',
} as const satisfies Record<string, StringKey>

interface Props {
  slideIndex: number
  items: AnimationItem[]
  /** Pane selection (the timing controls bind to it); -1 = none */
  selected: number
  onSelect: (idx: number) => void
  /** Move up/down one position */
  onMove: (idx: number, dir: -1 | 1) => void
  onDelete: (idx: number) => void
  /** Play preview */
  onPreview: () => void
  onCollapse: () => void
}

export function AnimationPane({
  slideIndex,
  items,
  selected,
  onSelect,
  onMove,
  onDelete,
  onPreview,
  onCollapse,
}: Props) {
  const { t } = useI18n()
  // Step numbering: onClick increments; withPrev/afterPrev share the previous entry's number (0 = auto-play)
  const stepNos: number[] = []
  let no = 0
  for (const it of items) {
    if (it.trigger === 'onClick') no++
    stepNos.push(no)
  }

  return (
    <aside className="comments-pane anim-pane">
      <div className="ai-panel-header">
        <span className="ai-panel-title">{t('paneAnimTitle', { n: slideIndex + 1 })}</span>
        <div className="ai-panel-header-actions">
          <button
            className="ai-header-btn"
            disabled={items.length === 0}
            onClick={onPreview}
            data-tip={t('paneAnimPreview')}
            aria-label={t('paneAnimPreview')}
          >
            ▶
          </button>
          <button
            className="ai-header-btn"
            onClick={onCollapse}
            data-tip={t('paneAnimCollapse')}
            aria-label={t('paneAnimCollapse')}
          >
            <IconSidebarCollapse size={15} />
          </button>
        </div>
      </div>

      <div className="comments-list">
        {items.length === 0 && (
          <div className="comments-empty">
            <p>{t('paneAnimEmpty')}</p>
            <p className="comments-empty-sub">{t('paneAnimEmptySub')}</p>
          </div>
        )}
        {items.map((it, i) => (
          <div
            key={i}
            className={`anim-row anim-${animClassOf(it.effect)} ${selected === i ? 'selected' : ''}`}
            onClick={() => onSelect(selected === i ? -1 : i)}
            data-tip={`${t(TRIGGER_KEY[it.trigger])} · ${t('paneAnimSeconds', { sec: (it.durationMs / 1000).toFixed(2) })}${it.delayMs ? ` · ${t('paneAnimDelaySeconds', { sec: (it.delayMs / 1000).toFixed(2) })}` : ''}`}
          >
            <span className="anim-row-no">{stepNos[i] || ''}</span>
            <span className="anim-row-trigger" data-tip={t(TRIGGER_KEY[it.trigger])}>
              {TRIGGER_GLYPH[it.trigger]}
            </span>
            <span className="anim-row-main">
              <span className="anim-row-effect">{t(EFFECT_KEY[it.effect])}</span>
              <span className="anim-row-target">
                {it.targetName}
                {it.paragraph != null && ` · ${t('paneAnimParaN', { n: it.paragraph + 1 })}`}
              </span>
            </span>
            <span className="anim-row-actions" onClick={(e) => e.stopPropagation()}>
              <button
                disabled={i === 0}
                data-tip={t('paneMoveUp')}
                aria-label={t('paneMoveUp')}
                onClick={() => onMove(i, -1)}
              >
                ▲
              </button>
              <button
                disabled={i === items.length - 1}
                data-tip={t('paneMoveDown')}
                aria-label={t('paneMoveDown')}
                onClick={() => onMove(i, 1)}
              >
                ▼
              </button>
              <button
                data-tip={t('paneAnimDelete')}
                aria-label={t('paneAnimDelete')}
                onClick={() => onDelete(i)}
              >
                ✕
              </button>
            </span>
          </div>
        ))}
      </div>
    </aside>
  )
}
