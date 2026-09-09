import { useEffect, useRef, useState } from 'react'
import {
  RibbonCollapseButton,
  RibbonExpandButton,
  useDismissablePopover,
  useRibbonCollapse,
} from '@genoffice/ui'
import { useI18n } from '../i18n/locale'
import type { StringKey } from '../i18n/locale'
import { GensparkMark } from '../ai/AiPanel'
import {
  IconChevronDown,
  IconCode,
  IconExpand,
  IconGlobe,
  IconPlay,
  IconPalette,
  IconPreview,
  IconRedo,
  IconSave,
  IconSearch,
  IconSplitView,
  IconSummarize,
  IconUndo,
  IconWand,
} from './icons'

export type ViewMode = 'preview' | 'split' | 'source'

const THEME_DIRECTIONS: StringKey[] = [
  'aiThemeMinimal',
  'aiThemeEditorial',
  'aiThemeTech',
  'aiThemePlayful',
  'aiThemeDark',
]

export const VIEW_MODES: ViewMode[] = ['preview', 'split', 'source']

const VIEW_LABEL: Record<ViewMode, StringKey> = {
  preview: 'viewPreview',
  split: 'viewSplit',
  source: 'viewSource',
}

const VIEW_ICON: Record<ViewMode, (p: { size?: number }) => React.JSX.Element> = {
  preview: IconPreview,
  split: IconSplitView,
  source: IconCode,
}

interface Props {
  disabled: boolean
  dirty: boolean
  onSave: () => void
  onFind: () => void
  canUndo: boolean
  canRedo: boolean
  onUndo: () => void
  onRedo: () => void
  autoSave: boolean
  onToggleAutoSave: (on: boolean) => void
  view: ViewMode
  onView: (view: ViewMode) => void
  aiOpen: boolean
  onToggleAi: () => void
  /** page-wide AI actions: send this instruction to the assistant right away */
  onAiPreset: (text: string) => void
  canvasMode: CanvasMode
  onPresent: (kind: PresentKind) => void
}

export type CanvasMode = 'edit' | 'present'

/** tab = chrome-free in this view; fullscreen = tab + the whole screen; newTab = a separate present tab/window */
export type PresentKind = 'tab' | 'fullscreen' | 'newTab'
const PRESENT_ITEMS: Array<{
  kind: PresentKind
  label: StringKey
  Icon: (p: { size?: number }) => React.JSX.Element
}> = [
  { kind: 'tab', label: 'presentInTab', Icon: IconExpand },
  { kind: 'fullscreen', label: 'presentFullscreen', Icon: IconPlay },
  { kind: 'newTab', label: 'presentNewTab', Icon: IconGlobe },
]

const ICON = 20

export function Ribbon(p: Props) {
  const { t } = useI18n()
  const collapse = useRibbonCollapse('htmlapp.ribbonCollapsed')
  const [themeOpen, setThemeOpen] = useState(false)
  const themeRef = useRef<HTMLDivElement>(null)
  useDismissablePopover(themeOpen, () => setThemeOpen(false), {
    inside: () => [themeRef.current],
  })
  useEffect(() => {
    if (!themeOpen) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setThemeOpen(false)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [themeOpen])
  const [presentOpen, setPresentOpen] = useState(false)
  const presentRef = useRef<HTMLDivElement>(null)
  useDismissablePopover(presentOpen, () => setPresentOpen(false), {
    inside: () => [presentRef.current],
  })
  useEffect(() => {
    if (!presentOpen) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setPresentOpen(false)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [presentOpen])

  const off = p.disabled

  return (
    <div className={`ribbon ${collapse.rootClass}`} ref={collapse.rootRef}>
      <div className="ribbon-tabs">
        <button
          type="button"
          className="qa-btn"
          data-tip={t('save')}
          aria-label={t('save')}
          disabled={off || !p.dirty}
          onMouseDown={(e) => e.preventDefault()}
          onClick={p.onSave}
        >
          <IconSave size={16} />
        </button>
        <button
          type="button"
          className="qa-btn"
          data-tip={t('undo')}
          aria-label={t('undo')}
          disabled={off || !p.canUndo}
          onMouseDown={(e) => e.preventDefault()}
          onClick={p.onUndo}
        >
          <IconUndo size={16} />
        </button>
        <button
          type="button"
          className="qa-btn"
          data-tip={t('redo')}
          aria-label={t('redo')}
          disabled={off || !p.canRedo}
          onMouseDown={(e) => e.preventDefault()}
          onClick={p.onRedo}
        >
          <IconRedo size={16} />
        </button>
        <button
          type="button"
          className="qa-btn"
          data-tip={t('findTip')}
          aria-label={t('findTip')}
          disabled={off}
          onMouseDown={(e) => e.preventDefault()}
          onClick={p.onFind}
        >
          <IconSearch size={16} />
        </button>
        <label className={`autosave-toggle${p.autoSave ? ' on' : ''}`} data-tip={t('autoSaveTip')}>
          <span className="autosave-knob" />
          <span className="autosave-text">{t('autoSave')}</span>
          <input
            type="checkbox"
            checked={p.autoSave}
            onChange={(e) => p.onToggleAutoSave(e.target.checked)}
          />
        </label>
        <RibbonExpandButton state={collapse} label={t('ribbonExpand')} />
      </div>

      <div className="ribbon-body" data-ribbon-body="">
        <div className="ribbon-group">
          <div className="ribbon-group-items">
            <button
              type="button"
              className={`rb-big ai-entry${p.aiOpen ? ' active' : ''}`}
              data-tip={t('aiOpenAssistant')}
              aria-pressed={p.aiOpen}
              disabled={off}
              onMouseDown={(e) => e.preventDefault()}
              onClick={p.onToggleAi}
            >
              <span className="rb-big-icon">
                <GensparkMark size={26} />
              </span>
              <span>Genspark AI</span>
            </button>
            <button
              type="button"
              className="rb-big ai-entry"
              data-tip={t('aiRestyleBtn')}
              disabled={off}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => p.onAiPreset(t('aiRestylePrompt'))}
            >
              <span className="rb-big-icon">
                <span className="ai-feature-icon" aria-hidden="true">
                  <IconWand size={24} />
                </span>
              </span>
              <span>{t('aiRestyleBtn')}</span>
            </button>
            <div className="rb-menu-wrap" ref={themeRef}>
              <button
                type="button"
                className={`rb-big ai-entry${themeOpen ? ' active' : ''}`}
                data-tip={t('aiThemeBtn')}
                aria-haspopup="menu"
                aria-expanded={themeOpen}
                disabled={off}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => setThemeOpen((v) => !v)}
              >
                <span className="rb-big-icon">
                  <span className="ai-feature-icon" aria-hidden="true">
                    <IconPalette size={24} />
                  </span>
                </span>
                <span>{t('aiThemeBtn')}</span>
              </button>
              {themeOpen && (
                <div className="rb-menu" role="menu">
                  {THEME_DIRECTIONS.map((key) => (
                    <button
                      key={key}
                      type="button"
                      role="menuitem"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => {
                        setThemeOpen(false)
                        p.onAiPreset(t('aiThemePrompt', { direction: t(key) }))
                      }}
                    >
                      {t(key)}
                    </button>
                  ))}
                </div>
              )}
            </div>
            <button
              type="button"
              className="rb-big ai-entry"
              data-tip={t('aiSummarizeBtn')}
              disabled={off}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => p.onAiPreset(t('aiSummarizePrompt'))}
            >
              <span className="rb-big-icon">
                <span className="ai-feature-icon" aria-hidden="true">
                  <IconSummarize size={24} />
                </span>
              </span>
              <span>{t('aiSummarizeBtn')}</span>
            </button>
          </div>
        </div>

        <div className="rb-sep" />

        <div className="ribbon-group" role="tablist">
          <div className="ribbon-group-items">
            {VIEW_MODES.map((mode) => {
              const Icon = VIEW_ICON[mode]
              return (
                <button
                  key={mode}
                  type="button"
                  role="tab"
                  className={`rb-btn rb-view${p.view === mode ? ' active' : ''}`}
                  aria-selected={p.view === mode}
                  data-tip={t(VIEW_LABEL[mode])}
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => p.onView(mode)}
                >
                  <Icon size={ICON} />
                  <span>{t(VIEW_LABEL[mode])}</span>
                </button>
              )
            })}
          </div>
        </div>

        <div className="rb-sep" />

        <div className="ribbon-group">
          <div className="ribbon-group-items">
            <div className="rb-menu-wrap" ref={presentRef}>
              <button
                type="button"
                className={`rb-btn rb-view${p.canvasMode === 'present' ? ' active' : ''}`}
                data-tip={t('modePresent')}
                aria-haspopup="menu"
                aria-expanded={presentOpen}
                disabled={off}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => setPresentOpen((v) => !v)}
              >
                <IconPlay size={ICON} />
                <span>{t('modePresent')}</span>
                <IconChevronDown size={14} />
              </button>
              {presentOpen && (
                <div className="rb-menu" role="menu">
                  {PRESENT_ITEMS.map(({ kind, label, Icon }) => (
                    <button
                      key={kind}
                      type="button"
                      role="menuitem"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => {
                        setPresentOpen(false)
                        p.onPresent(kind)
                      }}
                    >
                      <Icon size={16} />
                      {t(label)}
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
      <RibbonCollapseButton
        state={collapse}
        labels={{ collapse: t('ribbonCollapse'), pin: t('ribbonPin') }}
      />
    </div>
  )
}
