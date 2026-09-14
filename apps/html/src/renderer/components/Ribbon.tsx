import { useEffect, useRef, useState } from 'react'
import {
  HomeButton,
  RibbonCollapseButton,
  RibbonExpandButton,
  useDismissablePopover,
  useRibbonCollapse,
} from '@genoffice/ui'
import { useI18n } from '../i18n/locale'
import type { StringKey } from '../i18n/locale'
import { GensparkMark } from '../ai/AiPanel'
import type { InsertKind, InsertOptions } from '../document/insert-presets'
import {
  IconBullets,
  IconButton,
  IconChevronDown,
  IconCode,
  IconDivider,
  IconExpand,
  IconGlobe,
  IconHeading,
  IconPlay,
  IconPalette,
  IconPicture,
  IconPilcrow,
  IconPlus,
  IconPreview,
  IconRedo,
  IconSave,
  IconSearch,
  IconSection,
  IconSplitView,
  IconSummarize,
  IconTable,
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
  canInsert: boolean
  /** images come from a picked file by default; `url` places a remote image instead; tables take the picker's rows × cols */
  onInsert: (kind: InsertKind, opts?: InsertOptions) => void
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

const INSERT_LABEL: Record<InsertKind, StringKey> = {
  heading: 'insertHeading',
  paragraph: 'insertParagraph',
  list: 'insertList',
  button: 'insertButton',
  image: 'insertImage',
  table: 'insertTable',
  section: 'insertSection',
  divider: 'insertDivider',
}

type InsertIcon = (p: { size?: number }) => React.JSX.Element

/** the everyday kinds sit on the ribbon; the rest live under "More" */
const MORE_KINDS: Array<{ kind: InsertKind; Icon: InsertIcon }> = [
  { kind: 'list', Icon: IconBullets },
  { kind: 'button', Icon: IconButton },
  { kind: 'section', Icon: IconSection },
  { kind: 'divider', Icon: IconDivider },
]

/** which insert popover is open: the image-URL form, the table size grid or the "More" menu */
type InsertPopover = 'imageUrl' | 'table' | 'more'

const TABLE_PICKER_ROWS = 8
const TABLE_PICKER_COLS = 10

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

  const [insertPop, setInsertPop] = useState<InsertPopover | null>(null)
  const [imageUrl, setImageUrl] = useState('')
  /** hovered table size in the picker grid; 0 × 0 = nothing hovered */
  const [grid, setGrid] = useState({ r: 0, c: 0 })
  const imageRef = useRef<HTMLDivElement>(null)
  const tableRef = useRef<HTMLDivElement>(null)
  const moreRef = useRef<HTMLDivElement>(null)
  const closeInsert = () => {
    setInsertPop(null)
    setImageUrl('')
    setGrid({ r: 0, c: 0 })
  }
  const toggleInsert = (pop: InsertPopover) =>
    insertPop === pop ? closeInsert() : setInsertPop(pop)
  useDismissablePopover(insertPop !== null, closeInsert, {
    inside: () => [imageRef.current, tableRef.current, moreRef.current],
  })
  useEffect(() => {
    if (insertPop === null) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') closeInsert()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [insertPop])
  const submitImageUrl = () => {
    const url = imageUrl.trim()
    if (!url) return
    closeInsert()
    p.onInsert('image', { url })
  }
  const insert = (kind: InsertKind, opts?: InsertOptions) => {
    closeInsert()
    p.onInsert(kind, opts)
  }

  const off = p.disabled
  const insertOff = off || !p.canInsert

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
        <HomeButton />
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

        <div className="ribbon-group">
          <div className="ribbon-group-items">
            {/* the ribbon has no Insert tab: the row itself says what these buttons do */}
            <span className="rb-group-lead" aria-hidden="true">
              {t('ribbonGroupInsert')}
            </span>
            <button
              type="button"
              className="rb-btn rb-view"
              data-tip={t('insertHeading')}
              disabled={insertOff}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => insert('heading')}
            >
              <IconHeading size={ICON} />
              <span>{t('insertHeading')}</span>
            </button>
            <button
              type="button"
              className="rb-btn rb-view"
              data-tip={t('insertParagraph')}
              disabled={insertOff}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => insert('paragraph')}
            >
              <IconPilcrow size={ICON} />
              <span>{t('insertParagraph')}</span>
            </button>
            <div className="rb-menu-wrap rb-split" ref={imageRef}>
              <button
                type="button"
                className="rb-btn rb-view rb-split-main"
                data-tip={t('insertImage')}
                disabled={insertOff}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => insert('image')}
              >
                <IconPicture size={ICON} />
                <span>{t('insertImage')}</span>
              </button>
              <button
                type="button"
                className={`rb-btn rb-split-caret${insertPop === 'imageUrl' ? ' active' : ''}`}
                data-tip={t('insertImageUrl')}
                aria-label={t('insertImageUrl')}
                aria-haspopup="dialog"
                aria-expanded={insertPop === 'imageUrl'}
                disabled={insertOff}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => toggleInsert('imageUrl')}
              >
                <IconChevronDown size={14} />
              </button>
              {insertPop === 'imageUrl' && (
                <form
                  className="rb-menu rb-url-form"
                  onSubmit={(e) => {
                    e.preventDefault()
                    submitImageUrl()
                  }}
                >
                  <input
                    type="url"
                    autoFocus
                    placeholder="https://"
                    value={imageUrl}
                    onChange={(e) => setImageUrl(e.target.value)}
                  />
                  <button type="submit" className="rb-url-submit" disabled={!imageUrl.trim()}>
                    {t('insertConfirm')}
                  </button>
                </form>
              )}
            </div>
            <div className="rb-menu-wrap" ref={tableRef}>
              <button
                type="button"
                className={`rb-btn rb-view${insertPop === 'table' ? ' active' : ''}`}
                data-tip={t('insertTable')}
                aria-haspopup="dialog"
                aria-expanded={insertPop === 'table'}
                disabled={insertOff}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => toggleInsert('table')}
              >
                <IconTable size={ICON} />
                <span>{t('insertTable')}</span>
                <IconChevronDown size={14} />
              </button>
              {insertPop === 'table' && (
                <div
                  className="table-picker"
                  role="dialog"
                  aria-label={t('insertTablePickSize')}
                  onMouseLeave={() => setGrid({ r: 0, c: 0 })}
                >
                  <div className="table-picker-title">
                    {grid.r > 0
                      ? t('insertTableSize', { r: grid.r, c: grid.c })
                      : t('insertTablePickSize')}
                  </div>
                  <div className="table-picker-grid">
                    {Array.from({ length: TABLE_PICKER_ROWS }, (_, ri) =>
                      Array.from({ length: TABLE_PICKER_COLS }, (_, ci) => (
                        <button
                          key={`${ri}-${ci}`}
                          type="button"
                          className={`table-cell${ri < grid.r && ci < grid.c ? ' hot' : ''}`}
                          aria-label={t('insertTableSize', { r: ri + 1, c: ci + 1 })}
                          onMouseEnter={() => setGrid({ r: ri + 1, c: ci + 1 })}
                          onMouseDown={(e) => e.preventDefault()}
                          onClick={() => insert('table', { rows: ri + 1, cols: ci + 1 })}
                        />
                      )),
                    )}
                  </div>
                </div>
              )}
            </div>
            <div className="rb-menu-wrap" ref={moreRef}>
              <button
                type="button"
                className={`rb-btn rb-view${insertPop === 'more' ? ' active' : ''}`}
                data-tip={t('insertMore')}
                aria-haspopup="menu"
                aria-expanded={insertPop === 'more'}
                disabled={insertOff}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => toggleInsert('more')}
              >
                <IconPlus size={ICON} />
                <span>{t('insertMore')}</span>
                <IconChevronDown size={14} />
              </button>
              {insertPop === 'more' && (
                <div className="rb-menu" role="menu">
                  {MORE_KINDS.map(({ kind, Icon }) => (
                    <button
                      key={kind}
                      type="button"
                      role="menuitem"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => insert(kind)}
                    >
                      <Icon size={16} />
                      {t(INSERT_LABEL[kind])}
                    </button>
                  ))}
                </div>
              )}
            </div>
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
