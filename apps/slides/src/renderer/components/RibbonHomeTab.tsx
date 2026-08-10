/** Home tab of the slides ribbon. Extracted from Ribbon.tsx. */
import { useState } from 'react'
import { platformShortcuts } from '@genoffice/i18n'
import { saveEditSelection } from '../TextEditOverlay'
import { armColorInput } from '../color-input'
import { displayFontFamily } from '../konva-adapter'
import {
  GensparkMark,
  IconAiBeautify,
  IconAiFactCheck,
  IconAiImage,
  IconAlignCenter,
  IconAlignJustify,
  IconAlignLeft,
  IconAlignRight,
  IconBullets,
  IconClearFormat,
  IconCopy,
  IconCut,
  IconFind,
  IconFormatPainter,
  IconGrowFont,
  IconIndentDec,
  IconIndentInc,
  IconLineSpacing,
  IconNewSlide,
  IconNumbered,
  IconObjAlignBottom,
  IconObjAlignCenterH,
  IconObjAlignLeft,
  IconObjAlignMiddle,
  IconObjAlignRight,
  IconObjAlignTop,
  IconObjDistributeH,
  IconObjDistributeV,
  IconObjFlipH,
  IconObjFlipV,
  IconPaste,
  IconPlayCurrent,
  IconPlayFromStart,
  IconPosition,
  IconSection,
  IconShrinkFont,
  IconSlideLayout,
} from './icons'
import {
  BIG,
  FONT_FAMILIES,
  FONT_SIZES,
  Group,
  LayoutList,
  RbCaret,
  TEXT_COLORS,
  closeSiblingPanels,
  type RibbonTabCtx,
} from './ribbon-shared'

export function RibbonHomeTab({ rb }: { rb: RibbonTabCtx }) {
  const {
    aiOpen,
    brushMode,
    canDistribute,
    canPaste,
    curBulletChar,
    curAlign,
    curFontFamily,
    curFontSizeMixed,
    curFontSizePt,
    deckEmpty,
    editing,
    formatOpen,
    hasBrushFormat,
    hasDoc,
    hasSelection,
    hasTextSelection,
    layouts,
    layoutSize,
    onAddSection,
    onAddSlide,
    onAddSlideWithLayout,
    onAiPreset,
    onAlign,
    onArrange,
    onFlip,
    onCopy,
    onCut,
    onElementTextColor,
    onFindReplace,
    onFontFamily,
    onFontSize,
    onFormat,
    onFormatBrushClick,
    onFormatBrushDoubleClick,
    onParagraphFormat,
    onPaste,
    onResetLayout,
    onSetLayout,
    onSlideShow,
    onStrike,
    onTextColor,
    onTextToggle,
    onToggleAi,
    onToggleFormat,
    arrangeOpen,
    closePanels,
    collapseOpen,
    collapsedGroups,
    colorOpen,
    commitFontDraft,
    commitSizeDraft,
    fmtBtn,
    fontDraft,
    fontOpen,
    setFontDraft,
    lastBulletColor,
    lastColor,
    layoutOpen,
    layoutPickOpen,
    lineSpacingOpen,
    onCustomBulletColor,
    onCustomTextColor,
    paraOpen,
    recentColors,
    setArrangeOpen,
    setCollapseOpen,
    setColorOpen,
    setFontOpen,
    setLastColor,
    setLayoutOpen,
    setLayoutPickOpen,
    setLineSpacingOpen,
    setParaOpen,
    setSizeDraft,
    setSizeOpen,
    setSlideShowFromStart,
    setSlideShowOpen,
    sizeDraft,
    sizeOpen,
    slideShowFromStart,
    slideShowOpen,
    t,
  } = rb
  const [hangDraft, setHangDraft] = useState('')
  // Typed-ahead font query: only what the user actually typed filters the menu
  // (opening via the caret or focusing shows the full list)
  const [fontFilter, setFontFilter] = useState('')
  const EMU_PER_PX = 9525
  const commitHangDraft = () => {
    const px = parseFloat(hangDraft.replace(',', '.'))
    if (!Number.isFinite(px) || px < 0 || !hasSelection) return
    onParagraphFormat({ bulletHangEmu: Math.round(px * EMU_PER_PX) })
  }
  return (
    <>
      <Group label="Genspark AI">
        <button
          className={`rb-big ai-entry${aiOpen ? ' active' : ''}`}
          data-tip={t('aiOpenAssistant')}
          onClick={onToggleAi}
        >
          <span className="rb-big-icon">
            <GensparkMark size={26} />
          </span>
          <span>Genspark AI</span>
        </button>
        <button
          className="rb-big ai-entry"
          disabled={!hasDoc || deckEmpty}
          data-tip={t('aiBeautifyBtn')}
          onClick={() => onAiPreset(t('aiBeautifyPrompt'), { slideShot: true })}
        >
          <span className="rb-big-icon">
            <span className="ai-feature-icon" aria-hidden="true">
              <IconAiBeautify />
            </span>
          </span>
          <span>{t('aiBeautifyBtn')}</span>
        </button>
        <button
          className="rb-big ai-entry"
          disabled={!hasDoc || deckEmpty}
          data-tip={t('aiFactCheckBtn')}
          onClick={() => onAiPreset(t('aiFactCheckPrompt'))}
        >
          <span className="rb-big-icon">
            <span className="ai-feature-icon" aria-hidden="true">
              <IconAiFactCheck />
            </span>
          </span>
          <span>{t('aiFactCheckBtn')}</span>
        </button>
        <button
          className="rb-big ai-entry"
          disabled={!hasDoc || deckEmpty}
          data-tip={t('aiImageBtn')}
          onClick={() => onAiPreset(t('aiImagePrompt'))}
        >
          <span className="rb-big-icon">
            <span className="ai-feature-icon" aria-hidden="true">
              <IconAiImage />
            </span>
          </span>
          <span>{t('aiImageBtn')}</span>
        </button>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupClipboard')}>
        <button
          className="rb-big"
          disabled={!hasDoc || !canPaste}
          onClick={onPaste}
          data-tip={canPaste ? t('ribbonPasteTip') : t('ribbonPasteTipDisabled')}
        >
          <span className="rb-big-icon">
            <IconPaste size={BIG} />
          </span>
          <span>{t('ribbonPaste')}</span>
        </button>
        <div className="rb-col rb-clip-col">
          <button
            className="rb-icon"
            disabled={!hasSelection}
            onClick={onCut}
            data-tip={t('ribbonCutTip')}
            aria-label={t('ribbonCutTip')}
          >
            <IconCut size={14} />
          </button>
          <button
            className="rb-icon"
            disabled={!hasSelection}
            onClick={onCopy}
            data-tip={t('ribbonCopyTip')}
            aria-label={t('ribbonCopyTip')}
          >
            <IconCopy size={14} />
          </button>
          <button
            className={`rb-icon${brushMode ? ' on' : ''}`}
            disabled={!hasSelection}
            data-tip={
              !hasSelection
                ? t('ribbonBrushTipNoSelection')
                : brushMode === 'continuous'
                  ? t('ribbonBrushTipContinuous')
                  : brushMode === 'once'
                    ? t('ribbonBrushTipOnce')
                    : hasBrushFormat
                      ? t('ribbonBrushTipHasFormat')
                      : t('ribbonBrushTipDefault')
            }
            aria-label={
              !hasSelection
                ? t('ribbonBrushTipNoSelection')
                : brushMode === 'continuous'
                  ? t('ribbonBrushTipContinuous')
                  : brushMode === 'once'
                    ? t('ribbonBrushTipOnce')
                    : hasBrushFormat
                      ? t('ribbonBrushTipHasFormat')
                      : t('ribbonBrushTipDefault')
            }
            onClick={onFormatBrushClick}
            onDoubleClick={(e) => {
              e.preventDefault()
              onFormatBrushDoubleClick()
            }}
          >
            <IconFormatPainter size={14} />
          </button>
        </div>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonTabSlideShow')}>
        <div className="rb-drop-wrap">
          <button
            className="rb-big"
            disabled={!hasDoc}
            onClick={() => onSlideShow(slideShowFromStart)}
            data-tip={t(slideShowFromStart ? 'ribbonFromBeginningTip' : 'ribbonFromCurrentTip')}
          >
            <span className="rb-big-icon">
              {slideShowFromStart ? (
                <IconPlayFromStart size={BIG} />
              ) : (
                <IconPlayCurrent size={BIG} />
              )}
              <span
                className={`rb-caret-hit${slideShowOpen ? ' active' : ''}`}
                onMouseDown={(e) => {
                  e.stopPropagation()
                  closeSiblingPanels(e, closePanels, 'slideShow')
                }}
                onClick={(e) => {
                  e.stopPropagation()
                  if (hasDoc) setSlideShowOpen((v) => !v)
                }}
              >
                <RbCaret />
              </span>
            </span>
            <span>{t(slideShowFromStart ? 'ribbonFromBeginning' : 'ribbonFromCurrent')}</span>
          </button>
          {slideShowOpen && (
            <div className="rb-drop rb-menu" onMouseDown={(e) => e.stopPropagation()}>
              <button
                onClick={() => {
                  setSlideShowOpen(false)
                  setSlideShowFromStart(true)
                  onSlideShow(true)
                }}
                data-tip={t('ribbonFromBeginningTip')}
              >
                <span className="rb-menu-glyph">
                  <IconPlayFromStart size={20} />
                </span>
                {t('ribbonFromBeginning')}
              </button>
              <button
                onClick={() => {
                  setSlideShowOpen(false)
                  setSlideShowFromStart(false)
                  onSlideShow(false)
                }}
                data-tip={t('ribbonFromCurrentTip')}
              >
                <span className="rb-menu-glyph">
                  <IconPlayCurrent size={20} />
                </span>
                {t('ribbonFromCurrent')}
              </button>
            </div>
          )}
        </div>
      </Group>
      <div className="ribbon-sep" />
      <Group
        label={t('ribbonGroupSlides')}
        groupId="slides"
        collapse={{
          collapsed: collapsedGroups.includes('slides'),
          open: collapseOpen === 'slides',
          onToggle: () => {
            closePanels(['collapse'])
            setCollapseOpen((v) => (v === 'slides' ? null : 'slides'))
          },
          icon: <IconNewSlide size={BIG} />,
        }}
      >
        <div className="rb-drop-wrap">
          <button
            className="rb-big"
            disabled={!hasDoc}
            onClick={onAddSlide}
            data-tip={t('ribbonNewSlideTip')}
          >
            <span className="rb-big-icon">
              <IconNewSlide size={BIG} />
              <span
                className={`rb-caret-hit${layoutOpen ? ' active' : ''}`}
                data-tip={t('ribbonChooseLayoutNew')}
                onMouseDown={(e) => {
                  e.stopPropagation()
                  closeSiblingPanels(e, closePanels, 'layout')
                }}
                onClick={(e) => {
                  e.stopPropagation()
                  if (hasDoc) setLayoutOpen((v) => !v)
                }}
              >
                <RbCaret />
              </span>
            </span>
            <span>{t('ribbonNewSlide')}</span>
          </button>
          {layoutOpen && (
            <div className="rb-drop rb-layout-panel" onMouseDown={(e) => e.stopPropagation()}>
              <div className="rb-drop-title">{t('ribbonChooseLayoutNew')}</div>
              <LayoutList
                layouts={layouts}
                size={layoutSize}
                onPick={(path) => {
                  setLayoutOpen(false)
                  onAddSlideWithLayout(path)
                }}
              />
            </div>
          )}
        </div>
        <div className="rb-drop-wrap">
          <button
            className={`rb-big ${layoutPickOpen ? 'active' : ''}`}
            disabled={!hasDoc}
            onMouseDown={(e) => {
              e.stopPropagation()
              closeSiblingPanels(e, closePanels, 'layoutPick')
            }}
            onClick={() => setLayoutPickOpen((v) => !v)}
            data-tip={t('ribbonLayoutTip')}
          >
            <span className="rb-big-icon">
              <IconSlideLayout size={BIG} />
              <RbCaret />
            </span>
            <span>{t('ribbonLayout')}</span>
          </button>
          {layoutPickOpen && (
            <div className="rb-drop rb-layout-drop" onMouseDown={(e) => e.stopPropagation()}>
              <div className="rb-drop-title">{t('ribbonChooseLayoutChange')}</div>
              <LayoutList
                layouts={layouts}
                size={layoutSize}
                onPick={(path) => {
                  setLayoutPickOpen(false)
                  onSetLayout(path)
                }}
              />
              <div className="rb-menu-div" />
              <button
                className="rb-layout-reset"
                onClick={() => {
                  setLayoutPickOpen(false)
                  onResetLayout()
                }}
              >
                {t('ribbonResetLayout')}
              </button>
            </div>
          )}
        </div>
        <button
          className="rb-big"
          disabled={!hasDoc}
          onClick={onAddSection}
          data-tip={t('ribbonAddSectionTip')}
        >
          <span className="rb-big-icon">
            <IconSection size={BIG} />
          </span>
          <span>{t('ribbonAddSection')}</span>
        </button>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupFont')}>
        <div className="rb-col">
          <div className="rb-row">
            <div className="rb-drop-wrap">
              <div
                className={`rb-font-btn rb-font-name rb-size-combo ${fontOpen ? 'active' : ''}${
                  !editing && !hasTextSelection ? ' rb-combo-disabled' : ''
                }`}
              >
                {/* Editable font field (combobox): any font name, Enter applies.
                    Same selection dance as the size combobox */}
                <input
                  className="rb-size-input rb-font-input"
                  data-keep-edit=""
                  disabled={!editing && !hasTextSelection}
                  data-tip={
                    editing || hasTextSelection ? t('ribbonGroupFont') : t('ribbonFontTipDisabled')
                  }
                  value={fontDraft ?? curFontFamily ?? ''}
                  onMouseDown={(e) => {
                    e.stopPropagation()
                    if (editing) saveEditSelection()
                  }}
                  onFocus={(e) => {
                    setFontDraft(curFontFamily ?? '')
                    setFontFilter('')
                    e.target.select()
                  }}
                  onChange={(e) => {
                    setFontDraft(e.target.value)
                    setFontFilter(e.target.value)
                    // Ribbon menus are mutually exclusive: close any open sibling
                    // popup (size/color/...) before opening the font list
                    closePanels(['font'])
                    setFontOpen(true)
                  }}
                  onBlur={() => {
                    setFontDraft(null)
                    setFontFilter('')
                  }}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      e.preventDefault()
                      commitFontDraft()
                      e.currentTarget.blur()
                    } else if (e.key === 'Escape') {
                      e.preventDefault()
                      e.currentTarget.blur()
                    }
                  }}
                />
                <button
                  className="rb-size-caret"
                  disabled={!editing && !hasTextSelection}
                  data-tip={
                    editing || hasTextSelection ? t('ribbonGroupFont') : t('ribbonFontTipDisabled')
                  }
                  aria-label={
                    editing || hasTextSelection ? t('ribbonGroupFont') : t('ribbonFontTipDisabled')
                  }
                  onMouseDown={(e) => {
                    e.preventDefault()
                    e.stopPropagation()
                    if (editing || hasTextSelection) {
                      closeSiblingPanels(e, closePanels, 'font')
                      setFontFilter('')
                      setFontOpen((v) => !v)
                    }
                  }}
                >
                  <RbCaret />
                </button>
              </div>
              {fontOpen && (
                <div
                  className="rb-drop rb-menu rb-menu-scroll rb-font-menu"
                  onMouseDown={(e) => e.stopPropagation()}
                >
                  {(curFontFamily && !FONT_FAMILIES.includes(curFontFamily)
                    ? [curFontFamily, ...FONT_FAMILIES]
                    : FONT_FAMILIES
                  )
                    .filter(
                      (f) =>
                        !fontFilter.trim() ||
                        f.toLowerCase().includes(fontFilter.trim().toLowerCase()),
                    )
                    .map((f) => (
                      <button
                        key={f}
                        className={f === curFontFamily ? 'on' : ''}
                        style={{ fontFamily: displayFontFamily(f) }}
                        onMouseDown={(e) => {
                          e.preventDefault()
                          onFontFamily(f)
                          setFontOpen(false)
                        }}
                      >
                        {f}
                      </button>
                    ))}
                </div>
              )}
            </div>
            <div className="rb-drop-wrap">
              <div
                className={`rb-font-btn rb-font-size rb-size-combo ${sizeOpen ? 'active' : ''}${
                  !editing && !hasTextSelection ? ' rb-combo-disabled' : ''
                }`}
              >
                {/* Editable size field (combobox): any positive pt value, Enter applies.
                          data-keep-edit: focusing it doesn't commit the text edit; the saved selection
                          is restored on apply so the size hits the selection */}
                <input
                  className="rb-size-input"
                  data-keep-edit=""
                  disabled={!editing && !hasTextSelection}
                  data-tip={
                    editing || hasTextSelection
                      ? t('ribbonFontSizeTip')
                      : t('ribbonFontSizeTipDisabled')
                  }
                  value={
                    sizeDraft ??
                    (curFontSizePt != null ? `${curFontSizePt}${curFontSizeMixed ? '+' : ''}` : '')
                  }
                  onMouseDown={(e) => {
                    e.stopPropagation()
                    if (editing) saveEditSelection()
                  }}
                  onFocus={(e) => {
                    setSizeDraft(curFontSizePt != null ? String(curFontSizePt) : '')
                    e.target.select()
                  }}
                  onChange={(e) => setSizeDraft(e.target.value)}
                  onBlur={() => setSizeDraft(null)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      e.preventDefault()
                      commitSizeDraft()
                      e.currentTarget.blur()
                    } else if (e.key === 'Escape') {
                      e.preventDefault()
                      e.currentTarget.blur()
                    }
                  }}
                />
                <button
                  className="rb-size-caret"
                  disabled={!editing && !hasTextSelection}
                  data-tip={t('ribbonFontSizeTip')}
                  aria-label={t('ribbonFontSizeTip')}
                  onMouseDown={(e) => {
                    e.preventDefault()
                    e.stopPropagation()
                    if (editing || hasTextSelection) {
                      closeSiblingPanels(e, closePanels, 'size')
                      setSizeOpen((v) => !v)
                    }
                  }}
                >
                  <RbCaret />
                </button>
              </div>
              {sizeOpen && (
                <div
                  className="rb-drop rb-menu rb-menu-scroll rb-size-menu"
                  onMouseDown={(e) => e.stopPropagation()}
                >
                  {FONT_SIZES.map((s) => (
                    <button
                      key={s}
                      className={s === curFontSizePt ? 'on' : ''}
                      onMouseDown={(e) => {
                        e.preventDefault()
                        onFontSize(s)
                        setSizeOpen(false)
                      }}
                    >
                      {s}
                    </button>
                  ))}
                </div>
              )}
            </div>
            {fmtBtn('fontSizeUp', <IconGrowFont size={18} />, t('ribbonGrowFont'))}
            {fmtBtn('fontSizeDown', <IconShrinkFont size={18} />, t('ribbonShrinkFont'))}
            {fmtBtn(
              'removeFormat',
              <IconClearFormat size={18} />,
              t('ribbonClearFormatting'),
              'rb-push-right',
            )}
          </div>
          <div className="rb-row">
            {(
              [
                ['bold', <b key="b">B</b>, t('ribbonBold')],
                ['italic', <i key="i">I</i>, t('ribbonItalic')],
                ['underline', <u key="u">U</u>, t('ribbonUnderline')],
              ] as const
            ).map(([kind, label, title]) => (
              <button
                key={kind}
                className="rb-icon"
                disabled={!editing && !hasSelection}
                data-tip={title}
                aria-label={title}
                onMouseDown={(e) => {
                  e.preventDefault()
                  // While editing change the selection; with only elements selected toggle the whole box
                  if (editing) onFormat(kind)
                  else if (hasSelection) onTextToggle(kind)
                }}
              >
                {label}
              </button>
            ))}
            <button
              className="rb-icon"
              disabled={!editing && !hasSelection}
              data-tip={t('ribbonStrikethrough')}
              onMouseDown={(e) => {
                e.preventDefault()
                if (editing) onFormat('strikeThrough')
                else if (hasSelection) onStrike()
              }}
            >
              <s>ab</s>
            </button>
            {fmtBtn(
              'superscript',
              <span>
                x<sup className="rb-accent">2</sup>
              </span>,
              t('ribbonSuperscript'),
            )}
            {fmtBtn(
              'subscript',
              <span>
                x<sub className="rb-accent">2</sub>
              </span>,
              t('ribbonSubscript'),
            )}
            <span className="rb-mini-sep" />
            <div className="rb-drop-wrap">
              <button
                className="rb-icon"
                disabled={!editing && !hasSelection}
                data-tip={t('ribbonFontColor')}
                onMouseDown={(e) => {
                  e.preventDefault()
                  e.stopPropagation()
                  if (editing || hasSelection) {
                    closeSiblingPanels(e, closePanels, 'color')
                    setColorOpen((v) => !v)
                  }
                }}
              >
                <span className="rb-font-color">
                  A<span className="rb-color-underbar" style={{ background: lastColor }} />
                </span>
                <span className="rb-font-caret">
                  <RbCaret />
                </span>
              </button>
              {colorOpen && (
                <div className="rb-drop rb-color-grid" onMouseDown={(e) => e.stopPropagation()}>
                  {[...TEXT_COLORS, ...recentColors.filter((c) => !TEXT_COLORS.includes(c))].map(
                    (c) => (
                      <button
                        key={c}
                        className="rb-swatch"
                        style={{ background: c }}
                        data-tip={c}
                        aria-label={c}
                        onMouseDown={(e) => {
                          e.preventDefault()
                          setLastColor(c)
                          if (editing) onTextColor(c)
                          else onElementTextColor(c)
                          setColorOpen(false)
                        }}
                      />
                    ),
                  )}
                  {/* Any-color entry: native picker, same as the shape-fill input in the Format pane.
                            data-keep-edit: opening it doesn't commit the text edit */}
                  <label
                    className="rb-color-more"
                    data-keep-edit=""
                    data-tip={t('ribbonMoreColors')}
                    onMouseDown={(e) => e.stopPropagation()}
                  >
                    <input
                      type="color"
                      value={lastColor}
                      onPointerDown={(e) => {
                        armColorInput(e.currentTarget)
                        if (editing) saveEditSelection()
                      }}
                      onChange={(e) => onCustomTextColor(e.target.value)}
                    />
                    {t('ribbonMoreColors')}
                  </label>
                </div>
              )}
            </div>
          </div>
        </div>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupParagraph')}>
        <div className="rb-drop-wrap">
          <button
            className={`rb-big ${paraOpen ? 'active' : ''}`}
            disabled={!hasDoc}
            data-tip={t('ribbonGroupParagraph')}
            data-keep-edit=""
            onMouseDown={(e) => {
              e.stopPropagation()
              closeSiblingPanels(e, closePanels, 'para')
              if (editing) saveEditSelection()
            }}
            onClick={() => setParaOpen((v) => !v)}
          >
            <span className="rb-big-icon">
              <IconAlignLeft size={BIG} />
              <RbCaret />
            </span>
            <span>{t('ribbonGroupParagraph')}</span>
          </button>
          {paraOpen && (
            <div
              className="rb-drop rb-para-drop"
              data-keep-edit=""
              onMouseDown={(e) => e.stopPropagation()}
            >
              <div className="rb-col">
                <div className="rb-row">
                  <button
                    className="rb-icon"
                    disabled={!hasSelection}
                    data-tip={t('ribbonBullets')}
                    aria-label={t('ribbonBullets')}
                    onMouseDown={(e) => {
                      e.preventDefault()
                      if (hasSelection) onParagraphFormat({ bullet: 'char' })
                    }}
                  >
                    <IconBullets size={20} />
                  </button>
                  <button
                    className="rb-icon"
                    disabled={!hasSelection}
                    data-tip={t('ribbonNumbering')}
                    aria-label={t('ribbonNumbering')}
                    onMouseDown={(e) => {
                      e.preventDefault()
                      if (hasSelection) onParagraphFormat({ bullet: 'number' })
                    }}
                  >
                    <IconNumbered size={20} />
                  </button>
                  <button
                    className="rb-icon"
                    disabled={!hasSelection}
                    data-tip={t('ribbonIndentDec')}
                    aria-label={t('ribbonIndentDec')}
                    onMouseDown={(e) => {
                      e.preventDefault()
                      if (hasSelection) onParagraphFormat({ indentDelta: -1 })
                    }}
                  >
                    <IconIndentDec size={20} />
                  </button>
                  <button
                    className="rb-icon"
                    disabled={!hasSelection}
                    data-tip={t('ribbonIndentInc')}
                    aria-label={t('ribbonIndentInc')}
                    onMouseDown={(e) => {
                      e.preventDefault()
                      if (hasSelection) onParagraphFormat({ indentDelta: 1 })
                    }}
                  >
                    <IconIndentInc size={20} />
                  </button>
                </div>
                <div className="rb-para-label">{t('ribbonBulletChar')}</div>
                <div className="rb-bullet-grid">
                  <button
                    className={`rb-bullet-tile rb-bullet-tile-none ${curBulletChar === '' ? 'on' : ''}`}
                    disabled={!hasSelection}
                    data-tip={t('ribbonNone')}
                    onMouseDown={(e) => {
                      e.preventDefault()
                      if (hasSelection) onParagraphFormat({ bullet: 'none' })
                    }}
                  >
                    {t('ribbonNone')}
                  </button>
                  {['•', '○', '▪', '◆', '-', '✓', '►', '※'].map((g) => (
                    <button
                      key={g}
                      className={`rb-bullet-tile ${curBulletChar === g ? 'on' : ''}`}
                      disabled={!hasSelection}
                      data-tip={t('ribbonBulletChar')}
                      onMouseDown={(e) => {
                        e.preventDefault()
                        if (hasSelection) onParagraphFormat({ bullet: 'char', bulletChar: g })
                      }}
                    >
                      {[0, 1, 2].map((i) => (
                        <span key={i} className="rb-bullet-tile-row">
                          <span className="rb-bullet-tile-glyph">{g}</span>
                          <span className="rb-bullet-tile-bar" />
                        </span>
                      ))}
                    </button>
                  ))}
                </div>
                <div className="rb-para-label">{t('ribbonBulletHang')}</div>
                <div className="rb-row">
                  {(
                    [
                      ['ribbonBulletHangNarrow', 114300],
                      ['ribbonBulletHangNormal', 228600],
                      ['ribbonBulletHangWide', 342900],
                    ] as const
                  ).map(([key, emu]) => (
                    <button
                      key={key}
                      className="rb-bullet-hang"
                      disabled={!hasSelection}
                      data-tip={t('ribbonBulletHang')}
                      onMouseDown={(e) => {
                        e.preventDefault()
                        if (hasSelection) onParagraphFormat({ bulletHangEmu: emu })
                      }}
                    >
                      {t(key)}
                    </button>
                  ))}
                  <input
                    className="rb-bullet-hang rb-bullet-hang-input"
                    disabled={!hasSelection}
                    data-tip={t('ribbonBulletHangCustomTip')}
                    placeholder="px"
                    value={hangDraft}
                    onMouseDown={(e) => e.stopPropagation()}
                    onChange={(e) => setHangDraft(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        e.preventDefault()
                        commitHangDraft()
                      } else if (e.key === 'Escape') {
                        e.preventDefault()
                        e.currentTarget.blur()
                      }
                    }}
                  />
                </div>
                <div className="rb-para-label">{t('ribbonBulletSize')}</div>
                <div className="rb-row">
                  {[75, 90, 100, 125, 150].map((pct) => (
                    <button
                      key={pct}
                      className="rb-bullet-hang"
                      disabled={!hasSelection}
                      data-tip={t('ribbonBulletSize')}
                      onMouseDown={(e) => {
                        e.preventDefault()
                        if (hasSelection) onParagraphFormat({ bulletSizePct: pct })
                      }}
                    >
                      {pct}%
                    </button>
                  ))}
                </div>
                <div className="rb-para-label">{t('ribbonBulletColor')}</div>
                <div className="rb-row rb-bullet-colors">
                  {TEXT_COLORS.map((c) => (
                    <button
                      key={c}
                      className="rb-swatch"
                      style={{ background: c }}
                      disabled={!hasSelection}
                      data-tip={c}
                      aria-label={c}
                      onMouseDown={(e) => {
                        e.preventDefault()
                        if (hasSelection) onParagraphFormat({ bulletColor: c })
                      }}
                    />
                  ))}
                  <label
                    className="rb-color-more rb-bullet-color-more"
                    data-tip={t('ribbonMoreColors')}
                    onMouseDown={(e) => e.stopPropagation()}
                  >
                    <input
                      type="color"
                      value={lastBulletColor}
                      onPointerDown={(e) => armColorInput(e.currentTarget)}
                      onChange={(e) => {
                        if (hasSelection) onCustomBulletColor(e.target.value)
                      }}
                    />
                    {t('ribbonMoreColors')}
                  </label>
                </div>
                <div className="rb-row">
                  {(
                    [
                      ['left', <IconAlignLeft key="l" size={20} />, t('ribbonAlignLeft')],
                      ['center', <IconAlignCenter key="c" size={20} />, t('ribbonAlignCenter')],
                      ['right', <IconAlignRight key="r" size={20} />, t('ribbonAlignRight')],
                      ['justify', <IconAlignJustify key="j" size={20} />, t('ribbonAlignJustify')],
                    ] as const
                  ).map(([align, icon, label]) => (
                    <button
                      key={align}
                      className={`rb-icon ${curAlign === align ? 'active' : ''}`}
                      disabled={!editing && !hasSelection}
                      data-tip={label}
                      aria-label={label}
                      onMouseDown={(e) => {
                        e.preventDefault()
                        if (editing || hasSelection) onAlign(align)
                      }}
                    >
                      {icon}
                    </button>
                  ))}
                  <span className="rb-mini-sep" />
                  <div className="rb-drop-wrap">
                    <button
                      className={`rb-icon ${lineSpacingOpen ? 'active' : ''}`}
                      disabled={!hasSelection}
                      data-tip={t('ribbonLineSpacing')}
                      aria-label={t('ribbonLineSpacing')}
                      onMouseDown={(e) => {
                        e.preventDefault()
                        e.stopPropagation()
                        if (hasSelection) {
                          closeSiblingPanels(e, closePanels, 'lineSpacing')
                          setLineSpacingOpen((v) => !v)
                        }
                      }}
                    >
                      <IconLineSpacing size={20} />
                    </button>
                    {lineSpacingOpen && (
                      <div className="rb-drop rb-menu" onMouseDown={(e) => e.stopPropagation()}>
                        {(
                          [
                            [100, '1.0'],
                            [115, '1.15'],
                            [150, '1.5'],
                            [200, '2.0'],
                            [250, '2.5'],
                            [300, '3.0'],
                          ] as const
                        ).map(([pct, label]) => (
                          <button
                            key={pct}
                            onMouseDown={(e) => {
                              e.preventDefault()
                              onParagraphFormat({ lineSpacingPct: pct })
                              setLineSpacingOpen(false)
                            }}
                          >
                            {label}
                          </button>
                        ))}
                        <div className="rb-menu-sep" />
                        {/* free pt values: the op plumbing has carried
                            spaceBeforePt/spaceAfterPt all along */}
                        {(
                          [
                            ['ribbonSpaceBefore', 'spaceBeforePt'],
                            ['ribbonSpaceAfter', 'spaceAfterPt'],
                          ] as const
                        ).map(([labelKey, field]) => (
                          <label key={field} className="rb-menu-input">
                            <span>{t(labelKey)}</span>
                            <input
                              type="text"
                              inputMode="decimal"
                              placeholder="pt"
                              aria-label={t(labelKey)}
                              onMouseDown={() => {
                                if (editing) saveEditSelection()
                              }}
                              onKeyDown={(e) => {
                                if (e.key !== 'Enter') return
                                const pt = Number(e.currentTarget.value.trim().replace(',', '.'))
                                if (!Number.isFinite(pt) || pt < 0 || pt > 500) return
                                onParagraphFormat({ [field]: pt })
                                setLineSpacingOpen(false)
                              }}
                            />
                          </label>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupPanes')}>
        <button
          className={`rb-big ${formatOpen ? 'active' : ''}`}
          disabled={!hasDoc}
          onClick={onToggleFormat}
          data-tip={t('ribbonFormatPaneTip')}
        >
          <span className="rb-big-icon">
            <IconPosition size={BIG} />
          </span>
          <span>{t('ribbonFormatPane')}</span>
        </button>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupArrange')}>
        <div className="rb-drop-wrap">
          <button
            className={`rb-big ${arrangeOpen ? 'active' : ''}`}
            disabled={!hasSelection || !onArrange}
            data-tip={
              hasSelection
                ? t('ribbonAlignMenuTip')
                : t('ribbonSelectFirstHint', { title: t('ribbonAlignMenu') })
            }
            onMouseDown={(e) => {
              e.stopPropagation()
              closeSiblingPanels(e, closePanels, 'arrange')
            }}
            onClick={() => setArrangeOpen((v) => !v)}
          >
            <span className="rb-big-icon">
              <IconObjAlignLeft size={BIG} />
              <RbCaret />
            </span>
            <span>{t('ribbonAlignMenu')}</span>
          </button>
          {arrangeOpen && (
            <div className="rb-drop rb-menu rb-menu-wide" onMouseDown={(e) => e.stopPropagation()}>
              {(
                [
                  ['left', <IconObjAlignLeft key="i" size={20} />, t('ribbonAlignLeft')],
                  ['center-h', <IconObjAlignCenterH key="i" size={20} />, t('ribbonAlignCenterH')],
                  ['right', <IconObjAlignRight key="i" size={20} />, t('ribbonAlignRight')],
                  ['top', <IconObjAlignTop key="i" size={20} />, t('ribbonAlignTop')],
                  ['center-v', <IconObjAlignMiddle key="i" size={20} />, t('ribbonAlignMiddle')],
                  ['bottom', <IconObjAlignBottom key="i" size={20} />, t('ribbonAlignBottom')],
                ] as const
              ).map(([kind, icon, label]) => (
                <button
                  key={kind}
                  onClick={() => {
                    setArrangeOpen(false)
                    onArrange?.(kind)
                  }}
                >
                  <span className="rb-menu-glyph">{icon}</span>
                  {label}
                </button>
              ))}
              <div className="rb-menu-div" />
              {(
                [
                  [
                    'distribute-h',
                    <IconObjDistributeH key="i" size={20} />,
                    t('ribbonDistributeH'),
                  ],
                  [
                    'distribute-v',
                    <IconObjDistributeV key="i" size={20} />,
                    t('ribbonDistributeV'),
                  ],
                ] as const
              ).map(([kind, icon, label]) => (
                <button
                  key={kind}
                  disabled={!canDistribute}
                  data-tip={canDistribute ? undefined : t('ribbonDistributeHint', { title: label })}
                  aria-label={
                    canDistribute ? undefined : t('ribbonDistributeHint', { title: label })
                  }
                  onClick={() => {
                    setArrangeOpen(false)
                    onArrange?.(kind)
                  }}
                >
                  <span className="rb-menu-glyph">{icon}</span>
                  {label}
                </button>
              ))}
              {onFlip && (
                <>
                  <div className="rb-menu-div" />
                  {/* Mirror: rotation can't express a single-axis flip, so arrows could
                      only ever point one way without this */}
                  {(
                    [
                      ['h', t('ribbonFlipH')],
                      ['v', t('ribbonFlipV')],
                    ] as const
                  ).map(([axis, label]) => (
                    <button
                      key={axis}
                      onClick={() => {
                        setArrangeOpen(false)
                        onFlip(axis)
                      }}
                    >
                      <span className="rb-menu-glyph">
                        {axis === 'h' ? <IconObjFlipH size={20} /> : <IconObjFlipV size={20} />}
                      </span>
                      {label}
                    </button>
                  ))}
                </>
              )}
            </div>
          )}
        </div>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupEditing')}>
        <button
          className="rb-big"
          disabled={!hasDoc}
          onClick={onFindReplace}
          data-tip={platformShortcuts('⌘F')}
        >
          <span className="rb-big-icon">
            <IconFind size={BIG} />
          </span>
          <span>{t('ribbonFindReplace')}</span>
        </button>
      </Group>
    </>
  )
}
