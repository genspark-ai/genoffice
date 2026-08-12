import { useState } from 'react'
import { type SectionSettings } from '@genoffice/docx-engine'
import { WRAP_OPTIONS } from './ContextMenu'
import { MarginDialog, cmFromTwips, marginsFitPage, type PageMargins } from './MarginDialog'
import { useI18n, type StringKey } from '../i18n/locale'
import {
  IconCaret,
  IconColumns,
  IconMargins,
  IconOrientation,
  IconPageBreak,
  IconPageSize,
  IconPosition,
  IconWrapText,
} from './icons'

/** icon size for the big icon-over-label ribbon buttons (slides ribbon parity) */
import { BIG, TabProps, activeParaAttrs, setParaAttrs, toggleDropdown } from './ribbon-tabs'

const MARGIN_PRESETS: Array<{
  key: string
  nameKey: StringKey
  descKey: StringKey
  top: number
  right: number
  bottom: number
  left: number
}> = [
  {
    key: 'normal',
    nameKey: 'ribbonMarginNormal',
    descKey: 'ribbonMarginNormalDesc',
    top: 1440,
    right: 1440,
    bottom: 1440,
    left: 1440,
  },
  {
    key: 'narrow',
    nameKey: 'ribbonMarginNarrow',
    descKey: 'ribbonMarginNarrowDesc',
    top: 720,
    right: 720,
    bottom: 720,
    left: 720,
  },
  {
    key: 'moderate',
    nameKey: 'ribbonMarginModerate',
    descKey: 'ribbonMarginModerateDesc',
    top: 1440,
    right: 1080,
    bottom: 1440,
    left: 1080,
  },
  {
    key: 'wide',
    nameKey: 'ribbonMarginWide',
    descKey: 'ribbonMarginWideDesc',
    top: 1440,
    right: 2880,
    bottom: 1440,
    left: 2880,
  },
]

const LAST_CUSTOM_MARGINS_KEY = 'aidocs.marginLastCustom'

function readLastCustomMargins(): PageMargins | null {
  try {
    const raw = localStorage.getItem(LAST_CUSTOM_MARGINS_KEY)
    if (!raw) return null
    const v = JSON.parse(raw) as PageMargins
    const sides = [v.top, v.right, v.bottom, v.left]
    return sides.every((n) => Number.isFinite(n) && n >= 0) ? v : null
  } catch {
    return null
  }
}

const PAPER_SIZES = [
  { key: 'a4', name: 'A4', desc: '21 × 29.7 cm', w: 11906, h: 16838 },
  { key: 'letter', name: 'Letter', desc: '21.59 × 27.94 cm', w: 12240, h: 15840 },
  { key: 'legal', name: 'Legal', desc: '21.59 × 35.56 cm', w: 12240, h: 20160 },
  { key: 'b5', name: 'B5 (JIS)', desc: '18.2 × 25.7 cm', w: 10319, h: 14572 },
]

interface LayoutTabProps extends TabProps {
  section: SectionSettings | null
  onSection: (next: SectionSettings) => void
  /** Multi-section documents: index of the cursor's section (0-based); null for single-section */
  activeSection: number | null
  onInsertSectionBreak: (type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage') => void
}

const PT_PER_TWIP = 1 / 20

export function LayoutTab({
  editor,
  hasDoc,
  dropdown,
  setDropdown,
  section,
  onSection,
  activeSection,
  onInsertSectionBreak,
}: LayoutTabProps) {
  const { t } = useI18n()
  const paraAttrs = activeParaAttrs(editor)
  const enabled = hasDoc && !!section
  const [marginDialog, setMarginDialog] = useState(false)

  const applyMargins = (m: PageMargins) => {
    if (!section || !marginsFitPage(m, section.pageWidth, section.pageHeight)) return
    onSection({
      ...section,
      marginTop: m.top,
      marginRight: m.right,
      marginBottom: m.bottom,
      marginLeft: m.left,
    })
  }

  const marginsActive = (m: PageMargins) =>
    !!section &&
    section.marginTop === m.top &&
    section.marginRight === m.right &&
    section.marginBottom === m.bottom &&
    section.marginLeft === m.left

  // Arrange group: enabled when a floating object (image/textbox) is selected; maps to Word's Position / Wrap Text
  const protAttrs = editor.getAttributes('docProtected')
  const isImage = protAttrs?.blockType === 'image' && !!protAttrs.imageDataUrl
  const isFloatingBox =
    Array.isArray(protAttrs?.textboxes) && (protAttrs.textboxes as unknown[]).length > 0
  const canWrap = hasDoc && (isImage || isFloatingBox)
  // position presets go through imagePatchOf (original-document images only); newly inserted objects work after saving
  const canPosition = hasDoc && isImage && protAttrs?.docxIndex != null
  const currentWrap = (protAttrs?.imageWrap as string | null) ?? null

  const applyWrap = (value: string | null) => {
    const cleared =
      value === null
        ? { imagePosH: null, imagePosV: null, imageOffsetXEmu: null, imageOffsetYEmu: null }
        : {}
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', { imageWrap: value, ...cleared })
      .run()
    setDropdown(() => null)
  }

  const applyPositionPreset = (h: 'left' | 'center' | 'right', v: 'top' | 'center' | 'bottom') => {
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', {
        imageWrap: h === 'right' ? 'square-right' : 'square-left',
        imagePosH: h,
        imagePosV: v,
        imageOffsetXEmu: null,
        imageOffsetYEmu: null,
      })
      .run()
    setDropdown(() => null)
  }

  const applyInlinePosition = () => applyWrap(null)

  const setOrientation = (orientation: 'portrait' | 'landscape') => {
    if (!section || section.orientation === orientation) return
    onSection({
      ...section,
      orientation,
      pageWidth: section.pageHeight,
      pageHeight: section.pageWidth,
    })
    setDropdown(() => null)
  }

  const setPaper = (w: number, h: number) => {
    if (!section) return
    const landscape = section.orientation === 'landscape'
    onSection({ ...section, pageWidth: landscape ? h : w, pageHeight: landscape ? w : h })
    setDropdown(() => null)
  }

  const ptInput = (attr: string, title: string) => {
    const twips = Number(paraAttrs[attr]) || 0
    return (
      <label className="layout-num" title={title}>
        <span>{title}</span>
        <input
          type="number"
          min={0}
          max={400}
          step={1}
          disabled={!hasDoc}
          value={Math.round(twips * PT_PER_TWIP)}
          onChange={(e) => {
            const pt = Math.max(0, Number(e.target.value) || 0)
            setParaAttrs(editor, { [attr]: pt > 0 ? Math.round(pt / PT_PER_TWIP) : null })
          }}
        />
        <span className="layout-unit">pt</span>
      </label>
    )
  }

  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!enabled}
              title={t('ribbonMargins')}
              onClick={() => toggleDropdown(setDropdown, 'margins')}
            >
              <span className="rb-big-icon">
                <IconMargins size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonMargins')}</span>
            </button>
            {dropdown === 'margins' && section && (
              <div className="layout-menu">
                {(() => {
                  const storedLastCustom = readLastCustomMargins()
                  const lastCustom =
                    storedLastCustom &&
                    marginsFitPage(storedLastCustom, section.pageWidth, section.pageHeight)
                      ? storedLastCustom
                      : null
                  return (
                    <>
                      {lastCustom && (
                        <button
                          className={marginsActive(lastCustom) ? 'active' : ''}
                          onClick={() => {
                            applyMargins(lastCustom)
                            setDropdown(() => null)
                          }}
                        >
                          <b>{t('ribbonMarginLastCustom')}</b>
                          <span>
                            {t('ribbonMarginTop')} {cmFromTwips(lastCustom.top)} ·{' '}
                            {t('ribbonMarginBottom')} {cmFromTwips(lastCustom.bottom)} ·{' '}
                            {t('ribbonMarginLeft')} {cmFromTwips(lastCustom.left)} ·{' '}
                            {t('ribbonMarginRight')} {cmFromTwips(lastCustom.right)} cm
                          </span>
                        </button>
                      )}
                      {MARGIN_PRESETS.map((m) => (
                        <button
                          key={m.key}
                          className={marginsActive(m) ? 'active' : ''}
                          disabled={!marginsFitPage(m, section.pageWidth, section.pageHeight)}
                          onClick={() => {
                            applyMargins(m)
                            setDropdown(() => null)
                          }}
                        >
                          <b>{t(m.nameKey)}</b>
                          <span>{t(m.descKey)}</span>
                        </button>
                      ))}
                      <button
                        onClick={() => {
                          setDropdown(() => null)
                          setMarginDialog(true)
                        }}
                      >
                        <b>{t('ribbonMarginCustom')}</b>
                      </button>
                    </>
                  )
                })()}
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!enabled}
              title={t('ribbonOrientation')}
              onClick={() => toggleDropdown(setDropdown, 'orient')}
            >
              <span className="rb-big-icon">
                <IconOrientation size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonOrientation')}</span>
            </button>
            {dropdown === 'orient' && section && (
              <div className="layout-menu">
                <button
                  className={section.orientation === 'portrait' ? 'active' : ''}
                  onClick={() => setOrientation('portrait')}
                >
                  <b>{t('ribbonPortrait')}</b>
                </button>
                <button
                  className={section.orientation === 'landscape' ? 'active' : ''}
                  onClick={() => setOrientation('landscape')}
                >
                  <b>{t('ribbonLandscape')}</b>
                </button>
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!enabled}
              title={t('ribbonPaperSize')}
              onClick={() => toggleDropdown(setDropdown, 'paper')}
            >
              <span className="rb-big-icon">
                <IconPageSize size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonPaperSize')}</span>
            </button>
            {dropdown === 'paper' && section && (
              <div className="layout-menu">
                {PAPER_SIZES.map((p) => {
                  const portraitW = Math.min(section.pageWidth, section.pageHeight)
                  return (
                    <button
                      key={p.key}
                      className={portraitW === p.w ? 'active' : ''}
                      onClick={() => setPaper(p.w, p.h)}
                    >
                      <b>{p.name}</b>
                      <span>{p.desc}</span>
                    </button>
                  )
                })}
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className={`rb-big ${section && section.columns > 1 ? 'active' : ''}`}
              disabled={!enabled}
              title={t('ribbonColumns')}
              onClick={() => toggleDropdown(setDropdown, 'columns')}
            >
              <span className="rb-big-icon">
                <IconColumns size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonColumns')}</span>
            </button>
            {dropdown === 'columns' && section && (
              <div className="layout-menu">
                {[1, 2, 3].map((n) => (
                  <button
                    key={n}
                    className={section.columns === n ? 'active' : ''}
                    onClick={() => {
                      onSection({ ...section, columns: n })
                      setDropdown(() => null)
                    }}
                  >
                    {n === 1
                      ? t('ribbonOneColumn')
                      : n === 2
                        ? t('ribbonTwoColumns')
                        : t('ribbonThreeColumns')}
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!enabled}
              title={t('ribbonSectionBreakTip')}
              onClick={() => toggleDropdown(setDropdown, 'sectbreak')}
            >
              <span className="rb-big-icon">
                <IconPageBreak size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonSectionBreak')}</span>
            </button>
            {dropdown === 'sectbreak' && (
              <div className="layout-menu">
                <button
                  onClick={() => {
                    onInsertSectionBreak('nextPage')
                    setDropdown(() => null)
                  }}
                >
                  <b>{t('ribbonBreakNextPage')}</b>
                  <span>{t('ribbonBreakNextPageDesc')}</span>
                </button>
                <button
                  onClick={() => {
                    onInsertSectionBreak('continuous')
                    setDropdown(() => null)
                  }}
                >
                  <b>{t('ribbonBreakContinuous')}</b>
                  <span>{t('ribbonBreakContinuousDesc')}</span>
                </button>
                <button
                  onClick={() => {
                    onInsertSectionBreak('evenPage')
                    setDropdown(() => null)
                  }}
                >
                  <b>{t('ribbonBreakEvenPage')}</b>
                  <span>{t('ribbonBreakEvenPageDesc')}</span>
                </button>
                <button
                  onClick={() => {
                    onInsertSectionBreak('oddPage')
                    setDropdown(() => null)
                  }}
                >
                  <b>{t('ribbonBreakOddPage')}</b>
                  <span>{t('ribbonBreakOddPageDesc')}</span>
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">
          {activeSection !== null
            ? t('ribbonGroupPageSetupSection', { n: activeSection + 1 })
            : t('ribbonGroupPageSetup')}
        </div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items layout-para">
          <div className="layout-col">
            {ptInput('indentLeft', t('ribbonIndentLeft'))}
            {ptInput('indentRight', t('ribbonIndentRight'))}
          </div>
          <div className="layout-col">
            {ptInput('spaceBefore', t('ribbonSpaceBefore'))}
            {ptInput('spaceAfter', t('ribbonSpaceAfter'))}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupParagraphSelected')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!canPosition}
              title={t('ribbonPosition')}
              onClick={() => toggleDropdown(setDropdown, 'arrange-pos')}
            >
              <span className="rb-big-icon">
                <IconPosition size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonPosition')}</span>
            </button>
            {dropdown === 'arrange-pos' && (
              <div className="layout-menu">
                <button
                  className={currentWrap ? '' : 'active'}
                  onClick={() => applyInlinePosition()}
                >
                  <b>{t('appWrapInline')}</b>
                </button>
                <div className="pos-grid">
                  {(['top', 'center', 'bottom'] as const).map((v) =>
                    (['left', 'center', 'right'] as const).map((h) => (
                      <button
                        key={`${v}-${h}`}
                        className={`pos-cell ph-${h} pv-${v}${
                          protAttrs?.imagePosH === h && protAttrs?.imagePosV === v ? ' active' : ''
                        }`}
                        title={t('ribbonPosition')}
                        onClick={() => applyPositionPreset(h, v)}
                      >
                        <span className="pos-dot" />
                      </button>
                    )),
                  )}
                </div>
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!canWrap}
              title={t('ribbonWrapText')}
              onClick={() => toggleDropdown(setDropdown, 'arrange-wrap')}
            >
              <span className="rb-big-icon">
                <IconWrapText size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonWrapText')}</span>
            </button>
            {dropdown === 'arrange-wrap' && (
              <div className="layout-menu">
                {WRAP_OPTIONS.map((opt) => (
                  <button
                    key={String(opt.value)}
                    className={currentWrap === opt.value ? 'active' : ''}
                    onClick={() => applyWrap(opt.value)}
                  >
                    <b>{t(opt.labelKey)}</b>
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupArrange')}</div>
      </div>

      {marginDialog && section && (
        <MarginDialog
          margins={{
            top: section.marginTop,
            right: section.marginRight,
            bottom: section.marginBottom,
            left: section.marginLeft,
          }}
          pageWidth={section.pageWidth}
          pageHeight={section.pageHeight}
          onApply={(m) => {
            applyMargins(m)
            localStorage.setItem(LAST_CUSTOM_MARGINS_KEY, JSON.stringify(m))
          }}
          onClose={() => setMarginDialog(false)}
        />
      )}
    </>
  )
}

/* ================= References ================= */
