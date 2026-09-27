import { useRef, useState } from 'react'
import { type SectionSettings } from '@genoffice/docx-engine'
import { WRAP_OPTIONS } from './ContextMenu'
import { MarginDialog, marginsFitPage, type PageMargins } from './MarginDialog'
import { LengthInput } from './LengthInput'
import { PaperSizeDialog } from './PaperSizeDialog'
import { insertColumnBreak, insertTextWrappingBreak } from '../editor/page-break'
import { useI18n, type StringKey } from '../i18n/locale'
import { useMeasurement } from '../use-measurement'
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
import {
  BIG,
  TabProps,
  activeParaAttrs,
  insertPageBreakAt,
  setParaAttrs,
  toggleDropdown,
} from './ribbon-tabs'

/** Word's Margins gallery order; Mirrored's left/right are inside/outside and switch on w:mirrorMargins */
export const MARGIN_PRESETS: Array<{
  key: string
  nameKey: StringKey
  top: number
  right: number
  bottom: number
  left: number
  mirror?: true
}> = [
  {
    key: 'normal',
    nameKey: 'ribbonMarginNormal',
    top: 1440,
    right: 1440,
    bottom: 1440,
    left: 1440,
  },
  {
    key: 'narrow',
    nameKey: 'ribbonMarginNarrow',
    top: 720,
    right: 720,
    bottom: 720,
    left: 720,
  },
  {
    key: 'moderate',
    nameKey: 'ribbonMarginModerate',
    top: 1440,
    right: 1080,
    bottom: 1440,
    left: 1080,
  },
  {
    key: 'wide',
    nameKey: 'ribbonMarginWide',
    top: 1440,
    right: 2880,
    bottom: 1440,
    left: 2880,
  },
  {
    key: 'mirrored',
    nameKey: 'ribbonMarginMirrored',
    top: 1440,
    right: 1440,
    bottom: 1440,
    left: 1800,
    mirror: true,
  },
  {
    key: 'office2003',
    nameKey: 'ribbonMarginOffice2003',
    top: 1440,
    right: 1800,
    bottom: 1440,
    left: 1800,
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

/** Word's Size gallery, in Word's order; portrait twips */
export const PAPER_SIZES: ReadonlyArray<{ key: string; name: string; w: number; h: number }> = [
  { key: 'letter', name: 'Letter', w: 12240, h: 15840 },
  { key: 'legal', name: 'Legal', w: 12240, h: 20160 },
  { key: 'executive', name: 'Executive', w: 10440, h: 15120 },
  { key: 'a3', name: 'A3', w: 16838, h: 23811 },
  { key: 'a4', name: 'A4', w: 11906, h: 16838 },
  { key: 'a5', name: 'A5', w: 8391, h: 11906 },
  { key: 'b4', name: 'B4 (JIS)', w: 14572, h: 20639 },
  { key: 'b5', name: 'B5 (JIS)', w: 10319, h: 14572 },
  { key: 'tabloid', name: 'Tabloid', w: 15840, h: 24480 },
  { key: 'statement', name: 'Statement', w: 7920, h: 12240 },
  { key: 'env10', name: 'Envelope #10', w: 5940, h: 13680 },
  { key: 'envdl', name: 'Envelope DL', w: 6237, h: 12474 },
  { key: 'envc5', name: 'Envelope C5', w: 9185, h: 12983 },
]

export function paperSizeCaption(
  wTwips: number,
  hTwips: number,
  format: (twips: number) => string,
): string {
  return `${format(wTwips)} × ${format(hTwips)}`
}

/** the gallery entry a page matches in either orientation (Word tolerates 1-2 twips of rounding) */
export function paperSizeKeyOf(pageWidth: number, pageHeight: number): string | null {
  const w = Math.min(pageWidth, pageHeight)
  const h = Math.max(pageWidth, pageHeight)
  return PAPER_SIZES.find((p) => Math.abs(p.w - w) <= 2 && Math.abs(p.h - h) <= 2)?.key ?? null
}

interface LayoutTabProps extends TabProps {
  section: SectionSettings | null
  onSection: (next: SectionSettings) => void
  /** Multi-section documents: index of the cursor's section (0-based); null for single-section */
  activeSection: number | null
  onInsertSectionBreak: (type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage') => void
  onPaperSizeAll: (
    portraitW: number,
    portraitH: number,
    orientation: SectionSettings['orientation'],
  ) => void
  mirrorMargins: boolean
  onMirrorMargins: (on: boolean) => void
}

export function LayoutTab({
  editor,
  hasDoc,
  dropdown,
  setDropdown,
  section,
  onSection,
  activeSection,
  onInsertSectionBreak,
  onPaperSizeAll,
  mirrorMargins,
  onMirrorMargins,
}: LayoutTabProps) {
  const { t } = useI18n()
  const { format } = useMeasurement()
  const paraAttrs = activeParaAttrs(editor)
  const enabled = hasDoc && !!section
  const [marginDialog, setMarginDialog] = useState(false)
  const [paperDialog, setPaperDialog] = useState(false)
  // Word's gallery captions; a mirrored preset reads inside/outside instead of left/right
  const marginCaption = (m: PageMargins & { mirror?: true }) =>
    `${t('ribbonMarginTop')} ${format(m.top)} · ${t('ribbonMarginBottom')} ${format(m.bottom)} · ${t(m.mirror ? 'ribbonMarginInside' : 'ribbonMarginLeft')} ${format(m.left)} · ${t(m.mirror ? 'ribbonMarginOutside' : 'ribbonMarginRight')} ${format(m.right)}`

  const applyMargins = (m: PageMargins, mirror?: boolean) => {
    if (!section || !marginsFitPage(m, section.pageWidth, section.pageHeight)) return
    // a user-set value is an ordinary margin, not the file's header-proof fixed one
    onSection({
      ...section,
      marginTop: m.top,
      marginRight: m.right,
      marginBottom: m.bottom,
      marginLeft: m.left,
      marginTopFixed: undefined,
      marginBottomFixed: undefined,
    })
    if (mirror !== undefined) onMirrorMargins(mirror)
  }

  const marginsActive = (m: PageMargins & { mirror?: true }) =>
    !!section &&
    section.marginTop === m.top &&
    section.marginRight === m.right &&
    section.marginBottom === m.bottom &&
    section.marginLeft === m.left &&
    !!m.mirror === mirrorMargins

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

  /** selection captured when a field takes focus (commit target on blur) */
  const inputTargetRef = useRef<{ from: number; to: number } | null>(null)
  const commitPara = (attr: string, twips: number | null) => {
    const target = inputTargetRef.current
    setParaAttrs(editor, { [attr]: twips ? twips : null }, target ?? undefined)
  }
  // the paragraphs the entry is FOR: by blur, a click may already have
  // moved the live selection elsewhere
  const captureTarget = () => {
    const { from, to } = editor.state.selection
    inputTargetRef.current = { from, to }
  }
  // Enter-blur (no relatedTarget) hands focus back to the editor; a blur
  // INTO another control must not steal it back
  const releaseTarget = (e: { relatedTarget: EventTarget | null }) => {
    inputTargetRef.current = null
    if (!e.relatedTarget) editor.commands.focus()
  }
  // selection identity in the key: a commit that leaves the LIVE selection's
  // value unchanged must still remount the field, or it keeps showing the
  // number just applied to a different paragraph. The value itself stays out
  // of the length field's key: an arrow-key step commits without remounting.
  const fieldKey = (attr: string, shown?: number) =>
    `${attr}:${shown ?? ''}:${hasDoc}:${editor.state.selection.from}`
  /** Word shows indents in the measurement unit and permits negative ones
   *  (text into the margin); the ribbon bounds them by the page width. */
  const indentInput = (attr: string, title: string) => {
    const twips = Number(paraAttrs[attr]) || 0
    const bound = section?.pageWidth ?? 12240
    return (
      <label className="layout-num" data-tip={title}>
        <span>{title}</span>
        <LengthInput
          key={fieldKey(attr)}
          value={twips}
          min={-bound}
          max={bound}
          disabled={!hasDoc}
          ariaLabel={title}
          onCommit={(next) => commitPara(attr, next)}
          onEnter={() => (document.activeElement as HTMLElement | null)?.blur()}
          onFocus={captureTarget}
          onBlur={releaseTarget}
        />
      </label>
    )
  }
  /** Spacing stays in points, as in Word. Commit on Enter/blur (not per
   *  keystroke) so a half-typed value is not clamped mid-entry. */
  const ptInput = (attr: string, title: string) => {
    const twips = Number(paraAttrs[attr]) || 0
    const shown = Math.round(twips / 20)
    return (
      <label className="layout-num" data-tip={title}>
        <span>{title}</span>
        <input
          type="number"
          min={0}
          max={400}
          step={1}
          disabled={!hasDoc}
          key={fieldKey(attr, shown)}
          defaultValue={shown}
          onKeyDown={(e) => {
            if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
          }}
          onFocus={captureTarget}
          onBlur={(e) => {
            const raw = e.target.value.trim()
            const pt = raw === '' ? 0 : Math.min(400, Math.max(0, Number(raw)))
            if (Number.isFinite(pt)) {
              // clamped input keeps the typed text (uncontrolled): sync it back
              if (String(pt) !== raw) e.target.value = String(pt)
              if (Math.round(pt * 20) !== twips) commitPara(attr, Math.round(pt * 20))
            }
            releaseTarget(e)
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
              data-tip={t('ribbonMargins')}
              data-tip-detail={
                activeSection !== null
                  ? t('ribbonPageSetupSectionTip', { n: activeSection + 1 })
                  : undefined
              }
              onClick={() => toggleDropdown(setDropdown, 'margins')}
            >
              <span className="rb-big-icon">
                <IconMargins size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonMargins')}</span>
            </button>
            {dropdown === 'margins' && section && (
              <div data-rb-panel="" className="layout-menu">
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
                          <span>{marginCaption(lastCustom)}</span>
                        </button>
                      )}
                      {MARGIN_PRESETS.map((m) => (
                        <button
                          key={m.key}
                          className={marginsActive(m) ? 'active' : ''}
                          disabled={!marginsFitPage(m, section.pageWidth, section.pageHeight)}
                          onClick={() => {
                            applyMargins(m, !!m.mirror)
                            setDropdown(() => null)
                          }}
                        >
                          <b>{t(m.nameKey)}</b>
                          <span>{marginCaption(m)}</span>
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
              data-tip={t('ribbonOrientation')}
              onClick={() => toggleDropdown(setDropdown, 'orient')}
            >
              <span className="rb-big-icon">
                <IconOrientation size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonOrientation')}</span>
            </button>
            {dropdown === 'orient' && section && (
              <div data-rb-panel="" className="layout-menu">
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
              data-tip={t('ribbonPaperSize')}
              onClick={() => toggleDropdown(setDropdown, 'paper')}
            >
              <span className="rb-big-icon">
                <IconPageSize size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonPaperSize')}</span>
            </button>
            {dropdown === 'paper' && section && (
              <div data-rb-panel="" className="layout-menu layout-menu-scroll">
                {(() => {
                  const activeKey = paperSizeKeyOf(section.pageWidth, section.pageHeight)
                  return PAPER_SIZES.map((p) => (
                    <button
                      key={p.key}
                      className={activeKey === p.key ? 'active' : ''}
                      onClick={() => setPaper(p.w, p.h)}
                    >
                      <b>{p.name}</b>
                      <span>{paperSizeCaption(p.w, p.h, format)}</span>
                    </button>
                  ))
                })()}
                <button
                  onClick={() => {
                    setDropdown(() => null)
                    setPaperDialog(true)
                  }}
                >
                  <b>{t('ribbonMorePaperSizes')}</b>
                </button>
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className={`rb-big ${section && section.columns > 1 ? 'active' : ''}`}
              disabled={!enabled}
              data-tip={t('ribbonColumns')}
              onClick={() => toggleDropdown(setDropdown, 'columns')}
            >
              <span className="rb-big-icon">
                <IconColumns size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonColumns')}</span>
            </button>
            {dropdown === 'columns' && section && (
              <div data-rb-panel="" className="layout-menu">
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
              data-tip={t('ribbonBreaksTip')}
              onClick={() => toggleDropdown(setDropdown, 'sectbreak')}
            >
              <span className="rb-big-icon">
                <IconPageBreak size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonBreaks')}</span>
            </button>
            {dropdown === 'sectbreak' && (
              <div data-rb-panel="" className="layout-menu">
                <div className="layout-menu-head">{t('ribbonBreaksPageGroup')}</div>
                <button
                  onClick={() => {
                    insertPageBreakAt(editor)
                    setDropdown(() => null)
                  }}
                >
                  <b>{t('ribbonBreakPage')}</b>
                  <span>{t('ribbonBreakPageDesc')}</span>
                </button>
                <button
                  onClick={() => {
                    insertColumnBreak(editor)
                    setDropdown(() => null)
                  }}
                >
                  <b>{t('ribbonBreakColumn')}</b>
                  <span>{t('ribbonBreakColumnDesc')}</span>
                </button>
                <button
                  onClick={() => {
                    insertTextWrappingBreak(editor)
                    setDropdown(() => null)
                  }}
                >
                  <b>{t('ribbonBreakTextWrapping')}</b>
                  <span>{t('ribbonBreakTextWrappingDesc')}</span>
                </button>
                <div className="layout-menu-head">{t('ribbonSectionBreak')}</div>
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
        <div className="ribbon-group-label">{t('ribbonGroupPageSetup')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items layout-para">
          <div className="layout-col">
            {indentInput('indentLeft', t('ribbonIndentLeft'))}
            {indentInput('indentRight', t('ribbonIndentRight'))}
          </div>
          <div className="layout-col">
            {ptInput('spaceBefore', t('ribbonSpaceBefore'))}
            {ptInput('spaceAfter', t('ribbonSpaceAfter'))}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupParagraph')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!canPosition}
              data-tip={t('ribbonPosition')}
              onClick={() => toggleDropdown(setDropdown, 'arrange-pos')}
            >
              <span className="rb-big-icon">
                <IconPosition size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonPosition')}</span>
            </button>
            {dropdown === 'arrange-pos' && (
              <div data-rb-panel="" className="layout-menu">
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
                        data-tip={t('ribbonPosition')}
                        aria-label={t('ribbonPosition')}
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
              data-tip={t('ribbonWrapText')}
              onClick={() => toggleDropdown(setDropdown, 'arrange-wrap')}
            >
              <span className="rb-big-icon">
                <IconWrapText size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonWrapText')}</span>
            </button>
            {dropdown === 'arrange-wrap' && (
              <div data-rb-panel="" className="layout-menu">
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
          mirror={mirrorMargins}
          pageWidth={section.pageWidth}
          pageHeight={section.pageHeight}
          onApply={(m, mirror) => {
            applyMargins(m, mirror)
            localStorage.setItem(LAST_CUSTOM_MARGINS_KEY, JSON.stringify(m))
          }}
          onClose={() => setMarginDialog(false)}
        />
      )}
      {paperDialog && section && (
        <PaperSizeDialog
          section={section}
          multiSection={activeSection !== null}
          onApply={(w, h, scope) => {
            // typed dimensions decide the orientation flag (Word flips the radio); a square keeps it
            const orientation = w === h ? section.orientation : w > h ? 'landscape' : 'portrait'
            if (scope === 'document') {
              onPaperSizeAll(Math.min(w, h), Math.max(w, h), orientation)
            } else {
              onSection({ ...section, pageWidth: w, pageHeight: h, orientation })
            }
          }}
          onClose={() => setPaperDialog(false)}
        />
      )}
    </>
  )
}

/* ================= References ================= */
