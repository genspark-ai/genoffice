import { Fragment, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { Block, CustomNumberingLevel, DocDefaults, ThemeFonts } from '@genoffice/docx-engine'
import { HIGHLIGHT_CSS } from '../editor/extensions'
import { fontFamiliesFor, isEastAsianFontName } from '../font-list'
import { cssFontFamily } from '../line-metrics'
import { useI18n } from '../i18n/locale'
import type { RibbonFormatState } from './ribbon-format-state'
import { findNumIdOfKind } from './Ribbon'
import {
  IconAlignCenter,
  IconAlignJustify,
  IconAlignLeft,
  IconAlignRight,
  IconBullets,
  IconCaret,
  IconClearFormat,
  IconComment,
  IconHighlight,
  IconIndentDec,
  IconIndentInc,
  IconLineSpacing,
  IconLink,
  IconNumbered,
  IconPicture,
} from './icons'

/** Google Docs' style-gallery labels; genoffice only has Normal/H1/H2/H3 paragraph
 *  styles (no Title/Subtitle/Heading 4-6 definitions), so the dropdown only offers
 *  what actually exists. */
const STYLES = [
  { key: 'p', label: 'Normal text' },
  { key: 'h1', label: 'Heading 1' },
  { key: 'h2', label: 'Heading 2' },
  { key: 'h3', label: 'Heading 3' },
] as const

const TEXT_COLORS = [
  '000000',
  '434343',
  '666666',
  '999999',
  'CCCCCC',
  'FFFFFF',
  'C00000',
  'E06666',
  'FF9900',
  'FFD966',
  '38761D',
  '6AA84F',
  '1155CC',
  '4472C4',
  '674EA7',
  '741B47',
]

function IconPrint({ size = 18 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" aria-hidden>
      <path
        d="M6 9V3h12v6M6 18H4a1 1 0 0 1-1-1v-6a1 1 0 0 1 1-1h16a1 1 0 0 1 1 1v6a1 1 0 0 1-1 1h-2M6 14h12v7H6z"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function IconUndoM({ size = 18 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" aria-hidden>
      <path
        d="M8 7 4 11l4 4M4 11h10a6 6 0 1 1 0 12h-1"
        stroke="currentColor"
        strokeWidth="1.7"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function IconMoreVert({ size = 18 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="currentColor" aria-hidden>
      <circle cx="12" cy="5" r="1.6" />
      <circle cx="12" cy="12" r="1.6" />
      <circle cx="12" cy="19" r="1.6" />
    </svg>
  )
}

function IconRedoM({ size = 18 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" aria-hidden>
      <path
        d="M16 7l4 4-4 4M20 11H10a6 6 0 1 0 0 12h1"
        stroke="currentColor"
        strokeWidth="1.7"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

interface GDocsToolbarProps {
  editor: Editor
  formatState: RibbonFormatState
  hasDoc: boolean
  blocks: Block[]
  allocateNumId?: (kind: 'bullet' | 'ordered') => string | null
  createListDef?: (levels: CustomNumberingLevel[]) => string | null
  docDefaults?: DocDefaults
  themeFonts?: ThemeFonts | null
  zoom: number
  onZoom: (zoom: number) => void
  onZoomFit: (mode: 'width' | 'page') => void
  onLink: () => void
  onInsertImage: () => void
  onNewComment: () => void
  canComment: boolean
}

const ZOOM_PRESETS = [50, 75, 90, 100, 125, 150, 200]
const LINE_SPACINGS = [1, 1.15, 1.5, 2]

/** Priority order for adaptive collapse, highest priority first (kept visible
 *  longest, matching Google Docs' own collapse ordering). When the pill can't
 *  fit every group at the current width, trailing (lowest-priority) groups
 *  fold into the "More" (⋮) menu, same as Google Docs does. Clear-formatting
 *  is bundled with the text-color/highlight group (not called out separately
 *  in the spec) since it lives in the same "text decoration" cluster. */
const GROUP_IDS = [
  'undoRedo',
  'biu',
  'font',
  'fontSize',
  'styles',
  'colorHighlight',
  'lists',
  'align',
  'linkImgComment',
  'zoom',
  'print',
] as const

// Reserved width (px) for the "⋮ More" button + its separator when at least
// one group is collapsed — kept out of the fit calculation so the button
// itself never gets crowded out.
const MORE_BUTTON_RESERVE = 70
// Round measured widths to this grid before comparing to the last processed
// width so sub-pixel resize churn can't flip groups in/out (no jitter).
const QUANTIZE = 8
// Small safety margin subtracted from the available width so a group never
// paints flush against (or past) the pill's rounded edge.
const SAFETY_BUFFER = 6

export function GDocsToolbar({
  editor,
  formatState: fs,
  hasDoc,
  blocks,
  allocateNumId,
  docDefaults,
  themeFonts,
  zoom,
  onZoom,
  onZoomFit,
  onLink,
  onInsertImage,
  onNewComment,
  canComment,
}: GDocsToolbarProps) {
  const [dropdown, setDropdown] = useState<string | null>(null)
  // Independent from `dropdown`: whether the collapsed "⋮ More" panel is
  // open. Kept separate so opening a nested dropdown INSIDE the More panel
  // (zoom presets, font list, styles, align, spacing, color, highlight —
  // anything that sets `dropdown` to its own key) doesn't unmount the panel
  // out from under it. Previously the panel's visibility was gated on
  // `dropdown === 'more'`, so any nested dropdown click inside it closed the
  // whole panel, including the submenu that had just opened.
  const [moreOpen, setMoreOpen] = useState(false)
  const ref = useRef<HTMLDivElement>(null)
  const sub = fs.sub
  const ed = sub ?? editor
  const canEdit = fs.editable && hasDoc

  useEffect(() => {
    if (!dropdown && !moreOpen) return
    const onDoc = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) {
        setDropdown(null)
        setMoreOpen(false)
      }
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setDropdown(null)
        setMoreOpen(false)
      }
    }
    window.addEventListener('mousedown', onDoc)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousedown', onDoc)
      window.removeEventListener('keydown', onKey)
    }
  }, [dropdown, moreOpen])

  const chain = () => (canEdit ? ed.chain().focus() : null)

  const setTextStyle = (patch: Record<string, unknown>) => {
    chain()?.setMark('docTextStyle', patch).run()
    setDropdown(null)
  }

  const setFont = (name: string | null) => {
    if (!name) setTextStyle({ font: null, fontAscii: null })
    else if (isEastAsianFontName(name)) setTextStyle({ font: name })
    else setTextStyle({ fontAscii: name })
  }

  const setParaAttr = (attrs: Record<string, unknown>) => {
    if (!canEdit) return
    if (sub) {
      ed.chain().focus().updateAttributes('docParagraph', attrs).run()
      setDropdown(null)
      return
    }
    editor
      .chain()
      .focus()
      .updateAttributes('docParagraph', attrs)
      .updateAttributes('docHeading', attrs)
      .updateAttributes('docListItem', attrs)
      .run()
    setDropdown(null)
  }

  const applyStyle = (key: string) => {
    if (sub || !canEdit) return
    if (key === 'p') editor.chain().focus().setNode('docParagraph').run()
    else
      editor
        .chain()
        .focus()
        .setNode('docHeading', { level: Number(key.slice(1)) })
        .run()
    setDropdown(null)
  }

  const toggleList = (kind: 'bullet' | 'ordered') => {
    if (sub || !canEdit) return
    if (editor.isActive('docListItem', { kind })) {
      editor.chain().focus().setNode('docParagraph').run()
      return
    }
    const numId = findNumIdOfKind(blocks, kind) ?? allocateNumId?.(kind) ?? null
    editor.chain().focus().setNode('docListItem', { kind, numId, ilvl: 0 }).run()
  }

  const changeIndent = (dir: 1 | -1) => {
    if (sub || !canEdit) return
    const cur = fs.indentLeft ?? 0
    const step = 720 // 0.5"
    setParaAttr({ indentLeft: Math.max(0, cur + dir * step) })
  }

  const stepFontSize = (dir: 1 | -1) => {
    if (!canEdit) return
    const next = Math.max(1, Math.round(fs.fontSizePt) + dir)
    setTextStyle({ sizeHalfPoints: next * 2 })
  }

  const { lang } = useI18n()
  const bodyFontName = docDefaults?.asciiFont?.trim() || themeFonts?.minor?.trim() || 'Calibri'
  const fontFamilies = fontFamiliesFor(lang)
  const currentFont = fs.fontFamily || ''
  const activeStyleKey = fs.headingLevel ? `h${fs.headingLevel}` : 'p'
  const activeAlign = fs.align ?? (fs.bidi ? 'right' : 'left')

  const ALIGN_ICONS: Record<string, typeof IconAlignLeft> = {
    left: IconAlignLeft,
    center: IconAlignCenter,
    right: IconAlignRight,
    justify: IconAlignJustify,
  }
  const AlignIcon = ALIGN_ICONS[activeAlign] ?? IconAlignLeft

  const btn = (
    key: string,
    icon: React.ReactNode,
    title: string,
    onClick: () => void,
    active = false,
    disabled = false,
  ) => (
    <button
      key={key}
      type="button"
      className={`gdh-tb-btn ${active ? 'active' : ''}`}
      title={title}
      disabled={disabled || !canEdit}
      onClick={onClick}
    >
      {icon}
    </button>
  )

  // ---- Responsive collapse -------------------------------------------------
  // Groups are measured off-screen (identical markup, same pill classes) so
  // real gaps/padding are accounted for automatically. visibleCount groups
  // (in priority order, see GROUP_IDS) render in the pill; the rest render
  // inside the "More" panel with the exact same handlers — nothing is a
  // disabled clone.
  const [visibleCount, setVisibleCount] = useState<number>(GROUP_IDS.length)
  const pillRef = useRef<HTMLDivElement>(null)
  const groupRefs = useRef<Array<HTMLDivElement | null>>([])
  const lastQuantizedWidth = useRef<number>(-1)

  const recalcCollapse = () => {
    const pill = pillRef.current
    if (!pill) return
    const total = pill.clientWidth
    if (total === 0) return
    const quantized = Math.floor(total / QUANTIZE) * QUANTIZE
    if (quantized === lastQuantizedWidth.current) return
    lastQuantizedWidth.current = quantized
    const available = total - SAFETY_BUFFER
    let n = GROUP_IDS.length
    for (; n >= 1; n--) {
      const el = groupRefs.current[n - 1]
      if (!el) continue
      const collapsedReserve = n < GROUP_IDS.length ? MORE_BUTTON_RESERVE : 0
      const required =
        el.offsetLeft + el.offsetWidth + 18 /* pill right padding */ + collapsedReserve
      if (required <= available) break
    }
    setVisibleCount(Math.max(1, n))
  }

  useLayoutEffect(() => {
    recalcCollapse()
    const pill = pillRef.current
    if (!pill) return
    const ro = new ResizeObserver(() => recalcCollapse())
    ro.observe(pill)
    return () => ro.disconnect()
  }, [])

  // Web fonts (Google Sans / Roboto in the toolbar itself, plus doc fonts
  // used by font-name pills) can finish loading after this first layout,
  // changing measured widths and making the collapse point stale. Recalc
  // once fonts settle, and again on each individual face swap-in.
  useEffect(() => {
    const forceRecalc = () => {
      lastQuantizedWidth.current = -1
      recalcCollapse()
    }
    void document.fonts.ready.then(forceRecalc)
    document.fonts.addEventListener('loadingdone', forceRecalc)
    return () => document.fonts.removeEventListener('loadingdone', forceRecalc)
  }, [])

  useEffect(() => {
    lastQuantizedWidth.current = -1 // force a fresh recalc; label/text widths may have changed
    recalcCollapse()
  }, [currentFont, activeStyleKey, fs.fontSizePt, zoom, canEdit, lang])

  // ---- Group content (index-aligned with GROUP_IDS) ------------------------
  const groupContents: React.ReactNode[] = [
    // 0: undo / redo
    <>
      {btn(
        'undo',
        <IconUndoM />,
        'Undo (⌘Z)',
        () => editor.chain().focus().undo().run(),
        false,
        !hasDoc,
      )}
      {btn(
        'redo',
        <IconRedoM />,
        'Redo (⌘⇧Z)',
        () => editor.chain().focus().redo().run(),
        false,
        !hasDoc,
      )}
    </>,

    // 1: bold / italic / underline
    <>
      {btn('bold', <b>B</b>, 'Bold (⌘B)', () => chain()?.toggleMark('bold').run(), fs.bold)}
      {btn('italic', <i>I</i>, 'Italic (⌘I)', () => chain()?.toggleMark('italic').run(), fs.italic)}
      {btn(
        'underline',
        <u>U</u>,
        'Underline (⌘U)',
        () => chain()?.toggleMark('underline').run(),
        fs.underline,
      )}
    </>,

    // 2: font family
    <div className="gdh-tb-dd gdh-tb-font">
      <button
        type="button"
        className="gdh-tb-select"
        disabled={!canEdit}
        onClick={() => setDropdown((v) => (v === 'font' ? null : 'font'))}
      >
        {currentFont || bodyFontName}
        <IconCaret />
      </button>
      {dropdown === 'font' && (
        <div className="gdh-tb-menu gdh-tb-menu-font">
          <button className={!currentFont ? 'active' : ''} onClick={() => setFont(null)}>
            {bodyFontName} (body)
          </button>
          {fontFamilies
            .filter((f) => f !== bodyFontName)
            .map((f) => (
              <button
                key={f}
                className={f === currentFont ? 'active' : ''}
                style={{ fontFamily: cssFontFamily(f) }}
                onClick={() => setFont(f)}
              >
                {f}
              </button>
            ))}
        </div>
      )}
    </div>,

    // 3: font size stepper
    <div className="gdh-tb-fontsize">
      <button
        type="button"
        className="gdh-tb-step"
        disabled={!canEdit}
        title="Decrease font size"
        onClick={() => stepFontSize(-1)}
      >
        −
      </button>
      <span className="gdh-tb-fontsize-val">{Math.round(fs.fontSizePt)}</span>
      <button
        type="button"
        className="gdh-tb-step"
        disabled={!canEdit}
        title="Increase font size"
        onClick={() => stepFontSize(1)}
      >
        +
      </button>
    </div>,

    // 4: styles dropdown
    <div className="gdh-tb-dd gdh-tb-styles">
      <button
        type="button"
        className="gdh-tb-select"
        disabled={!canEdit || !!sub}
        onClick={() => setDropdown((v) => (v === 'styles' ? null : 'styles'))}
      >
        {STYLES.find((s) => s.key === activeStyleKey)?.label ?? 'Normal text'}
        <IconCaret />
      </button>
      {dropdown === 'styles' && (
        <div className="gdh-tb-menu">
          {STYLES.map((s) => (
            <button
              key={s.key}
              className={activeStyleKey === s.key ? 'active' : ''}
              onClick={() => applyStyle(s.key)}
            >
              {s.label}
            </button>
          ))}
        </div>
      )}
    </div>,

    // 5: text color / highlight / clear formatting
    <>
      <div className="gdh-tb-dd">
        <button
          type="button"
          className="gdh-tb-btn"
          disabled={!canEdit}
          title="Text color"
          onClick={() => setDropdown((v) => (v === 'color' ? null : 'color'))}
        >
          <span className="gdh-tb-a">A</span>
          <span
            className="gdh-tb-colorbar"
            style={{ background: fs.textColor ? `#${fs.textColor}` : '#000' }}
          />
        </button>
        {dropdown === 'color' && (
          <div className="gdh-color-grid">
            <button className="gdh-color-reset" onClick={() => setTextStyle({ color: null })}>
              Reset color
            </button>
            <div className="gdh-color-swatches">
              {TEXT_COLORS.map((c) => (
                <button
                  key={c}
                  className="gdh-color-swatch"
                  style={{ background: `#${c}` }}
                  onClick={() => setTextStyle({ color: c })}
                />
              ))}
            </div>
          </div>
        )}
      </div>
      <div className="gdh-tb-dd">
        <button
          type="button"
          className={`gdh-tb-btn ${fs.highlight ? 'active' : ''}`}
          disabled={!canEdit}
          title="Highlight color"
          onClick={() => setDropdown((v) => (v === 'highlight' ? null : 'highlight'))}
        >
          <IconHighlight />
        </button>
        {dropdown === 'highlight' && (
          <div className="gdh-color-grid">
            <button className="gdh-color-reset" onClick={() => setTextStyle({ highlight: null })}>
              None
            </button>
            <div className="gdh-color-swatches">
              {Object.entries(HIGHLIGHT_CSS)
                .filter(([k]) => k !== 'none')
                .map(([k, css]) => (
                  <button
                    key={k}
                    className="gdh-color-swatch"
                    style={{ background: css as string }}
                    onClick={() => setTextStyle({ highlight: k })}
                  />
                ))}
            </div>
          </div>
        )}
      </div>
      {btn('clear', <IconClearFormat />, 'Clear formatting (⌘\\)', () =>
        chain()?.unsetAllMarks().run(),
      )}
    </>,

    // 6: lists / indent
    <>
      {btn(
        'bullets',
        <IconBullets />,
        'Bulleted list',
        () => toggleList('bullet'),
        fs.listBullet,
        !!sub,
      )}
      {btn(
        'numbered',
        <IconNumbered />,
        'Numbered list',
        () => toggleList('ordered'),
        fs.listOrdered,
        !!sub,
      )}
      {btn('outdent', <IconIndentDec />, 'Decrease indent', () => changeIndent(-1), false, !!sub)}
      {btn('indent', <IconIndentInc />, 'Increase indent', () => changeIndent(1), false, !!sub)}
    </>,

    // 7: align / spacing
    <>
      <div className="gdh-tb-dd">
        <button
          type="button"
          className="gdh-tb-btn"
          disabled={!canEdit || !!sub}
          title="Align"
          onClick={() => setDropdown((v) => (v === 'align' ? null : 'align'))}
        >
          <AlignIcon />
          <IconCaret />
        </button>
        {dropdown === 'align' && (
          <div className="gdh-tb-menu gdh-tb-menu-icons">
            {(['left', 'center', 'right', 'justify'] as const).map((a) => {
              const AIcon = ALIGN_ICONS[a]
              return (
                <button
                  key={a}
                  className={activeAlign === a ? 'active' : ''}
                  onClick={() => setParaAttr({ align: a })}
                >
                  <AIcon /> {a[0].toUpperCase() + a.slice(1)}
                </button>
              )
            })}
          </div>
        )}
      </div>
      <div className="gdh-tb-dd">
        <button
          type="button"
          className="gdh-tb-btn"
          disabled={!canEdit || !!sub}
          title="Line & paragraph spacing"
          onClick={() => setDropdown((v) => (v === 'spacing' ? null : 'spacing'))}
        >
          <IconLineSpacing />
        </button>
        {dropdown === 'spacing' && (
          <div className="gdh-tb-menu">
            {LINE_SPACINGS.map((v) => (
              <button
                key={v}
                className={fs.lineSpacing === v ? 'active' : ''}
                onClick={() => setParaAttr({ lineSpacing: v })}
              >
                {v.toFixed(2).replace(/0+$/, '').replace(/\.$/, '')}
              </button>
            ))}
          </div>
        )}
      </div>
    </>,

    // 8: link / image / comment
    <>
      {btn('link', <IconLink />, 'Insert link (⌘K)', onLink, false, !canEdit)}
      {btn('image', <IconPicture />, 'Insert image', onInsertImage, false, !canEdit)}
      {btn('comment', <IconComment />, 'Add comment', onNewComment, false, !canComment)}
    </>,

    // 9: zoom
    <div className="gdh-tb-dd">
      <button
        type="button"
        className="gdh-tb-select"
        disabled={!hasDoc}
        onClick={() => setDropdown((v) => (v === 'zoom' ? null : 'zoom'))}
      >
        {Math.round(zoom)}%
      </button>
      {dropdown === 'zoom' && (
        <div className="gdh-tb-menu gdh-tb-menu-zoom">
          <button onClick={() => (onZoomFit('width'), setDropdown(null))}>Fit to width</button>
          <button onClick={() => (onZoomFit('page'), setDropdown(null))}>Fit whole page</button>
          <div className="gdh-menu-divider" />
          {ZOOM_PRESETS.map((z) => (
            <button
              key={z}
              className={Math.round(zoom) === z ? 'active' : ''}
              onClick={() => (onZoom(z), setDropdown(null))}
            >
              {z}%
            </button>
          ))}
        </div>
      )}
    </div>,

    // 10: print
    <>{btn('print', <IconPrint />, 'Print (⌘P)', () => window.print(), false, !hasDoc)}</>,
  ]

  const collapsed = visibleCount < GROUP_IDS.length

  useEffect(() => {
    if (!collapsed) setMoreOpen(false)
  }, [collapsed])

  return (
    <div className="gdh-toolbar" ref={ref}>
      <div
        className={`gdh-toolbar-pill${collapsed ? ' gdh-toolbar-pill--left' : ''}`}
        ref={pillRef}
      >
        {GROUP_IDS.slice(0, visibleCount).map((id, i) => (
          <Fragment key={id}>
            <div className="gdh-tb-group">{groupContents[i]}</div>
            {i < visibleCount - 1 && <span className="gdh-tb-sep" />}
          </Fragment>
        ))}

        {collapsed && (
          <>
            <span className="gdh-tb-sep" />
            <div className="gdh-tb-dd">
              <button
                type="button"
                className="gdh-tb-btn"
                title="More options"
                onClick={() => setMoreOpen((v) => !v)}
              >
                <IconMoreVert />
              </button>
              {moreOpen && (
                <div className="gdh-tb-menu gdh-tb-menu-more">
                  {GROUP_IDS.slice(visibleCount).map((id, i) => (
                    <div className="gdh-tb-more-row" key={`more-${id}`}>
                      {groupContents[visibleCount + i]}
                    </div>
                  ))}
                </div>
              )}
            </div>
          </>
        )}
      </div>

      {/* Off-screen measurer: identical group markup inside the pill's real
          classes so real gaps/padding are baked into the widths. Never
          painted (visibility hidden + pointer-events none); used purely to
          decide how many groups fit before anything would escape the pill. */}
      <div
        className="gdh-toolbar-pill gdh-tb-measurer"
        aria-hidden="true"
        style={{
          position: 'fixed',
          top: 0,
          left: 0,
          visibility: 'hidden',
          pointerEvents: 'none',
          width: 'max-content',
          maxWidth: 'none',
          zIndex: -1,
        }}
      >
        {GROUP_IDS.map((id, i) => (
          <Fragment key={`m-${id}`}>
            <div
              className="gdh-tb-group"
              ref={(el) => {
                groupRefs.current[i] = el
              }}
            >
              {groupContents[i]}
            </div>
            {i < GROUP_IDS.length - 1 && <span className="gdh-tb-sep" />}
          </Fragment>
        ))}
      </div>
    </div>
  )
}
