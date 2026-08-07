import { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n, type StringKey } from '../i18n/locale'
import { fontFamiliesFor, isEastAsianFontName } from '../font-list'
import { cssFontFamily } from '../line-metrics'
import { setParaAttrs, activeParaAttrs } from './ribbon-tabs'
import { setSelectionAlign } from '../editor/direction'
import { IconSparkle } from './icons'
import { useModalKeys } from './modal-keys'

/**
 * Editor context menu (right click in the document body):
 * Cut/Copy/Paste · Font… · Paragraph… · Synonyms · Translate · Hyperlink… · New Comment.
 */

export interface ContextMenuState {
  x: number
  y: number
}

interface EditorContextMenuProps {
  editor: Editor
  menu: ContextMenuState
  onClose: () => void
  onFontDialog: () => void
  onParagraphDialog: () => void
  onLink: () => void
  onNewComment: () => void
  onAiPreset: (instruction: string) => void
  /** List items: restart numbering / continue numbering (shown when the cursor is on a docListItem) */
  onRestartNumbering?: () => void
  onContinueNumbering?: () => void
  /** F9 update fields (shown when the cursor is on an inline field) */
  onUpdateFields?: () => void
}

/** target languages mirrored from the Review → Translate dropdown; the localized label also goes into the LLM prompt */
const TRANSLATE_TARGETS: Array<{ labelKey: StringKey }> = [
  { labelKey: 'appLangEnglish' },
  { labelKey: 'appLangSimplifiedChinese' },
  { labelKey: 'appLangJapanese' },
  { labelKey: 'appLangKorean' },
  { labelKey: 'appLangFrench' },
  { labelKey: 'appLangGerman' },
  { labelKey: 'appLangSpanish' },
]

const MENU_WIDTH = 240

export function EditorContextMenu({
  editor,
  menu,
  onClose,
  onFontDialog,
  onParagraphDialog,
  onLink,
  onNewComment,
  onAiPreset,
  onRestartNumbering,
  onContinueNumbering,
  onUpdateFields,
}: EditorContextMenuProps) {
  const { t } = useI18n()
  const ref = useRef<HTMLDivElement>(null)
  const [submenu, setSubmenu] = useState<string | null>(null)

  const { from, to } = editor.state.selection
  const hasSelection = from !== to
  const canEdit = editor.isEditable
  const selectedText = hasSelection ? editor.state.doc.textBetween(from, to, ' ').trim() : ''
  // Synonyms targets a word / short phrase, not long selections
  const synonymText = selectedText.length > 0 && selectedText.length <= 20 ? selectedText : ''

  // keep the menu inside the viewport (flip up / clamp left near the edges)
  const [pos, setPos] = useState({ left: menu.x, top: menu.y })
  useEffect(() => {
    const el = ref.current
    if (!el) return
    const rect = el.getBoundingClientRect()
    let left = menu.x
    let top = menu.y
    if (left + rect.width > window.innerWidth - 8) left = window.innerWidth - rect.width - 8
    if (top + rect.height > window.innerHeight - 8) top = Math.max(8, menu.y - rect.height)
    setPos({ left, top })
  }, [menu])

  useEffect(() => {
    const close = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose()
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('mousedown', close)
    window.addEventListener('keydown', onKey)
    window.addEventListener('blur', onClose)
    return () => {
      window.removeEventListener('mousedown', close)
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('blur', onClose)
    }
  }, [onClose])

  const run = (action: () => void) => () => {
    onClose()
    action()
  }

  const protAttrs = editor.getAttributes('docProtected')
  const isImage = protAttrs?.blockType === 'image'
  const isFloating =
    isImage ||
    (Array.isArray(protAttrs?.textboxes) && (protAttrs.textboxes as unknown[]).length > 0)
  const currentWrap = (protAttrs?.imageWrap as string | null) ?? null
  const setWrap = (wrap: string | null) =>
    editor.chain().focus().updateAttributes('docProtected', { imageWrap: wrap }).run()

  /** Plain-text insertion: no HTML parsing (insertContent(string) would treat < > as tags) */
  const insertPlainText = (text: string) => {
    const lines = text.replace(/\r/g, '').split('\n')
    if (lines.length === 1) {
      editor.chain().focus().insertContent({ type: 'text', text: lines[0] }).run()
      return
    }
    editor
      .chain()
      .focus()
      .insertContent(
        lines.map((line) => ({
          type: 'docParagraph',
          ...(line ? { content: [{ type: 'text', text: line }] } : {}),
        })),
      )
      .run()
  }

  const clipboard = async (action: 'cut' | 'copy' | 'paste' | 'pastePlain') => {
    if (action === 'paste') {
      // rich paste: prefer HTML (parsed by editor paste rules, keeps formatting), otherwise plain text
      try {
        const items = await navigator.clipboard.read()
        for (const item of items) {
          if (item.types.includes('text/html')) {
            const html = await (await item.getType('text/html')).text()
            editor
              .chain()
              .focus()
              .insertContent(html, { parseOptions: { preserveWhitespace: true } })
              .run()
            return
          }
        }
      } catch {
        /* clipboard.read unavailable (permission/format): fall back to plain text */
      }
      const text = await navigator.clipboard.readText()
      if (text) insertPlainText(text)
    } else if (action === 'pastePlain') {
      const text = await navigator.clipboard.readText()
      if (text) insertPlainText(text)
    } else {
      editor.commands.focus()
      document.execCommand(action)
    }
  }

  const item = (
    label: string,
    opts: {
      key?: string
      disabled?: boolean
      onClick?: () => void
      submenuKey?: string
      ai?: boolean
    },
  ) => (
    <button
      className="ctx-item"
      disabled={opts.disabled}
      title={opts.ai ? t('appAiBadgeTip') : undefined}
      onMouseEnter={() => setSubmenu(opts.submenuKey ?? null)}
      onClick={opts.submenuKey ? undefined : opts.onClick}
    >
      <span className="ctx-label">{label}</span>
      {opts.ai && (
        <span className="copilot-badge copilot-badge-menu">
          <IconSparkle size={10} />
        </span>
      )}
      {opts.key && <span className="ctx-key">{opts.key}</span>}
      {opts.submenuKey && <span className="ctx-arrow">›</span>}
    </button>
  )

  return (
    <div
      ref={ref}
      className="ctx-menu"
      style={{ left: pos.left, top: pos.top, minWidth: MENU_WIDTH }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {item(t('appCut'), {
        key: '⌘X',
        disabled: !hasSelection || !canEdit,
        onClick: run(() => void clipboard('cut')),
      })}
      {item(t('appCopy'), {
        key: '⌘C',
        disabled: !hasSelection,
        onClick: run(() => void clipboard('copy')),
      })}
      {item(t('appPaste'), {
        key: '⌘V',
        disabled: !canEdit,
        onClick: run(() => void clipboard('paste')),
      })}
      {item(t('appPastePlain'), {
        disabled: !canEdit,
        onClick: run(() => void clipboard('pastePlain')),
      })}
      <div className="ctx-sep" />
      {item(t('appFontMenu'), { key: '⌘D', onClick: run(onFontDialog) })}
      {item(t('appParagraphMenu'), { key: '⌥⌘M', onClick: run(onParagraphDialog) })}
      {editor.isActive('instrField') && onUpdateFields && (
        <>
          <div className="ctx-sep" />
          {item(t('appUpdateField'), { key: 'F9', onClick: run(() => onUpdateFields()) })}
        </>
      )}
      {editor.isActive('docListItem') && !!editor.getAttributes('docListItem').numId && (
        <>
          <div className="ctx-sep" />
          {item(t('appRestartNumbering'), {
            disabled: !onRestartNumbering,
            onClick: run(() => onRestartNumbering?.()),
          })}
          {item(t('appContinueNumbering'), {
            disabled: !onContinueNumbering,
            onClick: run(() => onContinueNumbering?.()),
          })}
        </>
      )}
      <div className="ctx-sep" />
      {item(t('appSynonyms'), {
        disabled: !synonymText,
        ai: true,
        onClick: run(() => onAiPreset(t('appSynonymsPrompt', { text: synonymText }))),
      })}
      <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
        {item(t('appTranslate'), { disabled: !hasSelection, submenuKey: 'translate', ai: true })}
        {submenu === 'translate' && hasSelection && (
          <div className="ctx-submenu">
            {TRANSLATE_TARGETS.map((target) => (
              <button
                key={target.labelKey}
                className="ctx-item"
                onClick={run(() =>
                  onAiPreset(
                    t('appTranslateSelectionPrompt', {
                      lang: t(target.labelKey),
                      text: selectedText,
                    }),
                  ),
                )}
              >
                <span className="ctx-label">
                  {t('appTranslateTo', { lang: t(target.labelKey) })}
                </span>
              </button>
            ))}
          </div>
        )}
      </div>
      {isFloating && (
        <>
          <div className="ctx-sep" />
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('appWrapTextMenu'), { submenuKey: 'wrap' })}
            {submenu === 'wrap' && (
              <div className="ctx-submenu">
                {WRAP_OPTIONS.map((opt) => (
                  <button
                    key={String(opt.value)}
                    className="ctx-item"
                    onClick={run(() => setWrap(opt.value))}
                  >
                    <span className="ctx-label">
                      {currentWrap === opt.value ? '✓ ' : ''}
                      {t(opt.labelKey)}
                    </span>
                  </button>
                ))}
              </div>
            )}
          </div>
        </>
      )}
      <div className="ctx-sep" />
      {item(t('appHyperlinkMenu'), { key: '⌘K', onClick: run(onLink) })}
      {item(t('appNewComment'), { disabled: !hasSelection, onClick: run(onNewComment) })}
    </div>
  )
}

/** Wrap options from Word's image context menu (inline/square/top-bottom/behind text/in front of text); shared with the Picture Format tab */
export const WRAP_OPTIONS: Array<{ labelKey: StringKey; value: string | null }> = [
  { labelKey: 'appWrapInline', value: null },
  { labelKey: 'appWrapSquareLeft', value: 'square-left' },
  { labelKey: 'appWrapSquareRight', value: 'square-right' },
  { labelKey: 'appWrapTopBottom', value: 'topBottom' },
  { labelKey: 'appWrapBehind', value: 'behind' },
  { labelKey: 'appWrapFront', value: 'front' },
]

/* ================= Font dialog ================= */

const FONT_SIZES = [9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72]

const FONT_STYLES: Array<{ key: string; nameKey: StringKey }> = [
  { key: 'regular', nameKey: 'appFontRegular' },
  { key: 'italic', nameKey: 'appFontItalic' },
  { key: 'bold', nameKey: 'appFontBold' },
  { key: 'boldItalic', nameKey: 'appFontBoldItalic' },
]

export function FontDialog({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t, lang } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const fontFamilies = fontFamiliesFor(lang)
  const textAttrs = editor.getAttributes('docTextStyle')
  const initialStyle = editor.isActive('bold')
    ? editor.isActive('italic')
      ? 'boldItalic'
      : 'bold'
    : editor.isActive('italic')
      ? 'italic'
      : 'regular'

  const [font, setFont] = useState(
    (textAttrs.font as string | null) ?? (textAttrs.fontAscii as string | null) ?? '',
  )
  const [size, setSize] = useState(
    textAttrs.sizeHalfPoints ? Number(textAttrs.sizeHalfPoints) / 2 : 11,
  )
  const [style, setStyle] = useState<string>(initialStyle)
  const [color, setColor] = useState(`#${(textAttrs.color as string | null) ?? '000000'}`)
  const [underline, setUnderline] = useState(editor.isActive('underline'))
  const [strike, setStrike] = useState(editor.isActive('strike'))
  const [vertAlign, setVertAlign] = useState<string>((textAttrs.vertAlign as string | null) ?? '')

  const apply = () => {
    if (!editor.isEditable) {
      onClose()
      return
    }
    const hex = color.replace('#', '').toUpperCase()
    let chain = editor
      .chain()
      .focus()
      .setMark('docTextStyle', {
        color: hex === '000000' ? null : hex,
        sizeHalfPoints: Math.round(size * 2),
        // picks target only their script's rFonts slot; the other slot survives
        ...(!font
          ? { font: null, fontAscii: null }
          : isEastAsianFontName(font)
            ? { font }
            : { fontAscii: font }),
        highlight: textAttrs.highlight ?? null,
        vertAlign: vertAlign || null,
      })
    const wantBold = style === 'bold' || style === 'boldItalic'
    const wantItalic = style === 'italic' || style === 'boldItalic'
    chain = wantBold ? chain.setMark('bold') : chain.unsetMark('bold')
    chain = wantItalic ? chain.setMark('italic') : chain.unsetMark('italic')
    chain = underline ? chain.setMark('underline') : chain.unsetMark('underline')
    chain = strike ? chain.setMark('strike') : chain.unsetMark('strike')
    chain.run()
    onClose()
  }

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal">
        <h2>{t('appFontDialogTitle')}</h2>
        <div className="font-dialog-row">
          <label>
            {t('appFontFamilyLabel')}
            <select value={font} onChange={(e) => setFont(e.target.value)}>
              <option value="">{t('appDefaultBodyFont')}</option>
              {fontFamilies.map((f) => (
                <option key={f} value={f} style={{ fontFamily: cssFontFamily(f) }}>
                  {f}
                </option>
              ))}
              {font && !fontFamilies.includes(font) && <option value={font}>{font}</option>}
            </select>
          </label>
          <label>
            {t('appFontStyleLabel')}
            <select value={style} onChange={(e) => setStyle(e.target.value)}>
              {FONT_STYLES.map((s) => (
                <option key={s.key} value={s.key}>
                  {t(s.nameKey)}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t('appFontSizeLabel')}
            <select value={size} onChange={(e) => setSize(Number(e.target.value))}>
              {!FONT_SIZES.includes(size) && <option value={size}>{size}</option>}
              {FONT_SIZES.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </label>
        </div>
        <div className="font-dialog-row">
          <label>
            {t('appFontColor')}
            <input
              type="color"
              className="font-color-input"
              value={color}
              onChange={(e) => setColor(e.target.value)}
            />
          </label>
          <label className="font-check">
            <input
              type="checkbox"
              checked={underline}
              onChange={(e) => setUnderline(e.target.checked)}
            />
            {t('appUnderline')}
          </label>
          <label className="font-check">
            <input type="checkbox" checked={strike} onChange={(e) => setStrike(e.target.checked)} />
            {t('appStrikethrough')}
          </label>
          <label className="font-check">
            <input
              type="checkbox"
              checked={vertAlign === 'superscript'}
              onChange={(e) => setVertAlign(e.target.checked ? 'superscript' : '')}
            />
            {t('appSuperscript')}
          </label>
          <label className="font-check">
            <input
              type="checkbox"
              checked={vertAlign === 'subscript'}
              onChange={(e) => setVertAlign(e.target.checked ? 'subscript' : '')}
            />
            {t('appSubscript')}
          </label>
        </div>
        <div
          className="font-preview"
          style={{
            fontFamily: font ? cssFontFamily(font) : undefined,
            fontSize: `${Math.min(size, 28)}pt`,
            fontWeight: style === 'bold' || style === 'boldItalic' ? 600 : 400,
            fontStyle: style === 'italic' || style === 'boldItalic' ? 'italic' : 'normal',
            color,
            textDecoration:
              [underline ? 'underline' : '', strike ? 'line-through' : ''].join(' ').trim() ||
              undefined,
          }}
        >
          {t('appFontPreviewSample')}
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" onClick={apply}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

/* ================= Paragraph dialog ================= */

const PT_PER_TWIP = 1 / 20

type AlignValue = 'left' | 'center' | 'right' | 'justify'

/** visual alignment values; setSelectionAlign resolves them per paragraph direction */
const ALIGN_OPTIONS: Array<{ key: AlignValue; nameKey: StringKey }> = [
  { key: 'left', nameKey: 'appAlignLeft' },
  { key: 'center', nameKey: 'appAlignCenter' },
  { key: 'right', nameKey: 'appAlignRight' },
  { key: 'justify', nameKey: 'appAlignJustify' },
]

const LINE_SPACINGS: Array<{ value: number; nameKey: StringKey }> = [
  { value: 1, nameKey: 'appLineSingle' },
  { value: 1.15, nameKey: 'appLine115' },
  { value: 1.5, nameKey: 'appLine15' },
  { value: 2, nameKey: 'appLineDouble' },
  { value: 2.5, nameKey: 'appLine25' },
  { value: 3, nameKey: 'appLineTriple' },
]

export function ParagraphDialog({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const attrs = activeParaAttrs(editor)
  // unset align means "start": visually left in LTR, right in RTL (same as the ribbon)
  const [align, setAlign] = useState<AlignValue>(
    (attrs.align as AlignValue | null) ?? (attrs.bidi === true ? 'right' : 'left'),
  )
  // Line spacing rule: multiples are the preset select values; 'atLeast'/'exact'
  // take a pt value (w:spacing w:lineRule + w:line, already modeled by parse/save)
  const rawTwips = Number(attrs.lineRawTwips) || 0
  const initRule =
    attrs.lineRule === 'exact' || attrs.lineRule === 'atLeast' ? (attrs.lineRule as string) : ''
  const initMultiple =
    Number(attrs.lineSpacing) || (attrs.lineRule === 'auto' && rawTwips ? rawTwips / 240 : 1)
  const [lineRule, setLineRule] = useState<string>(initRule)
  const [lineSpacing, setLineSpacing] = useState<number>(Math.round(initMultiple * 100) / 100)
  const [linePt, setLinePt] = useState<number>(rawTwips ? Math.round(rawTwips / 2) / 10 : 12)
  const twipsToPt = (v: unknown) => Math.round((Number(v) || 0) * PT_PER_TWIP)
  const [indentLeft, setIndentLeft] = useState(twipsToPt(attrs.indentLeft))
  const [indentRight, setIndentRight] = useState(twipsToPt(attrs.indentRight))
  const [indentFirstLine, setIndentFirstLine] = useState(twipsToPt(attrs.indentFirstLine))
  const [spaceBefore, setSpaceBefore] = useState(twipsToPt(attrs.spaceBefore))
  const [spaceAfter, setSpaceAfter] = useState(twipsToPt(attrs.spaceAfter))

  const apply = () => {
    if (!editor.isEditable) {
      onClose()
      return
    }
    const ptToTwips = (pt: number) => (pt > 0 ? Math.round(pt / PT_PER_TWIP) : null)
    const spacing =
      lineRule === 'exact' || lineRule === 'atLeast'
        ? { lineSpacing: null, lineRule, lineRawTwips: Math.max(20, Math.round(linePt * 20)) }
        : {
            lineSpacing: lineSpacing === 1 ? null : lineSpacing,
            lineRule: null,
            lineRawTwips: null,
          }
    // align goes through setSelectionAlign so each paragraph resolves the
    // visual value against its own direction (null = start side)
    setSelectionAlign(editor, align)
    setParaAttrs(editor, {
      ...spacing,
      indentLeft: ptToTwips(indentLeft),
      indentRight: ptToTwips(indentRight),
      indentFirstLine: ptToTwips(indentFirstLine),
      spaceBefore: ptToTwips(spaceBefore),
      spaceAfter: ptToTwips(spaceAfter),
    })
    onClose()
  }

  const numInput = (label: string, value: number, set: (v: number) => void) => (
    <label>
      {label}
      <span className="para-num">
        <input
          type="number"
          min={0}
          max={400}
          value={value}
          onChange={(e) => set(Math.max(0, Number(e.target.value) || 0))}
        />
        <span className="para-unit">pt</span>
      </span>
    </label>
  )

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal">
        <h2>{t('appParagraph')}</h2>
        <div className="font-dialog-row">
          <label>
            {t('appAlignment')}
            <select value={align} onChange={(e) => setAlign(e.target.value as AlignValue)}>
              {ALIGN_OPTIONS.map((o) => (
                <option key={o.key} value={o.key}>
                  {t(o.nameKey)}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t('appLineSpacingLabel')}
            <select
              value={lineRule || String(lineSpacing)}
              onChange={(e) => {
                const v = e.target.value
                if (v === 'exact' || v === 'atLeast' || v === 'multiple') {
                  setLineRule(v === 'multiple' ? '' : v)
                  if (v === 'multiple') setLineSpacing(1.25)
                } else {
                  setLineRule('')
                  setLineSpacing(Number(v))
                }
              }}
            >
              {!lineRule && !LINE_SPACINGS.some((s) => s.value === lineSpacing) && (
                <option value={lineSpacing}>{t('appLineMultiple', { n: lineSpacing })}</option>
              )}
              {LINE_SPACINGS.map((s) => (
                <option key={s.value} value={s.value}>
                  {t(s.nameKey)}
                </option>
              ))}
              <option value="atLeast">{t('appLineAtLeast')}</option>
              <option value="exact">{t('appLineExactly')}</option>
            </select>
          </label>
          {lineRule === 'exact' || lineRule === 'atLeast' ? (
            <label>
              {t('appLineValue')}
              <span className="para-num">
                <input
                  type="number"
                  min={1}
                  max={1584}
                  step={0.5}
                  value={linePt}
                  onChange={(e) => setLinePt(Math.max(0, Number(e.target.value) || 0))}
                />
                <span className="para-unit">pt</span>
              </span>
            </label>
          ) : (
            <label>
              {t('appLineValue')}
              <span className="para-num">
                <input
                  type="number"
                  min={0.06}
                  max={132}
                  step={0.05}
                  value={lineSpacing}
                  onChange={(e) => setLineSpacing(Math.max(0.06, Number(e.target.value) || 1))}
                />
                <span className="para-unit">×</span>
              </span>
            </label>
          )}
        </div>
        <div className="font-dialog-row">
          {numInput(t('appIndentLeft'), indentLeft, setIndentLeft)}
          {numInput(t('appIndentRight'), indentRight, setIndentRight)}
          {numInput(t('appIndentFirstLine'), indentFirstLine, setIndentFirstLine)}
        </div>
        <div className="font-dialog-row">
          {numInput(t('appSpaceBefore'), spaceBefore, setSpaceBefore)}
          {numInput(t('appSpaceAfter'), spaceAfter, setSpaceAfter)}
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" onClick={apply}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
