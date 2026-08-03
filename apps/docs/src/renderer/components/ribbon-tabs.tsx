import { useState } from 'react'
import type { CSSProperties } from 'react'
import type { Editor, JSONContent } from '@tiptap/core'
import {
  buildShapeParagraphXml,
  buildTextboxParagraphXml,
  buildWordArtParagraphXml,
  WORDART_PRESETS,
  type HeaderFooter,
  type TextboxDisplay,
} from '@genoffice/docx-engine'
import type { DocsTabInfo } from '../../shared/ipc'
import { tableModelToPmNode } from '../editor/convert'
import type { InkTool } from '../editor/ink'
import { t, useI18n, type StringKey } from '../i18n/locale'
import {
  IconAccept,
  IconAiPanel,
  IconCaret,
  IconComment,
  IconCompare,
  IconCursor,
  IconEraser,
  IconHighlighterPen,
  IconPen,
  IconLock,
  IconMoon,
  IconNavPane,
  IconOutlineView,
  IconPageWidth,
  IconGridlines,
  IconNewWindow,
  IconPrintLayout,
  IconReadMode,
  IconRuler,
  IconSparkle,
  IconSplit,
  IconSwitchWindows,
  IconRedo,
  IconReject,
  IconTrackChanges,
  IconTranslate,
  IconUndo,
  IconWebLayout,
  IconWholePage,
  IconZoom100,
  IconZoomIn,
  IconZoomOut,
} from './icons'

export {
  BookmarkModal,
  ChartInsertModal,
  CrossRefModal,
  InsertTab,
  LinkInsertModal,
} from './ribbon-insert-tab'
export { DesignTab } from './ribbon-design-tab'
export { LayoutTab } from './ribbon-layout-tab'
export { ReferencesTab } from './ribbon-references-tab'

/** icon size for the big icon-over-label ribbon buttons (slides ribbon parity) */
export const BIG = 28

/* ---------- shared helpers ---------- */

export type SetDropdown = (updater: (prev: string | null) => string | null) => void

export const toggleDropdown = (setDropdown: SetDropdown, key: string) =>
  setDropdown((prev) => (prev === key ? null : key))

/** apply paragraph-level attrs to every block type in the selection */
export function setParaAttrs(editor: Editor, attrs: Record<string, unknown>): void {
  let chain = editor
    .chain()
    .focus()
    .updateAttributes('docParagraph', attrs)
    .updateAttributes('docHeading', attrs)
    .updateAttributes('docListItem', attrs)
  // alignment also applies to selected images (w:jc on the image paragraph)
  if ('align' in attrs) {
    chain = chain.updateAttributes('docProtected', { imageAlign: attrs.align ?? null })
  }
  chain.run()
}

/** attrs of the paragraph-like node at the cursor */
export function activeParaAttrs(editor: Editor): Record<string, unknown> {
  if (editor.isActive('docHeading')) return editor.getAttributes('docHeading')
  if (editor.isActive('docListItem')) return editor.getAttributes('docListItem')
  return editor.getAttributes('docParagraph')
}

export interface TabProps {
  editor: Editor
  hasDoc: boolean
  dropdown: string | null
  setDropdown: SetDropdown
}

/* ================= Insert ================= */

const MAX_IMAGE_WIDTH_PX = 620 // ~content width of a US Letter page at 96dpi

export async function imageSizeOf(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight })
    img.onerror = () => reject(new Error('Failed to read image'))
    img.src = dataUrl
  })
}

/* insert commands shared by the ribbon and the native application menu */

export function insertTableAt(editor: Editor, rows: number, cols: number): void {
  const table = {
    rows: Array.from({ length: rows }, () => Array.from({ length: cols }, () => ({ paras: [''] }))),
    colWidthsPct: Array.from({ length: cols }, () => 100 / cols),
  }
  // inside a cell a top-level docTable insert would split the outer table
  // — Word semantics is a nested child table at the end of the cell
  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth > 0; depth--) {
    const name = $from.node(depth).type.name
    if (name === 'docTableCell' || name === 'docTableHeader') {
      editor
        .chain()
        .focus()
        .insertContentAt($from.end(depth), { type: 'docNestedTable', attrs: { model: table } })
        .run()
      return
    }
  }
  editor.chain().focus().insertContent(tableModelToPmNode(table)).run()
}

/** Insert an inline image from a dataURL at the cursor (shared by paste/dialog; size scaled to content width) */
export async function insertImageFromDataUrl(
  editor: Editor,
  dataUrl: string,
  label = t('ribbonPicture'),
): Promise<boolean> {
  const m = /^data:([^;]+);base64,(.*)$/s.exec(dataUrl)
  if (!m) return false
  const mime = m[1]
  if (!/^image\/(png|jpeg|gif)$/.test(mime)) return false
  try {
    const natural = await imageSizeOf(dataUrl)
    const scale = Math.min(1, MAX_IMAGE_WIDTH_PX / natural.width)
    editor
      .chain()
      .focus()
      .insertContent({
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'image',
          label,
          imageDataUrl: dataUrl,
          imageWidthPx: Math.round(natural.width * scale),
          imageHeightPx: Math.round(natural.height * scale),
          genImage: {
            base64: m[2],
            mime,
            widthPx: Math.round(natural.width * scale),
            heightPx: Math.round(natural.height * scale),
          },
        },
      })
      .run()
    return true
  } catch {
    return false
  }
}

export async function insertImageViaDialog(editor: Editor): Promise<void> {
  const picked = await window.desktop.pickImage()
  if (!picked) return
  await insertImageFromDataUrl(
    editor,
    `data:${picked.mime};base64,${picked.base64}`,
    t('ribbonPictureLabel', { name: picked.name }),
  )
}

/** 5 cm × 3 cm default textbox size in EMU (1 cm = 360000 EMU) */
const TEXTBOX_WIDTH_EMU = 1800000
const TEXTBOX_HEIGHT_EMU = 1080000

/** Default TextboxDisplay model for a freshly inserted empty textbox */
function emptyTextboxDisplay(): TextboxDisplay {
  return {
    fill: 'FFFFFF',
    borderColor: '000000',
    widthPx: Math.round(TEXTBOX_WIDTH_EMU / 9525),
    heightPx: Math.round(TEXTBOX_HEIGHT_EMU / 9525),
    paras: [{ runs: [{ text: '' }] }],
  }
}

/** Insert a floating text box (wp:anchor + wps:wsp) at the current cursor. */
export function insertTextboxAt(editor: Editor): void {
  const xml = buildTextboxParagraphXml({
    widthEmu: TEXTBOX_WIDTH_EMU,
    heightEmu: TEXTBOX_HEIGHT_EMU,
    id: Math.floor(Math.random() * 900000) + 100000,
  })
  editor
    .chain()
    .focus()
    .insertContent({
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: t('ribbonTextBox'),
        genXml: xml,
        textboxes: [emptyTextboxDisplay()],
      },
    })
    .run()
}

/** Common DrawingML preset shapes for the shape picker dropdown. */
export const SHAPE_PRESETS: Array<{
  prst: string
  labelKey: StringKey
  clipPath?: string
  borderRadius?: string
}> = [
  { prst: 'rect', labelKey: 'ribbonShapeRect' },
  { prst: 'roundRect', labelKey: 'ribbonShapeRoundRect', borderRadius: '12%' },
  { prst: 'ellipse', labelKey: 'ribbonShapeEllipse', borderRadius: '50%' },
  {
    prst: 'triangle',
    labelKey: 'ribbonShapeTriangle',
    clipPath: 'polygon(50% 0%, 0% 100%, 100% 100%)',
  },
  {
    prst: 'rtTriangle',
    labelKey: 'ribbonShapeRtTriangle',
    clipPath: 'polygon(0% 0%, 0% 100%, 100% 100%)',
  },
  {
    prst: 'diamond',
    labelKey: 'ribbonShapeDiamond',
    clipPath: 'polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)',
  },
  {
    prst: 'parallelogram',
    labelKey: 'ribbonShapeParallelogram',
    clipPath: 'polygon(15% 0%, 100% 0%, 85% 100%, 0% 100%)',
  },
  {
    prst: 'pentagon',
    labelKey: 'ribbonShapePentagon',
    clipPath: 'polygon(50% 0%, 100% 38%, 82% 100%, 18% 100%, 0% 38%)',
  },
  {
    prst: 'hexagon',
    labelKey: 'ribbonShapeHexagon',
    clipPath: 'polygon(25% 0%, 75% 0%, 100% 50%, 75% 100%, 25% 100%, 0% 50%)',
  },
  {
    prst: 'star5',
    labelKey: 'ribbonShapeStar5',
    clipPath:
      'polygon(50% 0%, 61% 35%, 98% 35%, 68% 57%, 79% 91%, 50% 70%, 21% 91%, 32% 57%, 2% 35%, 39% 35%)',
  },
  {
    prst: 'rightArrow',
    labelKey: 'ribbonShapeRightArrow',
    clipPath: 'polygon(0% 25%, 65% 25%, 65% 0%, 100% 50%, 65% 100%, 65% 75%, 0% 75%)',
  },
  {
    prst: 'leftRightArrow',
    labelKey: 'ribbonShapeLeftRightArrow',
    clipPath:
      'polygon(0% 50%, 20% 0%, 20% 30%, 80% 30%, 80% 0%, 100% 50%, 80% 100%, 80% 70%, 20% 70%, 20% 100%)',
  },
]

/** Shape preset CSS: convert a prst name to inline style properties. */
export function shapePresetStyle(prst: string | undefined): CSSProperties {
  const preset = SHAPE_PRESETS.find((s) => s.prst === prst)
  if (!preset) return {}
  const style: CSSProperties = {}
  if (preset.clipPath) style.clipPath = preset.clipPath
  if (preset.borderRadius) style.borderRadius = preset.borderRadius
  return style
}

/**
 * Insert a block beside the current top-level block. Ribbon actions can be
 * invoked while the caret is nested in a table cell, where inserting a block
 * directly at the selection is invalid and TipTap otherwise fails silently.
 */
export function insertTopLevelBlockAtSelection(editor: Editor, content: JSONContent): boolean {
  const { $from } = editor.state.selection
  const position = $from.depth > 0 ? $from.after(1) : editor.state.selection.to
  return editor.chain().focus().insertContentAt(position, content).run()
}

/** Insert a floating preset shape (wps:wsp with prstGeom) at the cursor. */
export function insertShapeAt(editor: Editor, prst: string): void {
  const preset = SHAPE_PRESETS.find((s) => s.prst === prst) ?? SHAPE_PRESETS[0]
  const widthEmu = 1800000
  const heightEmu = 1080000
  const xml = buildShapeParagraphXml({
    prst,
    widthEmu,
    heightEmu,
    id: Math.floor(Math.random() * 900000) + 100000,
    // default Office blue fill + slightly darker border
    fillHex: '4472C4',
    borderHex: '2F5496',
    withTextbox: true,
  })
  const textbox: TextboxDisplay = {
    fill: '4472C4',
    borderColor: '2F5496',
    widthPx: Math.round(widthEmu / 9525),
    heightPx: Math.round(heightEmu / 9525),
    prst,
    paras: [{ runs: [{ text: '' }] }],
  }
  insertTopLevelBlockAtSelection(editor, {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('ribbonShapeLabel', { name: t(preset.labelKey) }),
      genXml: xml,
      textboxes: [textbox],
    },
  })
}

/** ~7.5 cm × 2 cm default size for WordArt in EMU */
const WORDART_WIDTH_EMU = 2700000
const WORDART_HEIGHT_EMU = 720000

/** Insert a floating WordArt text box at the cursor. */
export function insertWordArtAt(editor: Editor, wordArtId: string): void {
  const preset = WORDART_PRESETS.find((p) => p.id === wordArtId) ?? WORDART_PRESETS[0]
  const xml = buildWordArtParagraphXml({
    wordArtId,
    widthEmu: WORDART_WIDTH_EMU,
    heightEmu: WORDART_HEIGHT_EMU,
    id: Math.floor(Math.random() * 900000) + 100000,
  })
  const textbox: TextboxDisplay = {
    // no background fill; shape border is also absent (noFill)
    widthPx: Math.round(WORDART_WIDTH_EMU / 9525),
    heightPx: Math.round(WORDART_HEIGHT_EMU / 9525),
    wordArtId,
    paras: [
      {
        runs: [
          {
            text: t('ribbonWordArtDefaultText'),
            color: preset.colorHex,
            bold: true,
            sizeHalfPoints: 72,
          },
        ],
        align: 'center',
      },
    ],
  }
  editor
    .chain()
    .focus()
    .insertContent({
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: t('ribbonWordArtLabel', { name: preset.label }),
        genXml: xml,
        textboxes: [textbox],
      },
    })
    .run()
}

export function insertPageBreakAt(editor: Editor): void {
  editor
    .chain()
    .focus()
    .insertContent({ type: 'docParagraph', attrs: { pageBreakBefore: true } })
    .run()
}

/**
 * Word's "Blank Page": one empty paragraph that starts its own page, and —
 * only when content follows it — a break pushed onto that following block so
 * it starts the page after. Unconditionally inserting two break paragraphs
 * turned a 1-page document into 3 pages.
 */
export function insertBlankPageAt(editor: Editor): void {
  editor
    .chain()
    .focus()
    .insertContent({ type: 'docParagraph', attrs: { pageBreakBefore: true } })
    .run()
  // locate the paragraph just inserted: the caret block, or the block before
  // it when insertContent left the caret in the split-off remainder
  const { doc, selection } = editor.state
  const isBlankBreak = (index: number): boolean => {
    if (index < 0 || index >= doc.childCount) return false
    const n = doc.child(index)
    return n.type.name === 'docParagraph' && n.childCount === 0 && n.attrs.pageBreakBefore === true
  }
  const caretIndex = selection.$to.index(0)
  const blankIndex = isBlankBreak(caretIndex)
    ? caretIndex
    : isBlankBreak(caretIndex - 1)
      ? caretIndex - 1
      : -1
  if (blankIndex < 0) return
  let blankPos = 0
  for (let i = 0; i < blankIndex; i++) blankPos += doc.child(i).nodeSize
  // Word leaves the caret on the new blank page
  editor.commands.setTextSelection(blankPos + 1)
  const followIndex = blankIndex + 1
  if (followIndex >= doc.childCount) return // blank page is the last page
  const follow = doc.child(followIndex)
  const followPos = blankPos + doc.child(blankIndex).nodeSize
  if ('pageBreakBefore' in follow.attrs) {
    if (!follow.attrs.pageBreakBefore) {
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(followPos, undefined, {
          ...follow.attrs,
          pageBreakBefore: true,
        }),
      )
    }
    return
  }
  // following block can't carry the property (e.g. a table): spacer fallback
  const spacer = editor.state.schema.nodes.docParagraph!.createAndFill({ pageBreakBefore: true })
  if (spacer) editor.view.dispatch(editor.state.tr.insert(followPos, spacer))
}

export interface InsertTabProps extends TabProps {
  header: HeaderFooter | null
  onHeader: (next: HeaderFooter) => void
  footer: HeaderFooter | null
  onFooter: (next: HeaderFooter) => void
  /** Open the "Page Number Format" dialog (number format / start-at, writes sectPr w:pgNumType) */
  onPageNumFormat: () => void
  /** Insert an inline field (DATE/TIME/PAGE/NUMPAGES/FILENAME) */
  onInsertField: (instr: string) => void
  titlePg: boolean
  onTitlePg: (v: boolean) => void
  evenOddHf: boolean
  onEvenOddHf: (v: boolean) => void
  commentCount: number
  onShowComments: () => void
}

/** target languages of Word's Translate dropdown that the AI backend can serve;
 *  the localized label is also spliced into the instruction sent to the LLM */
const TRANSLATE_TARGETS: Array<{ labelKey: StringKey }> = [
  { labelKey: 'ribbonLangEnglish' },
  { labelKey: 'ribbonLangSimplifiedChinese' },
  { labelKey: 'ribbonLangJapanese' },
  { labelKey: 'ribbonLangKorean' },
  { labelKey: 'ribbonLangFrench' },
  { labelKey: 'ribbonLangGerman' },
  { labelKey: 'ribbonLangSpanish' },
]

/** One-time "AI rewrites the whole document" acknowledgement */
const AI_REWRITE_ACK_KEY = 'docs-ai-rewrite-ack'

/** Revision display modes: All Markup (default) / No Markup (as accepted) / Original (as rejected) */
export type RevisionDisplayMode = 'all' | 'none' | 'original'

interface ReviewTabProps extends TabProps {
  onAiPreset: (instruction: string) => void
  commentCount: number
  onShowComments: () => void
  /** create a comment on the current selection (disabled when selection is empty) */
  canComment: boolean
  onNewComment: () => void
  trackChanges: boolean
  onTrackChanges: (on: boolean) => void
  revisionDisplay: RevisionDisplayMode
  onRevisionDisplay: (mode: RevisionDisplayMode) => void
  revisionCount: number
  onAcceptRevision: (all: boolean) => void
  onRejectRevision: (all: boolean) => void
  onGotoRevision: (dir: 1 | -1) => void
  /** Restrict Editing (read-only) is enforced */
  isProtected: boolean
  onToggleProtection: () => void
  onCompare: () => void
}

export function ReviewTab({
  hasDoc,
  dropdown,
  setDropdown,
  onAiPreset,
  commentCount,
  onShowComments,
  canComment,
  onNewComment,
  trackChanges,
  onTrackChanges,
  revisionDisplay,
  onRevisionDisplay,
  revisionCount,
  onAcceptRevision,
  onRejectRevision,
  onGotoRevision,
  isProtected,
  onToggleProtection,
  onCompare,
}: ReviewTabProps) {
  const { t } = useI18n()
  // One-time acknowledgement before whole-document AI rewrites:
  // Editor / Translate send the full document to the agent, consume credits and
  // may rewrite everything — say so once before the first run.
  const confirmAiRewrite = () => {
    if (localStorage.getItem(AI_REWRITE_ACK_KEY) === '1') return true
    if (!window.confirm(t('ribbonAiRewriteConfirm'))) return false
    localStorage.setItem(AI_REWRITE_ACK_KEY, '1')
    return true
  }
  return (
    <>
      {/* Word: Proofing (Editor) sits leftmost */}
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={`${t('ribbonEditorTip')} — ${t('ribbonAiCreditNote')}`}
            onClick={() => {
              if (confirmAiRewrite()) onAiPreset(t('ribbonEditorPrompt'))
            }}
          >
            <span className="rb-big-icon">
              <span className="copilot-badge">
                <IconSparkle size={13} />
              </span>
            </span>
            <span>{t('ribbonEditorBtn')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupProofing')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              title={`${t('ribbonTranslateTip')} — ${t('ribbonAiCreditNote')}`}
              onClick={() => toggleDropdown(setDropdown, 'translate')}
            >
              <span className="rb-big-icon">
                <span className="copilot-badge">
                  <IconTranslate size={13} />
                </span>
                <IconCaret />
              </span>
              <span>{t('ribbonTranslate')}</span>
            </button>
            {dropdown === 'translate' && (
              <div className="layout-menu">
                {TRANSLATE_TARGETS.map((lang) => (
                  <button
                    key={lang.labelKey}
                    onClick={() => {
                      setDropdown(() => null)
                      if (confirmAiRewrite()) {
                        onAiPreset(t('ribbonTranslatePrompt', { lang: t(lang.labelKey) }))
                      }
                    }}
                  >
                    {t('ribbonTranslateTo', { lang: t(lang.labelKey) })}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupLanguage')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc || !canComment || isProtected}
            title={canComment ? t('ribbonNewCommentTip') : t('ribbonNewCommentSelectTip')}
            onClick={onNewComment}
          >
            <span className="rb-big-icon">
              <IconComment size={BIG} />
            </span>
            <span>{t('ribbonNewComment')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonShowCommentsTip', { count: commentCount })}
            onClick={onShowComments}
          >
            <span className="rb-big-icon">
              <IconComment size={BIG} />
            </span>
            <span>{t('ribbonShowComments')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupComments')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className={`rb-big ${trackChanges ? 'active' : ''}`}
            disabled={!hasDoc || isProtected}
            title={t('ribbonTrackChangesTip')}
            onClick={() => onTrackChanges(!trackChanges)}
          >
            <span className="rb-big-icon">
              <IconTrackChanges size={BIG} />
            </span>
            <span>{t('ribbonTrackChanges')}</span>
          </button>
          <div className="rb-split-wrap">
            <button
              className={`rb-big ${revisionDisplay !== 'all' ? 'active' : ''}`}
              disabled={!hasDoc}
              title={t('ribbonRevDisplayTip')}
              onClick={() => toggleDropdown(setDropdown, 'revDisplay')}
            >
              <span className="rb-big-icon">
                <IconReadMode size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonRevDisplay')}</span>
            </button>
            {dropdown === 'revDisplay' && (
              <div className="layout-menu">
                {(
                  [
                    ['all', t('ribbonRevDisplayAll')],
                    ['none', t('ribbonRevDisplayNone')],
                    ['original', t('ribbonRevDisplayOriginal')],
                  ] as Array<[RevisionDisplayMode, string]>
                ).map(([mode, label]) => (
                  <button
                    key={mode}
                    onClick={() => {
                      onRevisionDisplay(mode)
                      setDropdown(() => null)
                    }}
                  >
                    {label}
                    {revisionDisplay === mode ? ' ✓' : ''}
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc || revisionCount === 0 || isProtected}
              title={t('ribbonAcceptTip', { count: revisionCount })}
              onClick={() => toggleDropdown(setDropdown, 'acceptRev')}
            >
              <span className="rb-big-icon">
                <IconAccept size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonAccept')}</span>
            </button>
            {dropdown === 'acceptRev' && (
              <div className="layout-menu">
                <button
                  onClick={() => {
                    onAcceptRevision(false)
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonAcceptOne')}
                </button>
                <button
                  onClick={() => {
                    onAcceptRevision(true)
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonAcceptAll')}
                </button>
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc || revisionCount === 0 || isProtected}
              title={t('ribbonRejectTip', { count: revisionCount })}
              onClick={() => toggleDropdown(setDropdown, 'rejectRev')}
            >
              <span className="rb-big-icon">
                <IconReject size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonReject')}</span>
            </button>
            {dropdown === 'rejectRev' && (
              <div className="layout-menu">
                <button
                  onClick={() => {
                    onRejectRevision(false)
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonRejectOne')}
                </button>
                <button
                  onClick={() => {
                    onRejectRevision(true)
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonRejectAll')}
                </button>
              </div>
            )}
          </div>
          <button
            className="rb-big"
            disabled={!hasDoc || revisionCount === 0}
            title={t('ribbonPrevChangeTip')}
            onClick={() => onGotoRevision(-1)}
          >
            <span className="rb-big-icon">
              <IconUndo size={BIG} />
            </span>
            <span>{t('ribbonPrevChange')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc || revisionCount === 0}
            title={t('ribbonNextChangeTip')}
            onClick={() => onGotoRevision(1)}
          >
            <span className="rb-big-icon">
              <IconRedo size={BIG} />
            </span>
            <span>{t('ribbonNextChange')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupTracking')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonCompareTip')}
            onClick={onCompare}
          >
            <span className="rb-big-icon">
              <IconCompare size={BIG} />
            </span>
            <span>{t('ribbonCompare')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonCompare')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className={`rb-big ${isProtected ? 'active' : ''}`}
            disabled={!hasDoc}
            title={isProtected ? t('ribbonStopProtectionTip') : t('ribbonRestrictEditingTip')}
            onClick={onToggleProtection}
          >
            <span className="rb-big-icon">
              <IconLock size={BIG} />
            </span>
            <span>{t('ribbonRestrictEditing')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupProtect')}</div>
      </div>
    </>
  )
}

/* ================= View ================= */

/** document rendering modes: Print Layout / Web Layout / Outline */
export type ViewMode = 'print' | 'web' | 'outline'

interface ViewTabProps {
  hasDoc: boolean
  /** current document path, so New Window can open the same file */
  filePath: string | null
  zoom: number
  onZoom: (zoom: number) => void
  onZoomFit: (mode: 'width' | 'page') => void
  showAi: boolean
  onToggleAi: () => void
  darkCanvas: boolean
  onDarkCanvas: (v: boolean) => void
  showRuler: boolean
  onShowRuler: (v: boolean) => void
  showNav: boolean
  onShowNav: (v: boolean) => void
  viewMode: ViewMode
  onViewMode: (mode: ViewMode) => void
  readMode: boolean
  onReadMode: (v: boolean) => void
  showGrid: boolean
  onShowGrid: (v: boolean) => void
  splitView: boolean
  onSplitView: (v: boolean) => void
  onPagePreview: () => void
}

export function ViewTab({
  hasDoc,
  filePath,
  zoom,
  onZoom,
  onZoomFit,
  showAi,
  onToggleAi,
  darkCanvas,
  onDarkCanvas,
  showRuler,
  onShowRuler,
  showNav,
  onShowNav,
  viewMode,
  onViewMode,
  readMode,
  onReadMode,
  showGrid,
  onShowGrid,
  splitView,
  onSplitView,
  onPagePreview,
}: ViewTabProps) {
  const { t } = useI18n()
  const [winMenuOpen, setWinMenuOpen] = useState(false)
  const [windows, setWindows] = useState<DocsTabInfo[]>([])

  const toggleWinMenu = async () => {
    if (!winMenuOpen) setWindows(await window.desktop.listDocsTabs())
    setWinMenuOpen((v) => !v)
  }

  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className={`rb-big ${viewMode === 'print' && !readMode ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonPrintLayoutTip')}
            onClick={() => {
              onViewMode('print')
              onReadMode(false)
            }}
          >
            <span className="rb-big-icon">
              <IconPrintLayout size={BIG} />
            </span>
            <span>{t('ribbonPrintLayout')}</span>
          </button>
          <button
            className={`rb-big ${viewMode === 'web' ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonWebLayoutTip')}
            onClick={() => onViewMode(viewMode === 'web' ? 'print' : 'web')}
          >
            <span className="rb-big-icon">
              <IconWebLayout size={BIG} />
            </span>
            <span>{t('ribbonWebLayout')}</span>
          </button>
          <button
            className={`rb-big ${viewMode === 'outline' ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonOutlineViewTip')}
            onClick={() => onViewMode(viewMode === 'outline' ? 'print' : 'outline')}
          >
            <span className="rb-big-icon">
              <IconOutlineView size={BIG} />
            </span>
            <span>{t('ribbonOutlineView')}</span>
          </button>
          <button
            className={`rb-big ${readMode ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonReadModeTip')}
            onClick={() => onReadMode(!readMode)}
          >
            <span className="rb-big-icon">
              <IconReadMode size={BIG} />
            </span>
            <span>{t('ribbonReadMode')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc || viewMode !== 'print' || readMode}
            title={t('ribbonPagePreviewTip')}
            onClick={onPagePreview}
          >
            <span className="rb-big-icon">
              <IconWholePage size={BIG} />
            </span>
            <span>{t('ribbonPagePreview')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupViews')}</div>
      </div>

      <div className="ribbon-sep" />
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonZoomOut')}
            onClick={() => onZoom(Math.max(50, zoom - 10))}
          >
            <span className="rb-big-icon">
              <IconZoomOut size={BIG} />
            </span>
            <span>{t('ribbonZoomOut')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonZoomIn')}
            onClick={() => onZoom(Math.min(200, zoom + 10))}
          >
            <span className="rb-big-icon">
              <IconZoomIn size={BIG} />
            </span>
            <span>{t('ribbonZoomIn')}</span>
          </button>
          <button
            className={`rb-big ${zoom === 100 ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonZoom100Tip')}
            onClick={() => onZoom(100)}
          >
            <span className="rb-big-icon">
              <IconZoom100 size={BIG} />
            </span>
            <span>100%</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonPageWidthTip')}
            onClick={() => onZoomFit('width')}
          >
            <span className="rb-big-icon">
              <IconPageWidth size={BIG} />
            </span>
            <span>{t('ribbonPageWidth')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonWholePageTip')}
            onClick={() => onZoomFit('page')}
          >
            <span className="rb-big-icon">
              <IconWholePage size={BIG} />
            </span>
            <span>{t('ribbonWholePage')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupZoom')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className={`rb-big ${showAi ? 'active' : ''}`}
            title={t('ribbonAiPanelTip')}
            onClick={onToggleAi}
          >
            <span className="rb-big-icon">
              <IconAiPanel size={BIG} />
            </span>
            <span>{t('ribbonAiPanel')}</span>
          </button>
          <button
            className={`rb-big ${darkCanvas ? 'active' : ''}`}
            title={t('ribbonDarkModeTip')}
            onClick={() => onDarkCanvas(!darkCanvas)}
          >
            <span className="rb-big-icon">
              <IconMoon size={BIG} />
            </span>
            <span>{t('ribbonDarkMode')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupAppearance')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className={`rb-big ${showRuler ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonRulerTip')}
            onClick={() => onShowRuler(!showRuler)}
          >
            <span className="rb-big-icon">
              <IconRuler size={BIG} />
            </span>
            <span>{t('ribbonRuler')}</span>
          </button>
          <button
            className={`rb-big ${showGrid ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonGridlinesTip')}
            onClick={() => onShowGrid(!showGrid)}
          >
            <span className="rb-big-icon">
              <IconGridlines size={BIG} />
            </span>
            <span>{t('ribbonGridlines')}</span>
          </button>
          <button
            className={`rb-big ${showNav ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonNavPaneTip')}
            onClick={() => onShowNav(!showNav)}
          >
            <span className="rb-big-icon">
              <IconNavPane size={BIG} />
            </span>
            <span>{t('ribbonNavPane')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupShow')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            title={t('ribbonNewTabTip')}
            onClick={() => void window.desktop.openNewTab(filePath)}
          >
            <span className="rb-big-icon">
              <IconNewWindow size={BIG} />
            </span>
            <span>{t('ribbonNewTab')}</span>
          </button>
          <button
            className={`rb-big ${splitView ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonSplitTip')}
            onClick={() => onSplitView(!splitView)}
          >
            <span className="rb-big-icon">
              <IconSplit size={BIG} />
            </span>
            <span>{t('ribbonSplit')}</span>
          </button>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              title={t('ribbonSwitchTabsTip')}
              onClick={() => void toggleWinMenu()}
            >
              <span className="rb-big-icon">
                <IconSwitchWindows size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonSwitchTabs')}</span>
            </button>
            {winMenuOpen && (
              <div className="layout-menu align-right">
                {windows.map((w) => (
                  <button
                    key={w.id}
                    onClick={() => {
                      void window.desktop.focusDocsTab(w.id)
                      setWinMenuOpen(false)
                    }}
                  >
                    {w.focused ? '✓ ' : ''}
                    {w.title || 'GenOffice Docs'}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupWindow')}</div>
      </div>
    </>
  )
}
/* ================= Draw ================= */

/** per-pen settings the Draw tab remembers (Word keeps pen/highlighter separate) */
export interface InkPenSettings {
  /** hex without '#' */
  color: string
  width: number
}

const INK_COLORS = [
  '000000',
  'C00000',
  'FF0000',
  'FFC000',
  'FFFF00',
  '92D050',
  '00B050',
  '00B0F0',
  '0070C0',
  '7030A0',
]

const PEN_WIDTHS = [1, 2, 3.5, 5]
const HIGHLIGHTER_WIDTHS = [6, 10, 16]

interface DrawTabProps {
  hasDoc: boolean
  tool: InkTool
  onTool: (tool: InkTool) => void
  pen: InkPenSettings
  onPen: (settings: InkPenSettings) => void
  highlighter: InkPenSettings
  onHighlighter: (settings: InkPenSettings) => void
  annotationCount: number
  onClearAll: () => void
}

export function DrawTab({
  hasDoc,
  tool,
  onTool,
  pen,
  onPen,
  highlighter,
  onHighlighter,
  annotationCount,
  onClearAll,
}: DrawTabProps) {
  const { t } = useI18n()
  // color/thickness edit the active pen; with no pen active they configure the pen tool
  const editingHighlighter = tool === 'highlighter'
  const active = editingHighlighter ? highlighter : pen
  const setActive = editingHighlighter ? onHighlighter : onPen
  const widths = editingHighlighter ? HIGHLIGHTER_WIDTHS : PEN_WIDTHS

  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className={`rb-big ${tool === 'select' ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonSelectTip')}
            onClick={() => onTool('select')}
          >
            <span className="rb-big-icon">
              <IconCursor size={BIG} />
            </span>
            <span>{t('ribbonSelect')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonSelect')}</div>
      </div>
      <div className="ribbon-sep" />
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className={`rb-big ${tool === 'pen' ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonPenTip')}
            onClick={() => onTool('pen')}
          >
            <span className="rb-big-icon" style={{ color: `#${pen.color}` }}>
              <IconPen size={BIG} />
            </span>
            <span>{t('ribbonPen')}</span>
          </button>
          <button
            className={`rb-big ${tool === 'highlighter' ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonHighlighterTip')}
            onClick={() => onTool('highlighter')}
          >
            <span className="rb-big-icon" style={{ color: `#${highlighter.color}` }}>
              <IconHighlighterPen size={BIG} />
            </span>
            <span>{t('ribbonHighlighter')}</span>
          </button>
          <button
            className={`rb-big ${tool === 'eraser' ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonEraserTip')}
            onClick={() => onTool('eraser')}
          >
            <span className="rb-big-icon">
              <IconEraser size={BIG} />
            </span>
            <span>{t('ribbonEraser')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupDrawingTools')}</div>
      </div>
      <div className="ribbon-sep" />
      <div className="ribbon-group">
        <div className="ribbon-group-items ink-settings">
          <div className="ink-swatches">
            {INK_COLORS.map((hex) => (
              <button
                key={hex}
                className={`ink-swatch ${active.color === hex ? 'active' : ''}`}
                style={{ background: `#${hex}` }}
                title={`#${hex}`}
                disabled={!hasDoc}
                onClick={() => setActive({ ...active, color: hex })}
              />
            ))}
          </div>
          <div className="ink-widths">
            {widths.map((w) => (
              <button
                key={w}
                className={`ink-width ${active.width === w ? 'active' : ''}`}
                title={t('ribbonPixels', { w })}
                disabled={!hasDoc}
                onClick={() => setActive({ ...active, width: w })}
              >
                <span
                  className="ink-width-dot"
                  style={{
                    width: Math.min(16, w * 2 + 2),
                    height: Math.min(16, w * 2 + 2),
                    background: `#${active.color}`,
                  }}
                />
              </button>
            ))}
          </div>
        </div>
        <div className="ribbon-group-label">
          {editingHighlighter ? t('ribbonHighlighterStyle') : t('ribbonPenStyle')}
        </div>
      </div>
      <div className="ribbon-sep" />
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc || annotationCount === 0}
            title={t('ribbonClearAllTip')}
            onClick={onClearAll}
          >
            <span className="rb-big-icon">
              <IconEraser size={BIG} />
            </span>
            <span>{t('ribbonClearAll')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupClear')}</div>
      </div>
    </>
  )
}
