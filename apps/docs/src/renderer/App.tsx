import { useCallback, useEffect, useMemo, useReducer, useRef, useState } from 'react'
import type { CSSProperties, MouseEvent as ReactMouseEvent } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import type { Editor } from '@tiptap/core'
import { DOMParser as PmDOMParser, type Mark as PmMark } from '@tiptap/pm/model'
import { markdownPasteHtml } from './editor/markdown-paste'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  DEFAULT_SECTION,
  applySectionSettings,
  type Block,
  type CommentInfo,
  type CustomNumberingLevel,
  type DocProtection,
  type HeaderFooter,
  type HfImage,
  type NoteInfo,
  type SectionInfo,
  type SectionSettings,
  type SourceInfo,
  type StyleUpsert,
  type ThemeColors,
  type ThemeFonts,
} from '@genoffice/docx-engine'
import type { AiSettings, OpenFileResult } from '../shared/ipc'
import { AI_PROVIDERS } from '../shared/ipc'
import { AiPanel } from './ai/AiPanel'
import { asianCharCount, countWords, nonAsianWordCount } from './word-count'
import { toRoman } from './note-format'
import { CommentsPanel } from './components/CommentsPanel'
import { EquationModal } from './components/EquationModal'
import { HeaderFooterArea } from './components/HeaderFooterArea'
import { PaginationPreview } from './components/PaginationPreview'
import {
  appendEndnotesBlock,
  assignSections,
  effectiveHfRefs,
  lineStartAnchor,
  liveSections,
  nextLineAnchor,
  measureBlocks,
  docGridPitchPt,
  type LineAnchor,
  pageNumbers,
  sliceWithLineSplit,
  tableRowFlags,
  type BlockBox,
  type BlockMeta,
  type PageNoteItem,
  pageAt,
  singleCutCell,
  sectionColGeom,
  sectionColumns,
  sectionFirstPages,
  sectionGeoms,
  sectionPageBox,
  effectiveTopPx,
  effectiveBottomPx,
  formatPageNumber,
  visiblePageCount,
  type SectionGeom,
  type SectionHfHeights,
  type PageSlice,
} from './pagination'
import { GAP_BAND, setPageGaps, type PageGapSpec } from './editor/pagination-gaps'
import { hfHasVisibleContent, makeGapHfEl } from './editor/hf-dom'
import {
  estimateFootnoteHeight,
  estimateHfHeight,
  FOOTNOTE_SEPARATOR_H,
  textHasCjk,
} from './line-metrics'
import { saveUntilPersisted } from './save-until-persisted'
import { cachedByDoc } from './doc-cache'
import { useShallowStable, useStableCallbacks } from './use-stable'
import { FindPanel } from './components/FindPanel'
import { Ribbon } from './components/Ribbon'
import { computeFormatState } from './components/ribbon-format-state'
import { IconRedo, IconSave, IconUndo } from './components/icons'
import { ToastHost } from './components/toast'
import {
  LinkInsertModal,
  insertImageFromDataUrl,
  insertImageViaDialog,
  insertPageBreakAt,
  insertTableAt,
  type InkPenSettings,
  type RevisionDisplayMode,
  type ViewMode,
} from './components/ribbon-tabs'
import { ComparePanel } from './components/ComparePanel'
import {
  EditorContextMenu,
  FontDialog,
  ParagraphDialog,
  type ContextMenuState,
} from './components/ContextMenu'
import { PromptModal } from './components/PromptModal'
import { t, useI18n } from './i18n/locale'
import {
  getActiveSubEditor,
  setActiveSubEditor,
  subscribeSubEditorState,
} from './editor/active-editor'
import type { CompareEntry } from './editor/compare'
import { collectHeadings } from './editor/headings'
import { applyTocPageDisplays } from './editor/toc-refresh'
import { setSelectionAlign } from './editor/direction'

import {
  editorExtensions,
  resolvedCommentsPluginKey,
  revisionDisplayState,
} from './editor/extensions'
import { type InkAnnotation, type InkTool } from './editor/ink'
import { InkOverlay } from './components/InkOverlay'
import { collectRevisions, gotoRevision, type TrackChangesStorage } from './editor/revisions'
import { NavPane } from './components/NavPane'
import { Ruler } from './components/Ruler'
import { docBodyFont, docLineFactor, docThemeCss } from './doc-style-css'
import { isDocDirty } from './doc-dirty'
import {
  EMPTY_HF_VARIANTS,
  hfFromPart,
  restingHfAreaVariant,
  type DocState,
  type HfVariantKey,
  type HfVariantsState,
  type HfView,
  type PendingNumbering,
} from './doc-state'
import {
  exportPdf as exportPdfImpl,
  loadFile as loadFileImpl,
  newFile as newFileImpl,
  save as saveImpl,
  writeRecoveryCopy as writeRecoveryCopyImpl,
  type FileActionContext,
} from './file-actions'
import {
  allocateListNumId as allocateListNumIdImpl,
  continueNumbering as continueNumberingImpl,
  createCustomListDef as createCustomListDefImpl,
  restartNumbering as restartNumberingImpl,
  type NumberingContext,
} from './numbering-actions'
import {
  addInk as addInkImpl,
  cancelNewComment as cancelNewCommentImpl,
  clearInks as clearInksImpl,
  compareWithFile as compareWithFileImpl,
  deleteComment as deleteCommentImpl,
  deleteNote as deleteNoteImpl,
  handleRevision as handleRevisionImpl,
  removeInks as removeInksImpl,
  replyToComment as replyToCommentImpl,
  resolveComment as resolveCommentImpl,
  startNewComment as startNewCommentImpl,
  submitNewComment as submitNewCommentImpl,
  submitNote as submitNoteImpl,
  submitProtectModal as submitProtectModalImpl,
  toggleProtection as toggleProtectionImpl,
  type NotePrompt,
  type ProtectModalState,
  type ReviewContext,
} from './review-actions'

const _IS_MAC = navigator.platform.toLowerCase().includes('mac')

const twipsToPx = (twips: number) => (twips / 1440) * 96

/** tiny stable string hash for decoration keys */
function hashStr(str: string): number {
  let h = 0
  for (let i = 0; i < str.length; i++) h = (h * 31 + str.charCodeAt(i)) | 0
  return h
}

const EMPTY_BLOCKS: Block[] = []

// O(doc) derivations cached by PM doc reference: caret moves and unrelated
// state updates reuse the last result instead of re-walking the whole document
const wordCountOfDoc = cachedByDoc((d) => countWords(d.textContent))
const revisionCountOfDoc = cachedByDoc((d) => collectRevisions(d).length)

/**
 * Document position of a measured line-start DOM anchor. posAtDOM works from the DOM
 * tree, unlike posAtCoords whose viewport hit-testing returns degenerate positions for
 * lines scrolled off-screen (which misplaced page-break markers into table bodies,
 * inflating the canvas table by a phantom anonymous row and skewing pagination).
 */
function posFromAnchor(view: Editor['view'], anchor: LineAnchor): number | undefined {
  try {
    const pos = view.posAtDOM(anchor.node, anchor.charOffset)
    return pos >= 0 ? pos : undefined
  } catch {
    return undefined
  }
}

/** Clean pasted Word/web HTML: mso conditional comments, <o:p>, and unwrapping <li><p>x</p></li> */
function cleanPastedHtml(html: string): string {
  return html
    .replace(/<!--\[if[\s\S]*?<!\[endif\]-->/g, '')
    .replace(/<o:p>[\s\S]*?<\/o:p>/g, '')
    .replace(/<li([^>]*)>\s*<p[^>]*>([\s\S]*?)<\/p>\s*<\/li>/g, '<li$1>$2</li>')
}

/** All runs a footnote/endnote reference may live in: paragraph runs, plus table
 *  cell paragraphs (incl. nested tables) — refs inside cells still print their
 *  note at the page bottom like Word */
function blockNoteScanRuns(b: Block): NonNullable<Block['runs']> {
  const out: NonNullable<Block['runs']> = []
  if (b.runs) out.push(...b.runs)
  const walkTable = (table: NonNullable<Block['table']> | undefined): void => {
    for (const row of table?.rows ?? []) {
      for (const cell of row) {
        for (const p of cell.richParas ?? []) out.push(...p.runs)
        for (const nested of cell.nestedTables ?? []) walkTable(nested)
      }
    }
  }
  if (b.table) walkTable(b.table)
  return out
}

/** Footnote area at the top of a page gap (previous page's bottom): absolutely positioned in the content area, double-click an entry to edit */
function makeGapNotesEl(
  items: PageNoteItem[],
  leftPx: number,
  widthPx: number,
  heightPx: number,
  onEdit: (id: string) => void,
): HTMLElement {
  const wrap = document.createElement('div')
  wrap.className = 'page-gap-notes'
  wrap.style.left = `${leftPx}px`
  wrap.style.width = `${widthPx}px`
  wrap.style.height = `${heightPx}px`
  for (const it of items) {
    const row = document.createElement('div')
    row.className = 'page-gap-note'
    row.title = t('appDblclickEditFootnote')
    // entries get fixed heights from the estimates: rendering matches the engine's reservation bookkeeping exactly, no more overflow
    row.style.height = `${it.height}px`
    const sup = document.createElement('sup')
    sup.textContent = String(it.no)
    row.append(sup)
    if (it.richParas) {
      it.richParas.forEach((para, pi) => {
        if (pi > 0) row.append(document.createElement('br'))
        for (const run of para) {
          const span = document.createElement('span')
          span.textContent = run.text
          if (run.bold) span.style.fontWeight = '600'
          if (run.italic) span.style.fontStyle = 'italic'
          const deco = [run.underline && 'underline', run.strike && 'line-through'].filter(Boolean)
          if (deco.length > 0) span.style.textDecoration = deco.join(' ')
          if (run.color) span.style.color = `#${run.color}`
          if (run.sizeHalfPoints) span.style.fontSize = `${run.sizeHalfPoints / 2}pt`
          row.append(span)
        }
      })
    } else {
      const span = document.createElement('span')
      span.textContent = it.text
      row.append(span)
    }
    row.addEventListener('dblclick', () => onEdit(it.id))
    wrap.appendChild(row)
  }
  return wrap
}

/** fields of Word's Word Count dialog */
interface DocStats {
  pages: number
  words: number
  asianChars: number
  nonAsianWords: number
  charsNoSpace: number
  charsWithSpace: number
  paragraphs: number
  lines: number
}

const DEFAULT_SETTINGS: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((p) => [
      p.id,
      { apiKey: '', model: p.defaultModel, baseUrl: p.needsBaseUrl ? '' : undefined },
    ]),
  ) as AiSettings['providers'],
}

export function App() {
  // subscribe to language switches for re-render; strings all go through module-level t, so memoized callbacks never capture stale closures
  const { lang } = useI18n()
  const [doc, setDoc] = useState<DocState | null>(null)
  /** true until the pending-open / new-blank boot checks settle; the start screen stays hidden meanwhile */
  const bootPendingRef = useRef<Promise<[OpenFileResult | null, boolean]> | null>(null)
  const bootHandledRef = useRef(false)
  const [_recent, setRecent] = useState<string[]>([])
  const [settings, setSettings] = useState<AiSettings>(DEFAULT_SETTINGS)
  const [showAi, setShowAi] = useState(() => localStorage.getItem('aidocs.showAi') !== '0')
  /** Increments on every open/new document: AiPanel remounts by key to reset the conversation and history (save path changes don't bump it, so the session continues) */
  const [aiPanelKey, setAiPanelKey] = useState(0)
  const [ribbonTabRequest, setRibbonTabRequest] = useState<{ tab: string; nonce: number } | null>(
    null,
  )
  const [status, setStatus] = useState('')
  const [zoom, setZoom] = useState(100)
  const [darkCanvas, setDarkCanvas] = useState(false)
  const [section, setSection] = useState<SectionSettings | null>(null)
  /** All sections (readSections): pagination/preview use per-section geometry; layout edits apply to the cursor's section */
  const [sections, setSections] = useState<SectionInfo[]>([])
  const [sectionDirty, setSectionDirty] = useState(false)
  /** Index of the cursor's section (maintained by selectionUpdate) */
  const [activeSection, setActiveSection] = useState(0)
  /** Indexes of edited non-final sections (their section-break paragraphs' sectPr is rewritten on save) */
  const [sectionsDirty, setSectionsDirty] = useState<number[]>([])
  /** Start type pending write to the trailing sectPr after inserting a continuous section break */
  const [trailingStartType, setTrailingStartType] = useState<SectionInfo['startType'] | null>(null)
  const [pageColor, setPageColor] = useState<string | null>(null)
  const [pageColorDirty, setPageColorDirty] = useState(false)
  const [header, setHeader] = useState<HeaderFooter | null>(null)
  const [headerDirty, setHeaderDirty] = useState(false)
  const [footer, setFooter] = useState<HeaderFooter | null>(null)
  const [footerDirty, setFooterDirty] = useState(false)
  const [hfVariants, setHfVariants] = useState<HfVariantsState>(EMPTY_HF_VARIANTS)
  const [hfVariantsDirty, setHfVariantsDirty] = useState<HfVariantKey[]>([])
  const [titlePg, setTitlePg] = useState(false)
  const [titlePgDirty, setTitlePgDirty] = useState(false)
  const [evenOddHf, setEvenOddHf] = useState(false)
  const [evenOddHfDirty, setEvenOddHfDirty] = useState(false)
  const [hfView, setHfViewState] = useState<HfView>('default')
  /** false until the user picks a variant (chip/toggle); resting areas then follow their page's variant instead */
  const [hfViewTouched, setHfViewTouched] = useState(false)
  const setHfView = (v: HfView) => {
    setHfViewState(v)
    setHfViewTouched(true)
  }
  const [pageInfo, setPageInfo] = useState({ current: 1, total: 1 })
  // last page's number for the document-end footer: text for the page marker, num for
  // even/odd parity (section restarts / pageNumberFmt make both differ from the physical count)
  const [lastPageNo, setLastPageNo] = useState<{ text: string; num: number } | null>(null)
  /** Page-number-format dialog + pending final-section pgNumType write (non-final sections rewrite sectPr via pgNumDirtySections) */
  const [pgNumModal, setPgNumModal] = useState<{ fmt: string; start: string } | null>(null)
  const [pgNumEdit, setPgNumEdit] = useState<{ fmt?: string; start?: number } | null>(null)
  const [pgNumDirtySections, setPgNumDirtySections] = useState<number[]>([])
  /** Pending numbering definitions to append (saved via SaveOptions.numbering; restart numbering / new list definitions) */
  const [pendingNumbering, setPendingNumbering] = useState<PendingNumbering>({
    newDefs: [],
    restartNums: [],
  })
  const numberingDirty =
    pendingNumbering.newDefs.length > 0 || pendingNumbering.restartNums.length > 0
  /** Header/footer edits for non-final sections (key = `${lastBlockIndex}:${kind}`), saved via SaveOptions.sectionHf */
  const [sectionHfEdits, setSectionHfEdits] = useState<Record<string, HeaderFooter>>({})
  /** The canvas header/footer area always views/edits the first section (multi-section docs: later sections inherit its refs) */
  const hfSectionClamped = 0
  const multiHf = sections.length > 1
  const effRefsAll = useMemo(() => effectiveHfRefs(sections), [sections])
  const effHfView: HfView =
    (hfView === 'first' && !titlePg) || (hfView === 'even' && !evenOddHf) ? 'default' : hfView
  /** Multi-section: the selected section's effective header/footer (local edits > referenced part > final section falls back to global edit state) */
  const sectionHfValue = (kind: 'header' | 'footer'): HeaderFooter | null => {
    const idx = hfSectionClamped
    const sec = sections[idx]
    const globalState = kind === 'header' ? header : footer
    if (!sec) return globalState
    const refVariants = effRefsAll[idx]?.[kind]
    const fromPart = () => {
      const rId = refVariants?.default
      return rId ? hfFromPart(doc?.parsed.hfParts?.[rId]) : null
    }
    if (idx === sections.length - 1) {
      // final section: reuse the global edit state (save writes the final section's sectPr reference)
      const dirty = kind === 'header' ? headerDirty : footerDirty
      return dirty ? globalState : (fromPart() ?? globalState)
    }
    return sectionHfEdits[`${sec.lastBlockIndex}:${kind}`] ?? fromPart()
  }
  /** Variant an on-canvas area shows/edits: explicit chip choice wins; at rest the header area is page 1 and the footer area the last page */
  const areaView = (kind: 'header' | 'footer'): HfView =>
    hfViewTouched
      ? effHfView
      : restingHfAreaVariant(kind, {
          titlePg,
          evenOddHf,
          pageCount: pageInfo.total,
          ...(lastPageNo ? { lastPageNo: lastPageNo.num } : {}),
        })
  const headerAreaView = areaView('header')
  const footerAreaView = areaView('footer')
  const shownHeader =
    headerAreaView === 'first'
      ? hfVariants.headerFirst
      : headerAreaView === 'even'
        ? hfVariants.headerEven
        : multiHf
          ? sectionHfValue('header')
          : header
  const shownFooter =
    footerAreaView === 'first'
      ? hfVariants.footerFirst
      : footerAreaView === 'even'
        ? hfVariants.footerEven
        : multiHf
          ? sectionHfValue('footer')
          : footer
  /** images of the part in the current view (logos etc., display-only; the save path keeps their bytes untouched) */
  const hfImagesOf = (kind: 'header' | 'footer') => {
    const parsed = doc?.parsed
    if (!parsed) return undefined
    const view = kind === 'header' ? headerAreaView : footerAreaView
    if (view === 'first') {
      return (kind === 'header' ? parsed.headerFirst : parsed.footerFirst)?.images
    }
    if (view === 'even') {
      return (kind === 'header' ? parsed.headerEven : parsed.footerEven)?.images
    }
    if (multiHf) {
      const rId = effRefsAll[hfSectionClamped]?.[kind]?.default
      const fromPart = rId ? parsed.hfParts?.[rId]?.images : undefined
      if (fromPart?.length) return fromPart
    }
    return (kind === 'header' ? parsed.headerImages : parsed.footerImages) ?? undefined
  }

  const commitHf = (kind: 'header' | 'footer', next: HeaderFooter) => {
    const view = kind === 'header' ? headerAreaView : footerAreaView
    // multi-section and not the final one: use the per-section edit channel (that section becomes its own part; earlier sections are unaffected)
    if (view === 'default' && multiHf && hfSectionClamped < sections.length - 1) {
      const sec = sections[hfSectionClamped]
      setSectionHfEdits((m) => ({ ...m, [`${sec.lastBlockIndex}:${kind}`]: next }))
      return
    }
    if (view === 'default') {
      if (kind === 'header') {
        setHeader(next)
        setHeaderDirty(true)
      } else {
        setFooter(next)
        setFooterDirty(true)
      }
      return
    }
    const key: HfVariantKey =
      view === 'first'
        ? kind === 'header'
          ? 'headerFirst'
          : 'footerFirst'
        : kind === 'header'
          ? 'headerEven'
          : 'footerEven'
    setHfVariants((v) => ({ ...v, [key]: next }))
    setHfVariantsDirty((d) => (d.includes(key) ? d : [...d, key]))
  }
  /** "Different first page" toggle, shared by the ribbon checkbox and the on-page chip */
  const toggleTitlePg = (on: boolean) => {
    setTitlePg(on)
    setTitlePgDirty(true)
    setHfView(on ? 'first' : 'default')
    setStatus(on ? t('appTitlePgOn') : t('appTitlePgOff'))
  }
  /** Unsaved section header/footer overrides for the pagination preview (default variant): local edits > document parts */
  const sectionHfOverride = useCallback(
    (si: number, kind: 'header' | 'footer'): HeaderFooter | null => {
      const sec = sections[si]
      if (!sec) return null
      if (si === sections.length - 1) {
        const dirty = kind === 'header' ? headerDirty : footerDirty
        return dirty ? (kind === 'header' ? header : footer) : null
      }
      return sectionHfEdits[`${sec.lastBlockIndex}:${kind}`] ?? null
    },
    [sections, sectionHfEdits, header, footer, headerDirty, footerDirty],
  )
  const [showMarks, setShowMarks] = useState(false)
  const [showRuler, setShowRuler] = useState(false)
  const [showNav, setShowNav] = useState(false)
  const [viewMode, setViewMode] = useState<ViewMode>('print')
  const [readMode, setReadMode] = useState(false)
  const [showGrid, setShowGrid] = useState(false)
  const [splitView, setSplitView] = useState(false)
  const [showPagePreview, setShowPagePreview] = useState(false)
  const [splitHtml, setSplitHtml] = useState('')
  const [showFind, setShowFind] = useState(false)
  const [showComments, setShowComments] = useState(false)
  /** Style definitions pending write-back (key = styleId), saved via SaveOptions.styleUpserts */
  const [styleUpserts, setStyleUpserts] = useState<Record<string, StyleUpsert>>({})
  const [comments, setComments] = useState<CommentInfo[]>([])
  const [commentsDirty, setCommentsDirty] = useState(false)
  const [watermark, setWatermark] = useState<string | null>(null)
  const [watermarkDirty, setWatermarkDirty] = useState(false)
  const [inkAnnotations, setInkAnnotations] = useState<InkAnnotation[]>([])
  const [inksDirty, setInksDirty] = useState(false)
  const [inkTool, setInkTool] = useState<InkTool>('select')
  const [inkPen, setInkPen] = useState<InkPenSettings>({ color: 'C00000', width: 2 })
  const [inkHighlighter, setInkHighlighter] = useState<InkPenSettings>({
    color: 'FFFF00',
    width: 10,
  })
  const [footnotes, setFootnotes] = useState<NoteInfo[]>([])
  /** Footnote ids already shown in canvas page gaps (the end-of-document list skips them to avoid duplication) */
  const [gapNoteIds, setGapNoteIds] = useState<Set<string>>(new Set())
  const [endnotes, setEndnotes] = useState<NoteInfo[]>([])
  const [notesDirty, setNotesDirty] = useState(false)
  const [sources, setSources] = useState<SourceInfo[]>([])
  const [sourcesDirty, setSourcesDirty] = useState(false)
  const [themeFonts, setThemeFonts] = useState<ThemeFonts | null>(null)
  const [themeFontsDirty, setThemeFontsDirty] = useState(false)
  const [themeColors, setThemeColors] = useState<ThemeColors | null>(null)
  const [themeColorsDirty, setThemeColorsDirty] = useState(false)
  const [commentComposing, setCommentComposing] = useState(false)
  const [trackChanges, setTrackChanges] = useState(false)
  const [revisionDisplay, setRevisionDisplay] = useState<RevisionDisplayMode>('all')
  // the original view restores old formatting via decorations: sync the mode and trigger one repaint
  useEffect(() => {
    revisionDisplayState.mode = revisionDisplay
    if (editor) editor.view.dispatch(editor.state.tr.setMeta('addToHistory', false))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [revisionDisplay])
  const [protection, setProtection] = useState<DocProtection | null>(null)
  const [protectionDirty, setProtectionDirty] = useState(false)
  const [compareResult, setCompareResult] = useState<{
    otherName: string
    entries: CompareEntry[]
  } | null>(null)
  const [autoSave, setAutoSave] = useState(() => localStorage.getItem('aidocs.autoSave') === '1')
  // tab closed but this renderer kept alive (shell freeze workaround): go inert
  const [tornDown, setTornDown] = useState(false)
  const [aiPreset, setAiPreset] = useState<{
    text: string
    nonce: number
    autoRun?: boolean
  } | null>(null)
  const [docCss, setDocCss] = useState('')
  // Live CJK-ness of the body while editing; overrides docCss's --doc-line-factor
  const [liveDocCjk, setLiveDocCjk] = useState<boolean | null>(null)
  const [stats, setStats] = useState<DocStats | null>(null)
  const [showLinkModal, setShowLinkModal] = useState(false)
  const [showEquationModal, setShowEquationModal] = useState(false)
  const [eqEditTarget, setEqEditTarget] = useState<{
    pos: number
    latex: string
    kind: 'inline' | 'block'
  } | null>(null)
  const [ctxMenu, setCtxMenu] = useState<ContextMenuState | null>(null)
  const [showFontDialog, setShowFontDialog] = useState(false)
  const [showParaDialog, setShowParaDialog] = useState(false)
  const [, forceRender] = useReducer((x: number) => x + 1, 0)
  const dirtyRef = useRef(false)
  // serializes save(): overlapping saves (Cmd+S vs autosave timer vs blur) would
  // otherwise race on the write + reparse + setContent sequence
  const saveInFlightRef = useRef(false)
  // set when a save wrote successfully but the editor changed while it was in
  // flight, so the file on disk is already one step behind. save() still counts
  // as succeeded; callers that must not lose data (the close guard) save again.
  const saveIncompleteRef = useRef(false)

  const editorRef = useRef<Editor | null>(null)
  const editor = useEditor({
    extensions: editorExtensions,
    content: { type: 'doc', content: [{ type: 'docParagraph' }] },
    editorProps: {
      // Word checks spelling as you type by default
      attributes: { class: 'doc-page', spellcheck: 'true' },
      // Word/web HTML cleanup: strip mso comments and <o:p>, unwrap <li><p>x</p></li>
      // (docListItem is an inline container; block-level p would shatter the list)
      transformPastedHTML: cleanPastedHtml,
      // clipboard images (screenshots/copied images): become inline images; mixed content
      // with HTML still uses default parsing (text in the HTML takes priority)
      handlePaste: (view, event) => {
        const data = event.clipboardData
        if (!data) return false
        const html = data.getData('text/html')
        // pasting block HTML onto an empty paragraph: replace wholesale (ProseMirror's default
        // first-block merge would demote headings to body text; Word keeps the source block format on empty paragraphs)
        const { $from, empty: selEmpty } = view.state.selection
        if (
          html &&
          selEmpty &&
          $from.parent.isTextblock &&
          $from.parent.content.size === 0 &&
          $from.depth === 1
        ) {
          try {
            const dom = new window.DOMParser().parseFromString(cleanPastedHtml(html), 'text/html')
            const parsed = PmDOMParser.fromSchema(view.state.schema).parse(dom.body)
            if (parsed.content.childCount > 0) {
              view.dispatch(
                view.state.tr.replaceWith($from.before(1), $from.after(1), parsed.content),
              )
              return true
            }
          } catch {
            /* fall back to default paste on parse failure */
          }
        }
        const imageFile = [...data.items]
          .find((item) => item.type.startsWith('image/'))
          ?.getAsFile()
        const text = data.getData('text/plain')
        if (imageFile && !text.trim()) {
          const reader = new FileReader()
          reader.onload = () => {
            const ed = editorRef.current
            if (ed && typeof reader.result === 'string') {
              void insertImageFromDataUrl(ed, reader.result, 'Image (pasted)')
            }
          }
          reader.readAsDataURL(imageFile)
          return true
        }
        // text/plain-only Markdown (code blocks, terminals, .md files, LLM
        // output): convert and insert as formatted content instead of literal
        // "## Heading" / "**bold**" characters
        if (!html && text) {
          const markdownHtml = markdownPasteHtml(text)
          if (markdownHtml !== null) {
            try {
              // cleanPastedHtml unwraps <li><p>…</p></li> from loose lists —
              // docListItem only allows inline content and would shatter them.
              const dom = new window.DOMParser().parseFromString(
                cleanPastedHtml(markdownHtml),
                'text/html',
              )
              const parser = PmDOMParser.fromSchema(view.state.schema)
              if (
                selEmpty &&
                $from.parent.isTextblock &&
                $from.parent.content.size === 0 &&
                $from.depth === 1
              ) {
                // same wholesale replace as block HTML onto an empty paragraph
                const parsed = parser.parse(dom.body)
                if (parsed.content.childCount > 0) {
                  view.dispatch(
                    view.state.tr.replaceWith($from.before(1), $from.after(1), parsed.content),
                  )
                  return true
                }
              } else {
                view.dispatch(
                  view.state.tr.replaceSelection(parser.parseSlice(dom.body)).scrollIntoView(),
                )
                return true
              }
            } catch {
              /* fall back to default literal paste on parse failure */
            }
          }
        }
        return false
      },
    },
    onSelectionUpdate: () => forceRender(),
    // typing in the main document takes ribbon routing back from any textbox
    onFocus: () => setActiveSubEditor(null),
    onUpdate: () => {
      dirtyRef.current = true
      forceRender()
    },
  })

  // textbox sub-editors: re-render the ribbon on focus/selection changes and
  // mark the document dirty when their content changes
  useEffect(
    () =>
      subscribeSubEditorState((docChanged) => {
        if (docChanged) dirtyRef.current = true
        forceRender()
      }),
    [],
  )

  useEffect(() => {
    void window.desktop.getRecentFiles().then(setRecent)
    void window.desktop.getAiSettings().then(setSettings)
  }, [])

  useEffect(() => {
    localStorage.setItem('aidocs.showAi', showAi ? '1' : '0')
  }, [showAi])

  useEffect(() => {
    localStorage.setItem('aidocs.autoSave', autoSave ? '1' : '0')
  }, [autoSave])

  // Pinch-to-zoom: Chromium delivers trackpad pinch as a wheel event
  // with ctrlKey set. Also support ⌘+scroll. Must be non-passive to preventDefault.
  useEffect(() => {
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return
      if (!(e.target as HTMLElement | null)?.closest?.('.editor-scroll')) return
      e.preventDefault()
      setZoom((z) => Math.min(200, Math.max(50, z - e.deltaY * 0.6)))
    }
    window.addEventListener('wheel', onWheel, { passive: false })
    return () => window.removeEventListener('wheel', onWheel)
  }, [])

  const isProtected = !!protection?.enforced && protection.edit === 'readOnly'

  // Read Mode / Restrict Editing: the document becomes read-only; Esc leaves Read Mode
  useEffect(() => {
    if (!editor) return
    editor.setEditable(!readMode && !isProtected)
    if (!readMode) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setReadMode(false)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [editor, readMode, isProtected])

  // Track Changes: the recorder plugin reads its toggle from extension storage
  useEffect(() => {
    if (!editor) return
    const storage = editor.storage.trackChanges as TrackChangesStorage
    storage.enabled = trackChanges
  }, [editor, trackChanges])

  // window title follows the document, so the OS window list and Switch Window show file names
  useEffect(() => {
    document.title = doc ? doc.fileName : 'GenOffice Docs'
  }, [doc])

  useEffect(() => window.desktop.onTeardown?.(() => setTornDown(true)), [])

  // keep the native View menu's checkmarks (AI Sidebar / Dark Mode) in sync
  useEffect(() => {
    window.desktop.reportViewMenuState?.({ aiSidebar: showAi, darkCanvas })
  }, [showAi, darkCanvas])

  // Crash-recovery copy: while the document is dirty, push a serialized
  // copy to the main process every 30s; a normal save (or discarding on close) removes
  // it, and reopening the file offers Restore/Discard when a newer copy exists.
  useEffect(() => {
    if (tornDown) return
    const timer = window.setInterval(() => {
      void writeRecoveryCopyImpl(fileCtxRef.current)
    }, 30_000)
    return () => window.clearInterval(timer)
  }, [tornDown])

  // Recompute the document-level line-height factor while editing:
  // docStyleCss decides it once at parse time, so typing CJK into a blank document
  // kept the Western factor until save/reopen (whole page ~8% shorter than the file).
  useEffect(() => {
    setLiveDocCjk(null)
  }, [doc])
  useEffect(() => {
    if (!editor) return
    const recompute = () => {
      let has = false
      editor.state.doc.descendants((node) => {
        if (has) return false
        if (node.isText && node.text && textHasCjk(node.text)) has = true
        return !has
      })
      setLiveDocCjk(has)
    }
    let timer = 0
    const onUpdate = () => {
      window.clearTimeout(timer)
      timer = window.setTimeout(recompute, 300)
    }
    editor.on('update', onUpdate)
    return () => {
      window.clearTimeout(timer)
      editor.off('update', onUpdate)
    }
  }, [editor])

  // Split pane: keep the read-only bottom copy in sync with the editor (debounced)
  useEffect(() => {
    if (!splitView || !editor) return
    const sync = () => setSplitHtml(editor.getHTML())
    sync()
    let timer = 0
    const onUpdate = () => {
      window.clearTimeout(timer)
      timer = window.setTimeout(sync, 300)
    }
    editor.on('update', onUpdate)
    return () => {
      window.clearTimeout(timer)
      editor.off('update', onUpdate)
    }
  }, [splitView, editor])

  // Mixed-paper menu export: after the preview mounts and renders, the merge export resumes automatically
  const pendingMixedExportRef = useRef<boolean | string>(false)

  /** App state bundle for the extracted file actions (file-actions.ts); refreshed every render. */
  const fileCtxRef = useRef<FileActionContext>(null as unknown as FileActionContext)
  fileCtxRef.current = {
    editor,
    doc,
    dirtyRef,
    saveInFlightRef,
    saveIncompleteRef,
    pendingMixedExportRef,
    setStatus,
    setRecent,
    setShowAi,
    setDoc,
    setAiPanelKey,
    setDocCss,
    setShowPagePreview,
    section,
    sectionDirty,
    sections,
    sectionsDirty,
    trailingStartType,
    setSection,
    setSections,
    setSectionDirty,
    setSectionsDirty,
    setTrailingStartType,
    pageColor,
    pageColorDirty,
    setPageColor,
    setPageColorDirty,
    header,
    headerDirty,
    footer,
    footerDirty,
    setHeader,
    setHeaderDirty,
    setFooter,
    setFooterDirty,
    hfVariants,
    hfVariantsDirty,
    setHfVariants,
    setHfVariantsDirty,
    sectionHfEdits,
    setSectionHfEdits,
    titlePg,
    titlePgDirty,
    evenOddHf,
    evenOddHfDirty,
    setTitlePg,
    setTitlePgDirty,
    setEvenOddHf,
    setEvenOddHfDirty,
    // opening a document resets to the resting per-page variant selection
    setHfView: (v: HfView) => {
      setHfViewState(v)
      setHfViewTouched(false)
    },
    pgNumEdit,
    pgNumDirtySections,
    setPgNumEdit,
    setPgNumDirtySections,
    pendingNumbering,
    numberingDirty,
    setPendingNumbering,
    styleUpserts,
    setStyleUpserts,
    comments,
    commentsDirty,
    setComments,
    setCommentsDirty,
    setShowComments,
    setCommentComposing,
    watermark,
    watermarkDirty,
    setWatermark,
    setWatermarkDirty,
    inkAnnotations,
    inksDirty,
    setInkAnnotations,
    setInksDirty,
    setInkTool,
    footnotes,
    endnotes,
    notesDirty,
    setFootnotes,
    setEndnotes,
    setNotesDirty,
    sources,
    sourcesDirty,
    setSources,
    setSourcesDirty,
    themeFonts,
    themeFontsDirty,
    themeColors,
    themeColorsDirty,
    setThemeFonts,
    setThemeFontsDirty,
    setThemeColors,
    setThemeColorsDirty,
    setTrackChanges,
    protection,
    protectionDirty,
    setProtection,
    setProtectionDirty,
    setCompareResult,
  }

  const loadFile = useCallback(
    (result: OpenFileResult | null) => loadFileImpl(fileCtxRef.current, result),
    [],
  )

  // file renamed externally (renamed in the shell Home list) → sync the save path and title-bar file name (content unchanged)
  useEffect(
    () =>
      window.desktop.onRenamedDocx(({ oldPath, newPath }) => {
        setDoc((prev) =>
          prev && prev.filePath === oldPath
            ? {
                ...prev,
                filePath: newPath,
                fileName: newPath.split(/[\\/]/).pop() ?? prev.fileName,
              }
            : prev,
        )
      }),
    [],
  )

  useEffect(() => {
    if (!editor) return
    const unsubscribe = window.desktop.onOpenDocx((result) => {
      void loadFile(result)
    })
    // With no pending file the window lands directly in the editor on a blank document
    // (the AI panel carries the generate-from-prompt flow). StrictMode runs the mount
    // effect twice but the pending queues can only be consumed once, so the consume
    // Promise lives in a ref and its result is processed only once.
    bootPendingRef.current ??= Promise.all([
      window.desktop.consumePendingOpenDocx(),
      // Still consume the one-shot new-blank flag so it doesn't leak into the next open
      window.desktop.consumeNewBlankDoc(),
    ])
    void bootPendingRef.current
      .then(async ([pending]) => {
        if (bootHandledRef.current) return
        bootHandledRef.current = true
        if (pending) await loadFile(pending)
        else await newFile()
      })
      // Open failures also land on a blank document, or the tab stays at "Opening…" forever
      .catch(() => {
        if (bootHandledRef.current) return
        bootHandledRef.current = true
        void newFile().catch(() => {})
      })
    return unsubscribe
    // newFile depends on editor (already in deps); we capture it by closure
    // rather than listing it to avoid a forward-reference TypeScript error
    // (newFile is declared after this effect in source order).
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editor, loadFile])

  const openFile = useCallback(async () => {
    await loadFile(await window.desktop.openDocx())
  }, [loadFile])

  /** new document from the built-in blank template (AI can then generate into it) */
  const newFile = useCallback(() => newFileImpl(fileCtxRef.current), [])

  const openRecent = useCallback(
    async (path: string) => {
      await loadFile(await window.desktop.openDocxPath(path))
    },
    [loadFile],
  )

  const save = useCallback(
    (saveAs: boolean, auto = false) => saveImpl(fileCtxRef.current, saveAs, auto),
    [],
  )

  // inserting a section break needs one save for the new section to take effect; the
  // flag is consumed in the render after state commit, guaranteeing the save closure
  // sees the latest sectionsDirty/trailingStartType
  const pendingSectionSaveRef = useRef(false)
  useEffect(() => {
    if (pendingSectionSaveRef.current && doc?.filePath) {
      pendingSectionSaveRef.current = false
      void save(false, true)
    }
  })

  /**
   * Insert a section break: the new break paragraph takes a copy of the current
   * section's sectPr (content before the break keeps the original
   * section's settings); "continuous" is written to the following section's (the
   * original current section sectPr's) w:type.
   */
  const insertSectionBreak = useCallback(
    (type: SectionInfo['startType']) => {
      if (!editor || !doc) return
      const cur = sections[activeSection]
      const sectPrCopy =
        cur?.sectPrXml ||
        applySectionSettings(
          '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>',
          section ?? sections[0]?.settings ?? DEFAULT_SECTION,
        )
      const genXml = `<w:p><w:pPr>${sectPrCopy}</w:pPr></w:p>`
      const { $head } = editor.state.selection
      const pos = $head.depth > 0 ? $head.after(1) : $head.pos
      editor
        .chain()
        .focus()
        .insertContentAt(pos, {
          type: 'docProtected',
          attrs: {
            blockType: 'passthrough',
            label: 'Section break paragraph',
            previewText: '',
            genXml,
          },
        })
        .run()
      // the section after the break is terminated by the "original current section sectPr",
      // whose w:type decides how the new section starts: always write the chosen type back
      // (including a nextPage reset, so a leftover continuous from the original section doesn't linger)
      if (sections.length === 0 || activeSection === sections.length - 1) {
        setTrailingStartType(type)
      } else {
        setSections((prev) =>
          prev.map((s, i) => (i === activeSection ? { ...s, startType: type } : s)),
        )
        setSectionsDirty((d) => (d.includes(activeSection) ? d : [...d, activeSection]))
      }
      const labels: Record<SectionInfo['startType'], string> = {
        nextPage: t('appBreakNextPage'),
        continuous: t('appBreakContinuous'),
        evenPage: t('appBreakEvenPage'),
        oddPage: t('appBreakOddPage'),
      }
      if (doc.filePath) {
        pendingSectionSaveRef.current = true
        setStatus(t('appSectionBreakInserted', { type: labels[type] }))
      } else {
        setStatus(t('appSectionBreakPending'))
      }
    },
    [editor, doc, sections, activeSection, section],
  )

  // ---- List numbering: restart numbering / new lists reuse document definitions (numbering.xml write-back) ----

  /** Floor of numIds allocated within the same render cycle (prevents double allocation before setState commits) */
  const numIdFloorRef = useRef(0)

  /** Allocate an unused numId (above parsed definitions + pending appends + this cycle's allocations) */
  /** App state bundle for the extracted numbering actions (numbering-actions.ts); refreshed every render. */
  const numberingCtxRef = useRef<NumberingContext>(null as unknown as NumberingContext)
  numberingCtxRef.current = {
    editor,
    doc,
    pendingNumbering,
    setPendingNumbering,
    numIdFloorRef,
    setStatus,
  }

  const createCustomListDef = useCallback(
    (levels: CustomNumberingLevel[]) => createCustomListDefImpl(numberingCtxRef.current, levels),
    [],
  )
  const allocateListNumId = useCallback(
    (kind: 'bullet' | 'ordered') => allocateListNumIdImpl(numberingCtxRef.current, kind),
    [],
  )
  const restartNumbering = useCallback(() => restartNumberingImpl(numberingCtxRef.current), [])
  const continueNumbering = useCallback(() => continueNumberingImpl(numberingCtxRef.current), [])

  // ---- Inline fields: insertion and F9 update (cached results recomputed locally) ----

  const fieldValue = useCallback(
    (instr: string): string => {
      const kw = instr.trim().split(/\s+/)[0]?.toUpperCase()
      const now = new Date()
      switch (kw) {
        case 'DATE':
        case 'CREATEDATE':
        case 'SAVEDATE':
          return now.toLocaleDateString('zh-CN')
        case 'TIME':
          return now.toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' })
        case 'NUMPAGES':
          return String(pageInfo.total)
        case 'PAGE':
          return String(pageInfo.current)
        case 'FILENAME':
          return doc?.fileName ?? ''
        default:
          return ''
      }
    },
    [pageInfo, doc],
  )

  const insertField = useCallback(
    (instr: string) => {
      if (!editor) return
      const value = fieldValue(instr) || ' '
      editor
        .chain()
        .focus()
        .insertContent({
          type: 'text',
          text: value,
          marks: [{ type: 'instrField', attrs: { instr } }],
        })
        .unsetMark('instrField')
        .run()
      setStatus(t('appFieldInserted', { instr }))
    },
    [editor, fieldValue],
  )

  /** F9: recompute all inline field caches (PAGE/NUMPAGES/date-time/file name); REF/TOC are recomputed when Word opens the file */
  const updateFields = useCallback(() => {
    if (!editor) return
    const { state, view } = editor
    const jobs: Array<{ from: number; to: number; text: string; marks: readonly PmMark[] }> = []
    state.doc.descendants((node, pos) => {
      if (!node.isText) return
      const mark = node.marks.find((m) => m.type.name === 'instrField')
      if (!mark) return
      const next = fieldValue(String(mark.attrs.instr))
      if (next && next !== node.text) {
        jobs.push({ from: pos, to: pos + node.nodeSize, text: next, marks: node.marks })
      }
    })
    if (jobs.length === 0) {
      setStatus(t('appNoFieldsToUpdate'))
      return
    }
    let tr = state.tr
    for (const j of jobs.sort((a, b) => b.from - a.from)) {
      tr = tr.replaceWith(j.from, j.to, state.schema.text(j.text, [...j.marks]))
    }
    view.dispatch(tr)
    setStatus(t('appFieldsUpdated', { n: jobs.length }))
  }, [editor, fieldValue])

  // editorRef: lets the handlePaste closure (the useEditor config exists before the instance) reach the instance
  useEffect(() => {
    editorRef.current = editor
  }, [editor])

  /** Pasted list items lacking a numId (schema default null; saving would lose list semantics):
   *  reuse the numId of an existing same-kind instance, otherwise fall back to creating a definition */
  useEffect(() => {
    if (!editor) return
    const fill = () => {
      const missing: Array<{ pos: number; attrs: Record<string, unknown> }> = []
      const reuse: Partial<Record<'bullet' | 'ordered', string>> = {}
      editor.state.doc.descendants((node, pos) => {
        if (node.type.name !== 'docListItem') return
        const kind = (node.attrs.kind as 'bullet' | 'ordered') ?? 'bullet'
        if (node.attrs.numId !== null) {
          if (!reuse[kind]) reuse[kind] = String(node.attrs.numId)
          return
        }
        missing.push({ pos, attrs: node.attrs })
      })
      if (missing.length === 0) return
      const idOf: Partial<Record<'bullet' | 'ordered', string | null>> = {}
      let tr = editor.state.tr
      for (const item of missing) {
        const kind = (item.attrs.kind as 'bullet' | 'ordered') ?? 'bullet'
        if (!(kind in idOf)) idOf[kind] = reuse[kind] ?? allocateListNumId(kind)
        const numId = idOf[kind]
        if (numId) tr = tr.setNodeMarkup(item.pos, undefined, { ...item.attrs, numId })
      }
      if (tr.docChanged) editor.view.dispatch(tr)
    }
    editor.on('update', fill)
    return () => {
      editor.off('update', fill)
    }
  }, [editor, allocateListNumId])

  /** Open the page-number-format dialog (initial value = the cursor section's pgNumType) */
  const openPgNumModal = useCallback(() => {
    const sec = sections[Math.min(activeSection, sections.length - 1)]
    setPgNumModal({
      fmt: sec?.pageNumberFmt ?? 'decimal',
      start: sec?.pageNumberStart !== undefined ? String(sec.pageNumberStart) : '',
    })
  }, [sections, activeSection])

  /** Apply the page-number format to the cursor's section: final section via SaveOptions.pgNumType, others rewrite their sectPr */
  const applyPgNumFormat = useCallback(() => {
    if (!pgNumModal) return
    const fmt = pgNumModal.fmt === 'decimal' ? undefined : pgNumModal.fmt
    const start =
      pgNumModal.start.trim() === '' ? undefined : Math.max(0, parseInt(pgNumModal.start, 10) || 0)
    const idx = sections.length > 0 ? Math.min(activeSection, sections.length - 1) : -1
    if (idx >= 0) {
      setSections((prev) =>
        prev.map((s, i) => (i === idx ? { ...s, pageNumberFmt: fmt, pageNumberStart: start } : s)),
      )
    }
    if (idx < 0 || idx === sections.length - 1) {
      setPgNumEdit({
        ...(fmt !== undefined ? { fmt } : {}),
        ...(start !== undefined ? { start } : {}),
      })
    } else {
      setPgNumDirtySections((d) => (d.includes(idx) ? d : [...d, idx]))
    }
    setPgNumModal(null)
    setStatus(t('appPgNumFormatSet'))
  }, [pgNumModal, sections, activeSection])

  const exportPdf = useCallback(
    (outPath?: string) => exportPdfImpl(fileCtxRef.current, outPath),
    [],
  )

  // for real-device verification: trigger export directly via CDP (same as __pageDebug)
  useEffect(() => {
    ;(window as unknown as Record<string, unknown>).__exportPdf = exportPdf
  }, [exportPdf])

  useEffect(() => {
    if (!showPagePreview || pendingMixedExportRef.current === false) return
    const pending = pendingMixedExportRef.current
    pendingMixedExportRef.current = false
    const timer = window.setTimeout(() => {
      void exportPdf(typeof pending === 'string' ? pending : undefined)
    }, 1200)
    return () => window.clearTimeout(timer)
  }, [showPagePreview, exportPdf])

  // ---- References: footnotes / endnotes ----

  const [notePrompt, setNotePrompt] = useState<NotePrompt | null>(null)
  const [protectModal, setProtectModal] = useState<ProtectModalState | null>(null)

  /** App state bundle for the extracted review actions (review-actions.ts); refreshed every render. */
  const reviewCtxRef = useRef<ReviewContext>(null as unknown as ReviewContext)
  reviewCtxRef.current = {
    editor,
    doc,
    dirtyRef,
    setStatus,
    notePrompt,
    setNotePrompt,
    footnotes,
    endnotes,
    setFootnotes,
    setEndnotes,
    setNotesDirty,
    comments,
    setComments,
    setCommentsDirty,
    setCommentComposing,
    setShowComments,
    setInkAnnotations,
    setInksDirty,
    protection,
    setProtection,
    setProtectionDirty,
    protectModal,
    setProtectModal,
    setCompareResult,
  }

  const insertNote = useCallback((kind: 'footnote' | 'endnote') => {
    setNotePrompt({ kind })
  }, [])

  const editNote = useCallback((kind: 'footnote' | 'endnote', id: string) => {
    setNotePrompt({ kind, id })
  }, [])

  const submitNote = useCallback((text: string) => submitNoteImpl(reviewCtxRef.current, text), [])

  const deleteNote = useCallback(
    (kind: 'footnote' | 'endnote', id: string) => deleteNoteImpl(reviewCtxRef.current, kind, id),
    [],
  )

  // Word: body shading of resolved threads is hidden
  useEffect(() => {
    if (!editor) return
    const done = new Set(comments.filter((c) => c.done).map((c) => c.id))
    editor.view.dispatch(editor.state.tr.setMeta(resolvedCommentsPluginKey, done))
  }, [editor, comments])

  const cancelNewComment = useCallback(() => cancelNewCommentImpl(reviewCtxRef.current), [])
  const startNewComment = useCallback(() => startNewCommentImpl(reviewCtxRef.current), [])
  const submitNewComment = useCallback(
    (text: string) => submitNewCommentImpl(reviewCtxRef.current, text),
    [],
  )
  const replyToComment = useCallback(
    (parentId: string, text: string) => replyToCommentImpl(reviewCtxRef.current, parentId, text),
    [],
  )
  const resolveComment = useCallback(
    (id: string, done: boolean) => resolveCommentImpl(reviewCtxRef.current, id, done),
    [],
  )
  const deleteComment = useCallback((id: string) => deleteCommentImpl(reviewCtxRef.current, id), [])
  const handleRevision = useCallback(
    (action: 'accept' | 'reject', all: boolean) =>
      handleRevisionImpl(reviewCtxRef.current, action, all),
    [],
  )
  const addInk = useCallback(
    (annotation: InkAnnotation) => addInkImpl(reviewCtxRef.current, annotation),
    [],
  )
  const removeInks = useCallback((ids: string[]) => removeInksImpl(reviewCtxRef.current, ids), [])
  const clearInks = useCallback(() => clearInksImpl(reviewCtxRef.current), [])
  const toggleProtection = useCallback(() => toggleProtectionImpl(reviewCtxRef.current), [])
  const submitProtectModal = useCallback(() => submitProtectModalImpl(reviewCtxRef.current), [])
  const compareWithFile = useCallback(() => compareWithFileImpl(reviewCtxRef.current), [])

  const revisionCount = editor && doc ? revisionCountOfDoc(editor.state.doc) : 0

  /** Uncapped width-fit ratio (%) from the scroller's measured size */
  const rawFitWidth = useCallback(() => {
    if (!section) return null
    const scroller = document.querySelector('.editor-scroll')
    if (!scroller) return null
    return ((scroller.clientWidth - 48) / twipsToPx(section.pageWidth)) * 100
  }, [section])

  /** Last auto-fit: if current zoom still equals its value → "fit mode", re-fit on size changes */
  const lastFitRef = useRef<{ mode: 'width' | 'page'; value: number } | null>(null)
  const zoomLiveRef = useRef(zoom)
  useEffect(() => {
    zoomLiveRef.current = zoom
  }, [zoom])

  /** Word's page-width / whole-page zoom: compute the actual zoom ratio from the current window size */
  const zoomFit = useCallback(
    (mode: 'width' | 'page') => {
      if (!section) return
      const scroller = document.querySelector('.editor-scroll')
      if (!scroller) return
      const pad = 48
      const wFit = ((scroller.clientWidth - pad) / twipsToPx(section.pageWidth)) * 100
      const hFit = ((scroller.clientHeight - pad) / twipsToPx(section.pageHeight)) * 100
      // whole page = the entire page visible, so it must fit both dimensions;
      // floor, not round: rounding up would push the page past the pane edge
      const next = mode === 'width' ? wFit : Math.min(wFit, hFit)
      const applied = Math.min(200, Math.max(50, Math.floor(next)))
      lastFitRef.current = { mode, value: applied }
      setZoom(applied)
    },
    [section],
  )

  // On scroller size changes (window/AI-dock/nav-pane toggles): follow with a
  // re-fit while in fit mode, and clamp any manual zoom back down to width-fit
  // whenever the page no longer fits horizontally — the canvas must never
  // overflow the pane on a resize (same contract as the slides stage). A manual
  // zoom smaller than fit is left alone.
  useEffect(() => {
    const el = document.querySelector('.editor-scroll')
    if (!doc || !el) return
    const ro = new ResizeObserver(() => {
      const raw = rawFitWidth()
      if (raw == null) return
      const lf = lastFitRef.current
      if (lf != null && Math.abs(zoomLiveRef.current - lf.value) <= 0.5) {
        zoomFit(lf.mode) // fit mode: follow the container in the user's chosen fit
        return
      }
      // The overflow test uses the uncapped ratio: a manual zoom that still
      // fits (or is smaller than fit) is the user's choice and must be kept
      if (zoomLiveRef.current <= raw + 0.5) return
      zoomFit('width')
    })
    ro.observe(el)
    return () => ro.disconnect()
  }, [doc, zoomFit, rawFitWidth])

  // page-bottom height (px) reserved for footnote references inside a block: same estimation model as the parity runner
  const footnoteExtraOf = useCallback(
    (b: Block): number => {
      const runs = blockNoteScanRuns(b)
      if (runs.length === 0 || footnotes.length === 0) return 0
      let extra = 0
      for (const run of runs) {
        if (run.noteRef?.kind !== 'footnote') continue
        const sec =
          sections.find((s) => (b.docxIndex ?? 0) <= s.lastBlockIndex)?.settings ?? section
        if (!sec) continue
        const contentW = twipsToPx(sec.pageWidth - sec.marginLeft - sec.marginRight)
        const fnText = footnotes.find((f) => f.id === run.noteRef!.id)?.text ?? ''
        extra += estimateFootnoteHeight(fnText, contentW, sec.docGrid)
      }
      return extra > 0 ? extra + FOOTNOTE_SEPARATOR_H : 0
    },
    [footnotes, sections, section],
  )

  // typed w:docGrid line pitch (pt) when every section shares one; null = no snapping
  const gridPitchPt = useMemo(() => docGridPitchPt(sections), [sections])

  // single-section header/footer push-down: body top = max(marginTop, headerDist + header height)
  const singleHfPx = useMemo(() => {
    if (!section) return { headerPx: 0, footerPx: 0 }
    const contentW = twipsToPx(section.pageWidth - section.marginLeft - section.marginRight)
    return {
      headerPx: estimateHfHeight(header, contentW, doc?.parsed.headerImages),
      footerPx: estimateHfHeight(footer, contentW, doc?.parsed.footerImages),
    }
  }, [section, header, footer, doc])
  const effTopSingle = section ? effectiveTopPx(section, singleHfPx.headerPx) : 0
  const effBottomSingle = section ? effectiveBottomPx(section, singleHfPx.footerPx) : 0

  // multi-section: estimated heights of each section's default-variant header/footer (capacity per section, variant differences ignored)
  const hfHeightsOf = useCallback(
    (secs: SectionInfo[]): SectionHfHeights[] => {
      const refs = effectiveHfRefs(secs)
      return secs.map((s, i) => {
        const set = s.settings
        const contentW = twipsToPx(set.pageWidth - set.marginLeft - set.marginRight)
        const pick = (kind: 'header' | 'footer'): HeaderFooter | null => {
          if (i === secs.length - 1) return kind === 'header' ? header : footer
          const edited = sectionHfEdits[`${s.lastBlockIndex}:${kind}`]
          if (edited) return edited
          const rId = refs[i]?.[kind]?.default
          return rId ? hfFromPart(doc?.parsed.hfParts?.[rId]) : null
        }
        const imagesOf = (kind: 'header' | 'footer') => {
          const rId = refs[i]?.[kind]?.default
          const fromPart = rId ? doc?.parsed.hfParts?.[rId]?.images : undefined
          if (fromPart?.length) return fromPart
          if (i === secs.length - 1) {
            return (
              (kind === 'header' ? doc?.parsed.headerImages : doc?.parsed.footerImages) ?? undefined
            )
          }
          return undefined
        }
        return {
          headerPx: estimateHfHeight(pick('header'), contentW, imagesOf('header')),
          footerPx: estimateHfHeight(pick('footer'), contentW, imagesOf('footer')),
        }
      })
    },
    [doc, header, footer, sectionHfEdits],
  )

  // pagination-constraint injection: docxIndex → parse-layer semantics (keepNext/keepLines/widow/table-row flags).
  // DOM measurement only has geometry; these constraints decide page cut points (no orphan headings / unbreakable lines / repeated table headers).
  const blockMetaOf = useCallback(
    (docxIndex: number): BlockMeta | undefined => {
      const b = doc?.parsed.blocks.find((bl) => bl.docxIndex === docxIndex)
      if (!b) return undefined
      if (b.type === 'table') {
        if (!b.originalXml || !/tblHeader|cantSplit/.test(b.originalXml)) return undefined
        return { tableRowFlags: tableRowFlags(b.originalXml) }
      }
      const styleDisplay = b.styleId ? doc?.parsed.styles.get(b.styleId)?.display : undefined
      const keepNext = b.format?.keepNext ?? styleDisplay?.keepNext
      const keepLines = b.format?.keepLines ?? styleDisplay?.keepLines
      const widowOff = b.format?.widowControl === false
      const fnExtra = footnoteExtraOf(b)
      if (!keepNext && !keepLines && !widowOff && fnExtra === 0) return undefined
      return {
        ...(keepNext ? { keepNext: true } : {}),
        ...(keepLines ? { keepLines: true } : {}),
        ...(widowOff ? { widowControl: false as const } : {}),
        ...(fnExtra > 0 ? { footnoteExtraPx: fnExtra } : {}),
      }
    },
    [doc, footnoteExtraOf],
  )

  // per-page footnote collection: the page of the referencing block → that page's footnote entries (number/text/estimated height)
  const pageFootnotesOf = useCallback(
    (blocks: BlockBox[], slices: PageSlice[]): PageNoteItem[][] => {
      const out: PageNoteItem[][] = slices.map(() => [])
      if (!doc || footnotes.length === 0 || !section) return out
      const noOf = new Map(footnotes.map((f, i) => [f.id, i + 1]))
      for (const b of blocks) {
        if (b.docxIndex === undefined) continue
        const pb = doc.parsed.blocks.find((bl) => bl.docxIndex === b.docxIndex)
        if (!pb) continue
        const ids = blockNoteScanRuns(pb)
          .filter((r) => r.noteRef?.kind === 'footnote')
          .map((r) => r.noteRef!.id)
        if (ids.length === 0) continue
        const page = pageAt(slices, b.top + 0.5) - 1
        const sec = sections.find((s) => b.docxIndex! <= s.lastBlockIndex)?.settings ?? section
        const contentW = twipsToPx(sec.pageWidth - sec.marginLeft - sec.marginRight)
        for (const id of ids) {
          const fn = footnotes.find((f) => f.id === id)
          if (!fn) continue
          out[page]?.push({
            no: noOf.get(id) ?? 0,
            id,
            text: fn.text,
            ...(fn.richParas ? { richParas: fn.richParas } : {}),
            height: estimateFootnoteHeight(fn.text, contentW, sec.docGrid),
          })
        }
      }
      return out
    },
    [doc, footnotes, sections, section],
  )

  // endnote-area entries (placed together at the document end, shared by pagination preview and page slicing): height estimated with the final section's content width
  const endnoteItems = useMemo<PageNoteItem[]>(() => {
    if (!section || endnotes.length === 0) return []
    const sec = sections[sections.length - 1]?.settings ?? section
    const contentW = twipsToPx(sec.pageWidth - sec.marginLeft - sec.marginRight)
    return endnotes.map((n, i) => ({
      no: i + 1,
      id: n.id,
      text: n.text,
      ...(n.richParas ? { richParas: n.richParas } : {}),
      height: estimateFootnoteHeight(n.text, contentW, sec.docGrid),
    }))
  }, [endnotes, sections, section])

  // canvas column-flow geometry (non-null when the cursor's section has equal-width columns —
  // same follow-the-cursor rule as the canvas page box): shared by canvas CSS / measuring state / preview.
  // equalWidth="0" (unequal local layout columns) is not modeled, matching the engine's scope
  const colFlow = useMemo(() => {
    const cur = sections[Math.min(activeSection, sections.length - 1)]
    if (!cur || sectionColumns(cur) <= 1) return null
    return sectionColGeom(cur)
  }, [sections, activeSection])

  // single-flow measuring state for the columned canvas: temporarily drop CSS columns and
  // set the width to the column width so DOM measurement yields 1-D coordinates matching
  // the engine's column flow (synchronous layout round-trip, no visible flicker)
  const measureSingleFlow = useCallback(
    function run<T>(pm: HTMLElement, fn: () => T): T {
      if (!colFlow || viewMode !== 'print') return fn()
      pm.classList.add('measuring-columns')
      try {
        return fn()
      } finally {
        pm.classList.remove('measuring-columns')
      }
    },
    [colFlow, viewMode],
  )

  // column-flow geometry gate: when the canvas column CSS is inactive, measure as full-width single flow; the geometry must drop cols to match
  const colGeomsFor = useCallback(
    (geoms: SectionGeom[]): SectionGeom[] => {
      if (colFlow && viewMode === 'print') return geoms
      for (const g of geoms) if (g.cols) g.cols = undefined
      return geoms
    },
    [colFlow, viewMode],
  )

  // real TOC page-number backfill: compute each heading's (docHeading) page from the current real page slicing.
  // Returns page numbers matching docHeadings in document order 1:1; returns null when not computable (page numbers left blank).
  const headingPages = useCallback((): number[] | null => {
    if (!editor || !section) return null
    const pm = document.querySelector('.editor-scroll .ProseMirror') as HTMLElement | null
    if (!pm) return null
    const factor = zoom / 100
    const { mBlocks, slices, secs } = measureSingleFlow(pm, () => {
      const origin = pm.getBoundingClientRect().top + effTopSingle * factor
      const { blocks, totalHeight } = measureBlocks(pm, origin, factor)
      const live = liveSections(sections, blocks)
      let s: PageSlice[]
      if (live.length > 0) {
        assignSections(blocks, live)
        s = sliceWithLineSplit(
          blocks,
          colGeomsFor(sectionGeoms(live, hfHeightsOf(live))),
          totalHeight,
          factor,
          blockMetaOf,
        )
      } else {
        const contentH = twipsToPx(section.pageHeight) - effTopSingle - effBottomSingle
        s = sliceWithLineSplit(
          blocks,
          [{ contentHeight: contentH, forceBreak: false }],
          totalHeight,
          factor,
          blockMetaOf,
        )
      }
      return { mBlocks: blocks, slices: s, secs: live }
    })
    // same page-number algorithm as the footer path (w:pgNumType start offsets apply to single-section docs too)
    const nums = secs.length > 0 ? pageNumbers(slices, secs) : slices.map((_, i) => i + 1)
    const byEl = new Map(mBlocks.filter((b) => b.el).map((b) => [b.el as HTMLElement, b.top]))
    const pages: number[] = []
    for (const h of collectHeadings(editor.state.doc)) {
      const dom = editor.view.nodeDOM(h.pos) as HTMLElement | null
      const top = dom ? byEl.get(dom) : undefined
      const idx = top === undefined ? 1 : pageAt(slices, top + 1)
      pages.push(nums[Math.min(Math.max(idx, 1), nums.length) - 1] ?? idx)
    }
    return pages
  }, [
    editor,
    section,
    sections,
    zoom,
    blockMetaOf,
    effTopSingle,
    effBottomSingle,
    hfHeightsOf,
    measureSingleFlow,
    colGeomsFor,
  ])

  // status-bar page number: real page slicing (same algorithm as the pagination preview). Edits remeasure with debounce; scrolling only relocates
  useEffect(() => {
    if (!doc || !section) {
      setPageInfo({ current: 1, total: 1 })
      setLastPageNo(null)
      return
    }
    const scroller = document.querySelector('.editor-scroll')
    if (!scroller) return
    const factor = zoom / 100
    const mTopPx = effTopSingle
    const contentH = twipsToPx(section.pageHeight) - effTopSingle - effBottomSingle
    let slices: PageSlice[] = []
    let timer: number | null = null
    const pmEl = () => document.querySelector('.editor-scroll .ProseMirror') as HTMLElement | null
    const locate = () => {
      const pm = pmEl()
      if (!pm || slices.length === 0) return
      const pmRect = pm.getBoundingClientRect()
      const scrollRect = scroller.getBoundingClientRect()
      const origin = pmRect.top + mTopPx * factor
      const midScreen = Math.min(scrollRect.top + scrollRect.height / 2, pmRect.bottom)
      // slices use gapless virtual coordinates; subtract the page-gap height above the midpoint
      // (including mid-paragraph inline gaps and repeated-header clone rows, which carry
      // page-repeat-header but not page-gap)
      let gapAbove = 0
      for (const gap of pm.querySelectorAll('.page-gap, .page-repeat-header')) {
        const r = gap.getBoundingClientRect()
        if (r.top < midScreen) gapAbove += Math.min(r.height, midScreen - r.top)
      }
      const midY = (midScreen - origin - gapAbove) / factor
      // visible-page numbering so the status bar and F9 NUMPAGES agree with the gap widgets
      const current = visiblePageCount(slices, pageAt(slices, midY))
      const total = visiblePageCount(slices)
      setPageInfo((prev) =>
        prev.current === current && prev.total === total ? prev : { current, total },
      )
    }
    const remeasure = () => {
      const pm = pmEl()
      if (!pm) return
      const tStart = performance.now()
      let tMeasure = 0
      let tSlice = 0
      // a columned canvas measures + slices in the single-flow measuring state (fillLineBoxes
      // also reads the DOM for line sampling, so it must share the state); display-state DOM
      // reads like gap positioning happen outside the measuring state
      const measured = measureSingleFlow(pm, () => {
        const t0 = performance.now()
        const origin = pm.getBoundingClientRect().top + mTopPx * factor
        const { blocks, totalHeight } = measureBlocks(pm, origin, factor)
        tMeasure = performance.now() - t0
        // multi-section: assign blocks to sections by docxIndex; each section has its own content height / forced breaks.
        // liveSections: when a section-break block is deleted, that section merges into the next in real time (effective before saving)
        const secList = sections.length > 0 ? liveSections(sections, blocks) : null
        if (secList) assignSections(blocks, secList)
        // the endnote area takes part in page slicing (placed together at the document end; overflows continue on later pages)
        const withEndnotes = appendEndnotesBlock(
          blocks,
          totalHeight,
          endnoteItems,
          FOOTNOTE_SEPARATOR_H,
        )
        const flowH = withEndnotes?.totalHeight ?? totalHeight
        const hfHs = secList ? hfHeightsOf(secList) : null
        const t1 = performance.now()
        const s = secList
          ? sliceWithLineSplit(
              blocks,
              colGeomsFor(sectionGeoms(secList, hfHs!)),
              flowH,
              factor,
              blockMetaOf,
            )
          : sliceWithLineSplit(
              blocks,
              [{ contentHeight: contentH, forceBreak: false }],
              flowH,
              factor,
              blockMetaOf,
            )
        tSlice = performance.now() - t1
        return { blocks, secList, hfHs, s }
      })
      const { blocks, secList, hfHs } = measured
      slices = measured.s
      // document-end footer shows the last page's displayed number, not the physical count
      if (slices.length > 0) {
        const lastIdx = slices.length - 1
        const num = secList ? pageNumbers(slices, secList)[lastIdx] : slices.length
        setLastPageNo({
          num,
          text: secList
            ? formatPageNumber(
                num,
                secList[Math.min(slices[lastIdx].section, secList.length - 1)]?.pageNumberFmt,
              )
            : String(num),
        })
      }
      // Auto-refresh TOC page numbers from the fresh slicing, like Word's field
      // update. Gated on dirtyRef — our pagination approximates Word's, so a
      // pristine open keeps the file's numbers. History-exempt: a field result
      // change is not an edit to undo.
      if (editor && dirtyRef.current && slices.length > 0) {
        const nums = secList ? pageNumbers(slices, secList) : slices.map((_, n) => n + 1)
        const byEl = new Map(blocks.filter((b) => b.el).map((b) => [b.el as HTMLElement, b.top]))
        const headings = collectHeadings(editor.state.doc)
        // formatted with the owning section's pgNumType, like the header/footer numbers
        const displays = headings.map((h) => {
          const dom = editor.view.nodeDOM(h.pos) as HTMLElement | null
          const top = dom ? byEl.get(dom) : undefined
          if (top === undefined) return undefined
          const idx = Math.min(Math.max(pageAt(slices, top + 1), 1), nums.length)
          const fmt =
            secList?.[Math.min(slices[idx - 1].section, secList.length - 1)]?.pageNumberFmt
          return formatPageNumber(nums[idx - 1] ?? idx, fmt)
        })
        const tr = editor.state.tr
        if (applyTocPageDisplays(editor.state.doc, tr, headings, displays)) {
          tr.setMeta('addToHistory', false)
          editor.view.dispatch(tr)
        }
      }
      // M4 always-on pagination in the canvas: render page gaps before page-leading blocks
      // (print view only; gaps don't count as content — measureBlocks subtracts them, so
      // slice results are gap-independent and refresh is idempotent)
      let tGapsBuild: number | undefined
      let tSetGaps: number | undefined
      const tGaps0 = performance.now()
      if (editor) {
        const gaps: PageGapSpec[] = []
        const gapIds = new Set<string>()
        if (viewMode === 'print' && !readMode) {
          const pmRect = pm.getBoundingClientRect()
          const pageNotes = pageFootnotesOf(blocks, slices)
          // per-page header/footer for the gaps (previous page's footer + next page's
          // header), variant-selected like the pagination preview (first page / odd-even /
          // per-section references), with real page numbers for the '#' marker
          const nums = secList ? pageNumbers(slices, secList) : slices.map((_, n) => n + 1)
          const firsts = sectionFirstPages(slices)
          const effRefs = secList ? effectiveHfRefs(secList) : null
          const parsed = doc.parsed
          const pageHfOf = (
            pageIdx: number,
            kind: 'header' | 'footer',
          ): { value: HeaderFooter | null; images?: HfImage[] } => {
            const pageNo = nums[pageIdx]
            if (!secList || secList.length <= 1) {
              if (titlePg && pageIdx === 0) {
                return kind === 'header'
                  ? { value: hfVariants.headerFirst, images: parsed.headerFirst?.images }
                  : { value: hfVariants.footerFirst, images: parsed.footerFirst?.images }
              }
              if (evenOddHf && pageNo % 2 === 0) {
                return kind === 'header'
                  ? { value: hfVariants.headerEven, images: parsed.headerEven?.images }
                  : { value: hfVariants.footerEven, images: parsed.footerEven?.images }
              }
              return kind === 'header'
                ? { value: header, images: parsed.headerImages ?? undefined }
                : { value: footer, images: parsed.footerImages ?? undefined }
            }
            const pageSlice = slices[pageIdx]
            const sec = secList[Math.min(pageSlice.section, secList.length - 1)]
            const refs = effRefs![Math.min(pageSlice.section, effRefs!.length - 1)]
            const variant =
              sec.titlePg && firsts[pageIdx]
                ? 'first'
                : evenOddHf && pageNo % 2 === 0
                  ? 'even'
                  : 'default'
            const ov = variant === 'default' ? sectionHfOverride(pageSlice.section, kind) : null
            const rId = refs[kind][variant]
            return {
              value: ov ?? hfFromPart(rId ? parsed.hfParts?.[rId] : null),
              images: rId ? parsed.hfParts?.[rId]?.images : undefined,
            }
          }
          const pageNoTextOf = (pageIdx: number) =>
            formatPageNumber(
              nums[pageIdx],
              secList?.[Math.min(slices[pageIdx].section, secList.length - 1)]?.pageNumberFmt,
            )
          const hfSig = (v: HeaderFooter | null | undefined) =>
            v ? `${v.text}·${v.pageNumber ? 1 : 0}·${v.paras?.length ?? 0}` : ''
          const visiblePages = visiblePageCount(slices)
          // stopgap until per-section canvas geometry: a landscape section's header/footer
          // strip must not overflow the (portrait) canvas paper it is drawn on
          const canvasPaperW = pmRect.width / factor
          slices.slice(1).forEach((slice, k) => {
            // an even/odd section's zero-height blank page shares its start with the following page: draw only one gap band on the canvas
            if (slice.start === slices[k].start) return
            // gap = previous page's (its section's) bottom margin + inter-page band + this page's (its section's) top margin
            const prevSec = secList?.[slices[k].section]?.settings ?? section
            const nextSec = secList?.[slice.section]?.settings ?? section
            // effective margins after header/footer push-down (an over-tall header pushes the body down)
            const nextHf = hfHs?.[slice.section] ?? singleHfPx
            const prevHf = hfHs?.[slices[k].section] ?? singleHfPx
            const metrics = {
              marginTop: effectiveTopPx(nextSec, nextHf.headerPx),
              marginBottom: effectiveBottomPx(prevSec, prevHf.footerPx),
              marginLeft: twipsToPx(nextSec.marginLeft),
              marginRight: twipsToPx(nextSec.marginRight),
            }
            // previous page's footer (bottom-margin band) + next page's header (top-margin
            // band), so the canvas shows headers/footers on every page like Word
            const gapFooter = pageHfOf(k, 'footer')
            const gapHeader = pageHfOf(k + 1, 'header')
            const hfEls: HTMLElement[] = []
            if (hfHasVisibleContent(gapFooter.value, gapFooter.images)) {
              const box = sectionPageBox(prevSec)
              const el = makeGapHfEl({
                kind: 'footer',
                value: gapFooter.value ?? { text: '' },
                images: gapFooter.images,
                pageNo: pageNoTextOf(k),
                pageTotal: visiblePages,
              })
              el.style.top = 'auto'
              el.style.bottom = `${GAP_BAND + metrics.marginTop + box.footerDist}px`
              el.style.width = `${Math.min(box.width, canvasPaperW) - 120}px`
              hfEls.push(el)
            }
            if (hfHasVisibleContent(gapHeader.value, gapHeader.images)) {
              const box = sectionPageBox(nextSec)
              const el = makeGapHfEl({
                kind: 'header',
                value: gapHeader.value ?? { text: '' },
                images: gapHeader.images,
                pageNo: pageNoTextOf(k + 1),
                pageTotal: visiblePages,
              })
              el.style.bottom = 'auto'
              el.style.top = `calc(100% - ${Math.max(0, metrics.marginTop - box.headerDist)}px)`
              el.style.width = `${Math.min(box.width, canvasPaperW) - 120}px`
              hfEls.push(el)
            }
            const hfProps =
              hfEls.length > 0
                ? {
                    hfEls,
                    // key must cover everything baked into the widgets (both pages'
                    // formatted numbers + total count), or stale PAGE/NUMPAGES survive reuse
                    hfKey: `${pageNoTextOf(k)}·${pageNoTextOf(k + 1)}·${visiblePages}·${hfSig(gapFooter.value)}·${hfSig(gapHeader.value)}`,
                  }
                : {}
            // previous page's footnotes: rendered into the top of the gap (page-bottom area), with the gap enlarged by the reserved height.
            // in-table gaps carry no footnote area (absolute positioning inside a table-row is unreliable); footnotes stay in the end-of-document list
            const items = pageNotes[k] ?? []
            let notes: HTMLElement | undefined
            let notesKey: string | undefined
            let notesMetrics = metrics
            if (items.length > 0) {
              const fnH = items.reduce((s, n) => s + n.height, 0) + FOOTNOTE_SEPARATOR_H
              const contentW = twipsToPx(
                prevSec.pageWidth - prevSec.marginLeft - prevSec.marginRight,
              )
              notes = makeGapNotesEl(items, twipsToPx(prevSec.marginLeft), contentW, fnH, (id) =>
                editNote('footnote', id),
              )
              notesKey = `${Math.round(fnH)}:${items.map((n) => `${n.no}${n.text}`).join('|')}`
              notesMetrics = { ...metrics, marginBottom: metrics.marginBottom + fnH }
            }
            const markShown = () => items.forEach((n) => gapIds.add(n.id))
            const i = blocks.findIndex((b) => Math.abs(b.top - slice.start) < 0.5)
            if (i >= 0 && blocks[i].el) {
              gaps.push({
                el: blocks[i].el!,
                metrics: notesMetrics,
                ...(notes ? { notes, notesKey } : {}),
                ...hfProps,
              })
              markShown()
              return
            }
            // mid-paragraph page break (line-level cut point): insert an inline gap at the broken line. In-table cut points are not decorated yet
            const b = blocks.find(
              (bb) => bb.el && bb.top < slice.start && slice.start < bb.top + bb.height - 0.5,
            )
            if (!b?.el) return
            if (b.el.querySelector('tr')) {
              // in-table cut point: insert an in-table gap row (display:table-row widget)
              // before the broken row (next page's first row). Positioning must subtract the
              // gap's own height: a tr's DOM offset includes in-table gaps above it, so
              // subtract them before comparing with the slice's gapless virtual coordinates
              const gapRects = Array.from(b.el.querySelectorAll('.page-gap-inline')).map((g) =>
                g.getBoundingClientRect(),
              )
              const elTop = b.el.getBoundingClientRect().top
              let matched = false
              let cutRow: Element | null = null
              for (const tr of Array.from(b.el.querySelectorAll('tr')).filter(
                (r) =>
                  !r.closest('.doc-nested-table') && !r.classList.contains('page-repeat-header'),
              )) {
                const trTop = tr.getBoundingClientRect().top
                const gapsAbove = gapRects.reduce((s, g) => (g.top <= trTop ? s + g.height : s), 0)
                const off = (trTop - elTop - gapsAbove) / factor
                // last real row starting at/above the cut = the row the cut falls inside
                if (!tr.classList.contains('page-gap') && off <= slice.start - b.top + 0.5)
                  cutRow = tr
                if (Math.abs(off - (slice.start - b.top)) < 1.5) {
                  matched = true
                  try {
                    const $pos = editor.view.state.doc.resolve(editor.view.posAtDOM(tr, 0))
                    // w:tblHeader repetition: the engine reserved slice.repeatHeader.height
                    // at the top of this page's column, so cloning the source header rows
                    // below the gap fills exactly that space. Clones are decorations —
                    // page-gap-inline keeps them out of the virtual coordinates.
                    let repeatHeaderEls: HTMLElement[] | undefined
                    if (slice.repeatHeader) {
                      const tableEl = tr.closest('table')
                      const srcRows = tableEl
                        ? (Array.from(tableEl.querySelectorAll(':scope > tbody > tr')).filter(
                            (r) =>
                              !r.classList.contains('page-gap') &&
                              !r.classList.contains('page-repeat-header'),
                          ) as HTMLElement[])
                        : []
                      const els: HTMLElement[] = []
                      let acc = 0
                      for (const row of srcRows) {
                        if (acc >= slice.repeatHeader.height - 1.5) break
                        const clone = row.cloneNode(true) as HTMLElement
                        clone.classList.add('page-gap-inline', 'page-repeat-header')
                        clone.setAttribute('contenteditable', 'false')
                        els.push(clone)
                        acc += row.getBoundingClientRect().height / factor
                      }
                      if (els.length > 0) repeatHeaderEls = els
                    }
                    for (let d = $pos.depth; d > 0; d--) {
                      if ($pos.node(d).type.name === 'docTableRow') {
                        gaps.push({
                          pos: $pos.before(d),
                          kind: 'table',
                          metrics,
                          ...hfProps,
                          ...(repeatHeaderEls
                            ? {
                                repeatHeaderEls,
                                // content signature: header edits with unchanged height
                                // must still rebuild the widgets (same rule as hfKey)
                                repeatHeaderKey: `${repeatHeaderEls.length}-${Math.round(slice.repeatHeader!.height)}-${hashStr(repeatHeaderEls.map((e) => e.innerHTML).join('§'))}`,
                              }
                            : {}),
                        })
                        break
                      }
                    }
                  } catch {
                    /* if posAtDOM fails, skip decorating this round */
                  }
                  break
                }
              }
              if (!matched) {
                // in-row cut point (page break between a cell's lines): anchor at the first line after it
                const anchor = nextLineAnchor(b.el, slice.start - b.top, factor)
                const pos = anchor ? posFromAnchor(editor.view, anchor) : undefined
                if (anchor == null || pos == null) return
                const cutCell = singleCutCell(cutRow)
                if (cutCell) {
                  // single-column row: insert a real inline gap band; bleed to the paper
                  // edges from the anchor's block (the widget's containing block)
                  const r = (
                    anchor.node.parentElement?.closest('td > *, th > *') ?? cutCell
                  ).getBoundingClientRect()
                  gaps.push({
                    pos,
                    kind: 'cell',
                    metrics: {
                      ...metrics,
                      marginLeft: metrics.marginLeft + (r.left - pmRect.left) / factor,
                      marginRight: metrics.marginRight + (pmRect.right - r.right) / factor,
                    },
                    ...hfProps,
                  })
                } else {
                  // multi-cell row: keep the zero-height dashed marker (no hfProps — it can't host header/footer strips)
                  gaps.push({ pos, kind: 'cut', metrics })
                }
              }
              return
            }
            const anchor = lineStartAnchor(b.el, slice.start - b.top, factor)
            const pos = anchor ? posFromAnchor(editor.view, anchor) : undefined
            if (pos == null) return
            // fold the block's horizontal indent relative to the content area into negative margins so the gap spans the full paper width
            const elRect = b.el.getBoundingClientRect()
            gaps.push({
              pos,
              metrics: {
                ...notesMetrics,
                marginLeft: metrics.marginLeft + (elRect.left - pmRect.left) / factor,
                marginRight: metrics.marginRight + (pmRect.right - elRect.right) / factor,
              },
              ...(notes ? { notes, notesKey } : {}),
              ...hfProps,
            })
            markShown()
          })
          // TOC page numbers: the file's cached PAGEREF results are stale (generators
          // write them against a layout that never matches; Word silently refreshes on
          // open, we never write back). Backfill the display from the live layout —
          // DOM-only, inside contenteditable=false subtrees the save path never reads.
          const tocLines = pm.querySelectorAll<HTMLElement>('.doc-toc-line[data-toc-anchor]')
          if (tocLines.length > 0) {
            const anchorIdx = new Map<string, number>()
            for (const b of parsed.blocks) {
              if (b.docxIndex == null) continue
              for (const a of b.hiddenBookmarks ?? []) anchorIdx.set(a, b.docxIndex)
            }
            const topByIdx = new Map<number, number>()
            for (const b of blocks) {
              if (b.docxIndex != null) topByIdx.set(b.docxIndex, b.top)
            }
            for (const el of tocLines) {
              const idx = anchorIdx.get(el.getAttribute('data-toc-anchor') ?? '')
              const top = idx === undefined ? undefined : topByIdx.get(idx)
              if (top === undefined) continue
              const pageEl = el.querySelector('.doc-toc-page')
              // pageAt is 1-based; pageNoTextOf indexes nums/slices 0-based
              const pageIdx = Math.max(0, Math.min(pageAt(slices, top + 1) - 1, slices.length - 1))
              if (pageEl && slices.length > 0) pageEl.textContent = pageNoTextOf(pageIdx)
            }
          }
        }
        tGapsBuild = performance.now() - tGaps0
        const tSet0 = performance.now()
        setPageGaps(editor.view, gaps)
        tSetGaps = performance.now() - tSet0
        // the last page paints as a full sheet like the ones above it:
        // extend the canvas to that page's paper bottom, measured from the last gap
        const gapEls = pm.querySelectorAll('.page-gap')
        const lastGapEl = gapEls[gapEls.length - 1]
        if (lastGapEl && slices.length > 1) {
          const last = slices[slices.length - 1]
          const lastSec = secList?.[last.section]?.settings ?? section
          const lastHf = hfHs?.[last.section] ?? singleHfPx
          const paperTop =
            (lastGapEl.getBoundingClientRect().bottom - pm.getBoundingClientRect().top) / factor -
            effectiveTopPx(lastSec, lastHf.headerPx)
          pm.style.minHeight = `${Math.round(paperTop + twipsToPx(lastSec.pageHeight))}px`
        } else {
          pm.style.removeProperty('min-height')
        }
        // for real-device verification/troubleshooting: current slices and block geometry (read-only snapshot, no functional dependency)
        ;(window as unknown as Record<string, unknown>).__pageDebug = {
          slices,
          blocks: blocks.map((b) => ({ top: b.top, height: b.height, docxIndex: b.docxIndex })),
          remeasureMs: performance.now() - tStart,
          measureMs: tMeasure,
          sliceMs: tSlice,
          gapsBuildMs: tGapsBuild,
          setGapsMs: tSetGaps,
        }
        requestAnimationFrame(() => {
          const tf = performance.now()
          requestAnimationFrame(() => {
            const dbg = (window as unknown as Record<string, Record<string, unknown>>).__pageDebug
            if (dbg) dbg.frameMs = performance.now() - tf
          })
        })
        setGapNoteIds((prev) => {
          if (prev.size === gapIds.size && [...gapIds].every((id) => prev.has(id))) return prev
          return gapIds
        })
      }
      locate()
    }
    const onUpdate = () => {
      if (timer) window.clearTimeout(timer)
      timer = window.setTimeout(remeasure, 300)
    }
    remeasure()
    // async @font-face loading triggers a full reflow (line-break points change); pagination must be remeasured
    document.fonts.ready.then(onUpdate).catch(() => {})
    document.fonts.addEventListener('loadingdone', onUpdate)
    scroller.addEventListener('scroll', locate, { passive: true })
    editor?.on('update', onUpdate)
    return () => {
      if (timer) window.clearTimeout(timer)
      document.fonts.removeEventListener('loadingdone', onUpdate)
      scroller.removeEventListener('scroll', locate)
      editor?.off('update', onUpdate)
    }
  }, [
    doc,
    section,
    sections,
    zoom,
    editor,
    viewMode,
    readMode,
    blockMetaOf,
    pageFootnotesOf,
    endnoteItems,
    editNote,
    effTopSingle,
    effBottomSingle,
    singleHfPx,
    hfHeightsOf,
    measureSingleFlow,
    colGeomsFor,
    header,
    footer,
    hfVariants,
    titlePg,
    evenOddHf,
    sectionHfOverride,
  ])

  // section at the cursor: the target the Layout tab acts on
  useEffect(() => {
    if (!editor || sections.length === 0) {
      setActiveSection(0)
      return
    }
    const locateSection = () => {
      const { $head } = editor.state.selection
      const pmDoc = editor.state.doc
      const topIndex = $head.depth > 0 ? $head.index(0) : 0
      let docxIndex: number | null = null
      for (let i = Math.min(topIndex, pmDoc.childCount - 1); i >= 0; i--) {
        const di = pmDoc.child(i).attrs?.docxIndex as number | null | undefined
        if (di !== null && di !== undefined) {
          docxIndex = di
          break
        }
      }
      const s =
        docxIndex === null ? 0 : sections.findIndex((sec) => docxIndex! <= sec.lastBlockIndex)
      setActiveSection(s >= 0 ? s : sections.length - 1)
    }
    locateSection()
    editor.on('selectionUpdate', locateSection)
    return () => {
      editor.off('selectionUpdate', locateSection)
    }
  }, [editor, sections])

  /** Word word-count dialog: pages/lines estimated from the current layout */
  const openStats = useCallback(() => {
    if (!editor) return
    const text = editor.state.doc.textContent
    const pm = document.querySelector('.ProseMirror')
    const zoomFactor = zoom / 100
    let lines = 0
    if (pm) {
      for (const el of Array.from(pm.children) as HTMLElement[]) {
        const cs = getComputedStyle(el)
        let lh = parseFloat(cs.lineHeight)
        if (!Number.isFinite(lh) || lh <= 0) lh = (parseFloat(cs.fontSize) || 15) * 1.2
        lines += Math.max(1, Math.round(el.getBoundingClientRect().height / zoomFactor / lh))
      }
    }
    // Word's paragraph figure counts non-empty paragraphs, including those
    // inside table cells (descendants, not just top-level children)
    let paragraphs = 0
    editor.state.doc.descendants((node) => {
      if (node.isTextblock && node.textContent.trim() !== '') paragraphs++
      return true
    })
    setStats({
      pages: pageInfo.total,
      words: countWords(text),
      asianChars: asianCharCount(text),
      nonAsianWords: nonAsianWordCount(text),
      charsNoSpace: text.replace(/\s/g, '').length,
      charsWithSpace: text.replace(/\n/g, '').length,
      paragraphs,
      lines,
    })
  }, [editor, zoom, pageInfo.total])

  // "has unsaved changes" check shared by the close guard and autosave; refreshed on every
  // render (all edit paths forceRender), so the guard's query reads the latest value
  const anyDirtyRef = useRef(false)
  const hasUnsavedChanges = isDocDirty(fileCtxRef.current)
  anyDirtyRef.current = hasUnsavedChanges

  // close guard: the main process queries dirty state before closing a tab/window; choosing "Save" runs a full save and reports back
  useEffect(() => {
    const offCheck = window.desktop.onCloseCheck?.(() => {
      window.desktop.reportCloseCheck({
        dirty: !!doc && (anyDirtyRef.current || dirtyRef.current),
        autoSave: autoSave && !!doc?.filePath,
        filePath: doc?.filePath ?? null,
      })
    })
    const offSave = window.desktop.onCloseSaveRequest?.(() => {
      // Closing must not report success while edits are still unpersisted, so a
      // save that raced with typing is retried until the file catches up.
      // save(false) never prompts — a pathless first save lands silently in the
      // default folder — so retrying is always safe; reporting "persisted" just
      // because the snapshot had no path yet would close over mid-save edits.
      void saveUntilPersisted({
        save: () => save(false),
        wasIncomplete: () => saveIncompleteRef.current,
        hasPath: () => true,
      }).then(
        (ok) => window.desktop.reportCloseSaveResult(ok === true),
        () => window.desktop.reportCloseSaveResult(false),
      )
    })
    return () => {
      offCheck?.()
      offSave?.()
    }
  }, [doc, save, autoSave])

  // autosave: every 30s and on window blur, silently persist pending changes
  useEffect(() => {
    if (tornDown || !autoSave || !doc || !doc.filePath) return
    const tick = () => {
      if (!isDocDirty(fileCtxRef.current)) return
      if (editor?.view.composing) return // don't interrupt IME input
      const active = document.activeElement as HTMLElement | null
      if (active?.closest('td[contenteditable], .doc-textbox')) return // mid in-place edit
      void save(false, true)
    }
    const id = window.setInterval(tick, 30_000)
    window.addEventListener('blur', tick)
    return () => {
      window.clearInterval(id)
      window.removeEventListener('blur', tick)
    }
  }, [tornDown, autoSave, doc, editor, save])

  // After an AI run finishes on a never-saved document, silently save it once: the
  // first save derives the file name from the first heading (see deriveAutoFileName
  // in file-actions), which also renames the shell tab — mirrors slides, where AI
  // generation names and persists the draft deck.
  useEffect(() => {
    const handler = () => {
      const cur = fileCtxRef.current
      if (!cur.doc || cur.doc.filePath || !anyDirtyRef.current) return
      if (editor?.view.composing) return
      void save(false, true)
    }
    window.addEventListener('ai-docs-run-done', handler)
    return () => window.removeEventListener('ai-docs-run-done', handler)
  }, [editor, save])

  useEffect(() => {
    // editing shortcuts only fire when focus is in an editor surface (main
    // ProseMirror, textbox sub-editor, in-place table cell) or nowhere at all —
    // never while typing in the find box, AI input or other form fields
    const focusInEditor = () => {
      const el = document.activeElement
      if (!el || el === document.body) return true
      return !!(el as HTMLElement).closest('.ProseMirror, td[contenteditable]')
    }
    const handler = (e: KeyboardEvent) => {
      const canEdit = !!editor?.isEditable && focusInEditor()
      if ((e.metaKey || e.ctrlKey) && e.key === 's') {
        e.preventDefault()
        void save(e.shiftKey)
      }
      if ((e.metaKey || e.ctrlKey) && e.key === 'o') {
        e.preventDefault()
        void openFile()
      }
      if ((e.metaKey || e.ctrlKey) && e.key === 'f') {
        e.preventDefault()
        if (doc) setShowFind(true)
      }
      // Word dialog shortcuts: Font ⌘D / Paragraph ⌥⌘M / Hyperlink ⌘K
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey && e.key === 'd' && doc && canEdit) {
        e.preventDefault()
        setShowFontDialog(true)
      }
      if ((e.metaKey || e.ctrlKey) && e.altKey && e.code === 'KeyM' && doc && canEdit) {
        e.preventDefault()
        setShowParaDialog(true)
      }
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey && e.key === 'k' && doc && canEdit) {
        e.preventDefault()
        setShowLinkModal(true)
      }
      if (e.key === 'F9' && doc && editor?.isEditable) {
        e.preventDefault()
        updateFields()
      }
      // Word alignment shortcuts
      const ALIGN_KEYS: Record<string, 'left' | 'center' | 'right' | 'justify'> = {
        l: 'left',
        e: 'center',
        r: 'right',
        j: 'justify',
      }
      if (
        (e.metaKey || e.ctrlKey) &&
        !e.shiftKey &&
        !e.altKey &&
        e.key in ALIGN_KEYS &&
        editor &&
        canEdit
      ) {
        e.preventDefault()
        // route into a focused textbox sub-editor, like the ribbon does
        const target = getActiveSubEditor() ?? editor
        setSelectionAlign(target, ALIGN_KEYS[e.key])
      }
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [save, openFile, editor, doc, updateFields])

  // double-click an inline equation / click an equation block's edit button (ones with LaTeX source) → reopen the equation dialog for editing
  useEffect(() => {
    const handler = (e: Event) => {
      const detail = (e as CustomEvent<{ pos: number; latex: string; kind?: 'inline' | 'block' }>)
        .detail
      if (editor && doc && detail) setEqEditTarget({ kind: 'inline', ...detail })
    }
    window.addEventListener('ai-docs-edit-inline-math', handler)
    return () => window.removeEventListener('ai-docs-edit-inline-math', handler)
  }, [editor, doc])

  // native application menu → renderer commands
  useEffect(() => {
    return window.desktop.onMenuCommand((command, payload) => {
      const align = (value: 'left' | 'center' | 'right' | 'justify') =>
        editor && setSelectionAlign(editor, value)
      switch (command) {
        case 'new':
          void newFile()
          break
        case 'open':
          void openFile()
          break
        case 'open-path':
          if (payload) void openRecent(payload)
          break
        case 'save':
          void save(false)
          break
        case 'save-as':
          void save(true)
          break
        case 'undo':
          editor?.chain().focus().undo().run()
          break
        case 'redo':
          editor?.chain().focus().redo().run()
          break
        case 'zoom-in':
          setZoom((z) => Math.min(200, Math.round(z) + 10))
          break
        case 'zoom-out':
          setZoom((z) => Math.max(50, Math.round(z) - 10))
          break
        case 'zoom-100':
          setZoom(100)
          break
        case 'zoom-page-width':
          zoomFit('width')
          break
        case 'zoom-whole-page':
          zoomFit('page')
          break
        case 'toggle-ai':
          setShowAi((v) => !v)
          break
        case 'toggle-dark':
          setDarkCanvas((v) => !v)
          break
        case 'insert-table':
          if (editor && doc) insertTableAt(editor, 3, 3)
          break
        case 'insert-image':
          if (editor && doc) void insertImageViaDialog(editor)
          break
        case 'insert-page-break':
          if (editor && doc) insertPageBreakAt(editor)
          break
        case 'insert-link':
          if (editor && doc) setShowLinkModal(true)
          break
        case 'insert-equation':
          if (editor && doc) setShowEquationModal(true)
          break
        // Menu parity with the shortcuts and the right-click menu
        case 'insert-comment':
          if (editor && doc) {
            setShowComments(true)
            startNewComment()
          }
          break
        case 'font-dialog':
          if (doc) setShowFontDialog(true)
          break
        case 'paragraph-dialog':
          if (doc) setShowParaDialog(true)
          break
        case 'word-count':
          if (doc) openStats()
          break
        case 'bold':
          editor?.chain().focus().toggleMark('bold').run()
          break
        case 'italic':
          editor?.chain().focus().toggleMark('italic').run()
          break
        case 'underline':
          editor?.chain().focus().toggleMark('underline').run()
          break
        case 'align-left':
          align('left')
          break
        case 'align-center':
          align('center')
          break
        case 'align-right':
          align('right')
          break
        case 'align-justify':
          align('justify')
          break
        case 'page-setup':
          setRibbonTabRequest({ tab: 'layout', nonce: Date.now() })
          break
        case 'find':
          if (doc) setShowFind(true)
          break
        case 'export-pdf':
          void exportPdf()
          break
      }
    })
  }, [
    editor,
    doc,
    newFile,
    openFile,
    openRecent,
    save,
    exportPdf,
    zoomFit,
    openStats,
    startNewComment,
  ])

  // Word: TOC entries jump on ⌘/Ctrl+click only; a plain click just places the
  // caret. Resolve the bookmark anchor against the original block XML, fall
  // back to matching the heading text.
  const onDocClick = useCallback(
    (e: ReactMouseEvent) => {
      const commentSpan = (e.target as HTMLElement).closest('.doc-comment') as HTMLElement | null
      if (commentSpan) {
        const ids = (commentSpan.dataset.commentIds ?? '').split(' ')
        if (comments.some((c) => !c.done && ids.includes(c.id))) setShowComments(true)
      }
      if (!e.metaKey && !e.ctrlKey) return
      const line = (e.target as HTMLElement).closest('.doc-toc-line') as HTMLElement | null
      if (!line) return
      let target: Element | null = null
      const anchor = line.dataset.tocAnchor
      if (anchor && doc) {
        const block = doc.parsed.blocks.find((b) => b.originalXml?.includes(`w:name="${anchor}"`))
        if (block && block.docxIndex !== null) {
          target = document.querySelector(`.ProseMirror [data-idx="${block.docxIndex}"]`)
        }
      }
      if (!target) {
        const title = (line.dataset.tocTitle ?? '').replace(/\s+/g, '')
        if (title) {
          target =
            [
              ...document.querySelectorAll(
                '.ProseMirror h1, .ProseMirror h2, .ProseMirror h3, .ProseMirror h4, .ProseMirror h5, .ProseMirror h6',
              ),
            ].find((h) => (h.textContent ?? '').replace(/\s+/g, '') === title) ?? null
        }
      }
      target?.scrollIntoView({ behavior: 'smooth', block: 'start' })
    },
    [doc, comments],
  )

  /** Editor right-click menu, only inside the document body */
  const onDocContextMenu = useCallback(
    (e: ReactMouseEvent) => {
      if (readMode || isProtected) return
      if (!(e.target as HTMLElement).closest('.doc-page')) return
      e.preventDefault()
      // Word behavior: right-clicking outside the selection moves the cursor there first (menu items act on the clicked block)
      if (editor) {
        const hit = editor.view.posAtCoords({ left: e.clientX, top: e.clientY })
        if (hit) {
          const { from, to } = editor.state.selection
          if (hit.pos < from || hit.pos > to) {
            editor.commands.setTextSelection(hit.pos)
          }
        }
      }
      setCtxMenu({ x: e.clientX, y: e.clientY })
    },
    [readMode, isProtected, editor],
  )

  // e2e/automation hook: lets tests drive open/edit/save without native dialogs
  useEffect(() => {
    ;(window as unknown as Record<string, unknown>).__aidocs = {
      editor,
      openPath: (path: string) => openRecent(path),
      save: () => save(false),
      getStatus: () => status,
      exportPdfTo: (path: string) => exportPdf(path),
    }
  }, [editor, openRecent, save, status, exportPdf])

  // shallow-stable snapshot of every editor read the ribbon displays: caret moves
  // that change none of it keep the reference, so the memoized Ribbon skips
  const formatState = useShallowStable(
    computeFormatState(editor, doc?.parsed.styles, doc?.parsed.docDefaults),
  )

  const ribbonStyles = useMemo(() => (doc ? new Map(doc.parsed.styles) : undefined), [doc])

  /** every function prop of the memoized Ribbon, with stable identities (dispatches into the latest render's closures) */
  const ribbonActions = useStableCallbacks({
    allocateNumId: (kind: 'bullet' | 'ordered') => allocateListNumId(kind),
    createListDef: (levels: CustomNumberingLevel[]) => createCustomListDef(levels),
    onParagraphDialog: () => setShowParaDialog(true),
    onOpen: () => void openFile(),
    onSave: () => void save(false),
    onSaveAs: () => void save(true),
    onToggleAi: () => setShowAi((v) => !v),
    onSection: (next: SectionSettings) => {
      // layout applies to the cursor's section; the final section's sectPr goes through SaveOptions.section (also drives canvas geometry)
      setSections((prev) =>
        prev.map((s, i) => (i === activeSection ? { ...s, settings: next } : s)),
      )
      if (sections.length <= 1 || activeSection === sections.length - 1) {
        setSection(next)
        setSectionDirty(true)
      } else {
        setSectionsDirty((d) => (d.includes(activeSection) ? d : [...d, activeSection]))
        setStatus(t('appSectionSettingsApplied', { n: activeSection + 1 }))
      }
    },
    onInsertSectionBreak: (type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage') =>
      insertSectionBreak(type),
    onPageColor: (next: string | null) => {
      setPageColor(next)
      setPageColorDirty(true)
    },
    onWatermark: (next: string | null) => {
      setWatermark(next)
      setWatermarkDirty(true)
      setStatus(next ? t('appWatermarkSet', { text: next }) : t('appWatermarkRemoved'))
    },
    onThemeFonts: (fonts: ThemeFonts) => {
      setThemeFonts(fonts)
      setThemeFontsDirty(true)
      setStatus(t('appThemeFontsChanged'))
    },
    onThemeColors: (colors: ThemeColors) => {
      setThemeColors(colors)
      setThemeColorsDirty(true)
      setStatus(t('appThemeColorsApplied', { name: colors.name ?? '' }))
    },
    onInkTool: setInkTool,
    onInkPen: setInkPen,
    onInkHighlighter: setInkHighlighter,
    onInkClearAll: clearInks,
    onInsertNote: insertNote,
    onAddSource: (source: SourceInfo) => {
      setSources((prev) => [...prev, source])
      setSourcesDirty(true)
      setStatus(t('appSourceAdded', { title: source.title }))
    },
    headingPages,
    onZoom: setZoom,
    onZoomFit: zoomFit,
    onDarkCanvas: setDarkCanvas,
    onAiPreset: (text: string) => {
      // Word's Editor / Translate start working as soon as they're clicked
      setShowAi(true)
      setAiPreset({ text, nonce: Date.now(), autoRun: true })
    },
    onHeader: (next: HeaderFooter) => {
      setHeader(next)
      setHeaderDirty(true)
    },
    onPageNumFormat: openPgNumModal,
    onInsertField: insertField,
    onFooter: (next: HeaderFooter) => {
      setFooter(next)
      setFooterDirty(true)
    },
    onTitlePg: toggleTitlePg,
    onEvenOddHf: (on: boolean) => {
      setEvenOddHf(on)
      setEvenOddHfDirty(true)
      setHfView(on ? 'even' : 'default')
      setStatus(on ? t('appEvenOddOn') : t('appEvenOddOff'))
    },
    onShowMarks: setShowMarks,
    onShowRuler: setShowRuler,
    onShowNav: setShowNav,
    onShowComments: () => setShowComments(true),
    onNewComment: startNewComment,
    onTrackChanges: setTrackChanges,
    onRevisionDisplay: setRevisionDisplay,
    onAcceptRevision: (all: boolean) => handleRevision('accept', all),
    onRejectRevision: (all: boolean) => handleRevision('reject', all),
    onGotoRevision: (dir: 1 | -1) => {
      if (editor) gotoRevision(editor, dir)
    },
    onToggleProtection: toggleProtection,
    onCompare: () => void compareWithFile(),
    onViewMode: setViewMode,
    onReadMode: setReadMode,
    onShowGrid: setShowGrid,
    onSplitView: setSplitView,
    onPagePreview: () => setShowPagePreview(true),
  })

  const closeCommentsPanel = useCallback(() => {
    setShowComments(false)
    cancelNewComment()
  }, [cancelNewComment])

  const hasDoc = !!doc
  const quickActions = useMemo(
    () => (
      <>
        <button
          className="qa-btn"
          title={t('appSaveShortcutTip')}
          disabled={!hasDoc || !hasUnsavedChanges}
          onClick={() => void save(false)}
        >
          <IconSave size={16} />
        </button>
        <button
          className="qa-btn"
          title={t('appUndo')}
          disabled={!hasDoc}
          onClick={() => editor?.chain().focus().undo().run()}
        >
          <IconUndo size={16} />
        </button>
        <button
          className="qa-btn"
          title={t('appRedo')}
          disabled={!hasDoc}
          onClick={() => editor?.chain().focus().redo().run()}
        >
          <IconRedo size={16} />
        </button>
        <label className={`autosave-toggle ${autoSave ? 'on' : ''}`} title={t('appAutoSaveTip')}>
          <span className="autosave-knob" />
          <span className="autosave-text">{t('appAutoSave')}</span>
          <input
            type="checkbox"
            checked={autoSave}
            onChange={(e) => setAutoSave(e.target.checked)}
          />
        </label>
        <span className="qa-sep" aria-hidden="true" />
      </>
    ),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [hasDoc, hasUnsavedChanges, autoSave, editor, save, lang],
  )

  if (!editor) return null

  const wordCount = wordCountOfDoc(editor.state.doc)

  const canvasSection = sections[activeSection]?.settings ?? section
  const canvasBox = canvasSection ? sectionPageBox(canvasSection) : null
  const canvasTop = canvasSection ? effectiveTopPx(canvasSection, singleHfPx.headerPx) : 0
  const canvasBottom = canvasSection ? effectiveBottomPx(canvasSection, singleHfPx.footerPx) : 0
  const docZoomStyle = {
    zoom: zoom / 100,
    '--page-w': canvasBox ? `${canvasBox.width}px` : undefined,
    '--page-h': canvasBox ? `${canvasBox.height}px` : undefined,
    '--section-content-w': canvasBox ? `${canvasBox.contentWidth}px` : undefined,
    '--page-pad': canvasSection
      ? `${canvasTop}px ${twipsToPx(canvasSection.marginRight)}px ${canvasBottom}px ${twipsToPx(canvasSection.marginLeft)}px`
      : undefined,
    '--header-dist': canvasBox ? `${canvasBox.headerDist}px` : undefined,
    '--footer-dist': canvasBox ? `${canvasBox.footerDist}px` : undefined,
    '--page-bg': pageColor ? `#${pageColor}` : undefined,
    '--page-cols': colFlow && viewMode === 'print' ? colFlow.cols : undefined,
  } as CSSProperties

  const docZoomClass = [
    'doc-zoom',
    showMarks ? 'show-marks' : '',
    `view-${viewMode}`,
    showGrid ? 'show-grid' : '',
  ].join(' ')

  return (
    <div
      className={`app ${readMode ? 'read-mode' : ''}${revisionDisplay !== 'all' ? ` rev-display-${revisionDisplay}` : ''}`}
    >
      <ToastHost />
      {docCss && <style>{docCss}</style>}
      {doc && liveDocCjk != null && (
        <style>{`.doc-page { --doc-line-factor:${docLineFactor(doc.parsed, liveDocCjk)} }`}</style>
      )}
      {doc && gridPitchPt != null && (
        // typed w:docGrid: line-height round(up) expressions snap to this pitch
        <style>{`.doc-page { --doc-grid-pitch:${gridPitchPt}pt }`}</style>
      )}
      {doc && section && (
        // over-wide tables may spill into the right margin (Word/LO), capped at the paper edge
        <style>{`.doc-page { --doc-margin-right:${twipsToPx(section.marginRight)}px }`}</style>
      )}
      {/* Theme CSS comes from live state, so a Design ▸ Themes/Fonts/Colors pick shows
          on the page immediately instead of only in the saved file */}
      {doc && <style>{docThemeCss(themeFonts, themeColors, !!docBodyFont(doc.parsed))}</style>}
      {colFlow && viewMode === 'print' && (
        // columns (sectPr w:cols): column gap follows the document's w:space; measuring-columns
        // is the single-flow measuring state (columns removed, content-box width = column width,
        // toggled instantaneously for measurement, invisible).
        // .doc-page is border-box, so the measured width must add back the left/right margin padding
        <style>{`.editor-scroll .doc-page { column-count: ${colFlow.cols}; column-gap: ${colFlow.gapPx}px; column-fill: balance; }
.editor-scroll .doc-page.measuring-columns { column-count: auto; width: ${colFlow.colWidthPx + twipsToPx(canvasSection?.marginLeft ?? section?.marginLeft ?? 0) + twipsToPx(canvasSection?.marginRight ?? section?.marginRight ?? 0)}px; }`}</style>
      )}
      <Ribbon
        quickActions={quickActions}
        editor={editor}
        formatState={formatState}
        hasDoc={!!doc}
        blocks={doc?.parsed.blocks ?? EMPTY_BLOCKS}
        styles={ribbonStyles}
        docDefaults={doc?.parsed.docDefaults}
        showAi={showAi}
        section={sections[activeSection]?.settings ?? section}
        activeSection={sections.length > 1 ? activeSection : null}
        pageColor={pageColor}
        watermark={watermark}
        themeFonts={themeFonts}
        themeColors={themeColors}
        inkTool={inkTool}
        inkPen={inkPen}
        inkHighlighter={inkHighlighter}
        inkCount={inkAnnotations.length}
        sources={sources}
        zoom={Math.round(zoom)}
        darkCanvas={darkCanvas}
        tabRequest={ribbonTabRequest}
        header={header}
        footer={footer}
        titlePg={titlePg}
        evenOddHf={evenOddHf}
        showMarks={showMarks}
        showRuler={showRuler}
        showNav={showNav}
        commentCount={comments.length}
        canComment={!editor.state.selection.empty}
        trackChanges={trackChanges}
        revisionDisplay={revisionDisplay}
        revisionCount={revisionCount}
        isProtected={isProtected}
        filePath={doc?.filePath ?? null}
        viewMode={viewMode}
        readMode={readMode}
        showGrid={showGrid}
        splitView={splitView}
        {...ribbonActions}
      />

      <div className="app-main">
        {doc && (
          <div className={`ai-dock${showAi ? '' : ' collapsed'}`}>
            {/* always mounted: collapse must not drop state or in-flight runs */}
            <AiPanel
              key={aiPanelKey}
              editor={editor}
              blocks={doc.parsed.blocks}
              settings={settings}
              docEmpty={wordCount === 0}
              numIdFallback={
                doc.isBlank ? { bullet: BLANK_BULLET_NUM_ID, ordered: BLANK_ORDERED_NUM_ID } : null
              }
              preset={aiPreset}
              open={showAi}
              onExpand={() => setShowAi(true)}
              onCollapse={() => setShowAi(false)}
              filePath={doc?.filePath ?? null}
            />
          </div>
        )}
        <div className="app-content">
          <div className={`workspace ${darkCanvas ? 'workspace-dark' : ''}`}>
            {doc && showFind && <FindPanel editor={editor} onClose={() => setShowFind(false)} />}
            {doc && showNav && <NavPane editor={editor} doc={editor.state.doc} />}
            <div className="editor-area">
              <main className="editor-scroll">
                {doc ? (
                  <div
                    className={docZoomClass}
                    onClick={onDocClick}
                    onContextMenu={onDocContextMenu}
                    style={docZoomStyle}
                  >
                    {showRuler && section && (
                      <Ruler
                        section={section}
                        editor={editor}
                        onTabStopsChange={(stops) => {
                          if (!editor) return
                          editor
                            .chain()
                            .focus()
                            .updateAttributes('docParagraph', {
                              tabStops: stops ? JSON.stringify(stops) : null,
                            })
                            .updateAttributes('docHeading', {
                              tabStops: stops ? JSON.stringify(stops) : null,
                            })
                            .updateAttributes('docListItem', {
                              tabStops: stops ? JSON.stringify(stops) : null,
                            })
                            .run()
                        }}
                      />
                    )}
                    <div className={`page-wrap ${section?.pageBorder ? 'page-bordered' : ''}`}>
                      {watermark && (
                        <div className="page-watermark" aria-hidden="true">
                          {watermark}
                        </div>
                      )}
                      {!readMode && (
                        <div
                          className={`hf-variant-chips${titlePg || evenOddHf ? '' : ' hf-chips-idle'}`}
                        >
                          {/* Always-available entry point: the ribbon checkbox alone was
                          undiscoverable while editing the header, so the toggle also lives here
                          (revealed on header hover until enabled) */}
                          <button
                            className={`hf-first-toggle${titlePg ? ' on' : ''}`}
                            title={t('ribbonDiffFirstPageTip')}
                            onClick={() => toggleTitlePg(!titlePg)}
                          >
                            {titlePg ? '✓ ' : ''}
                            {t('ribbonDiffFirstPage')}
                          </button>
                          {titlePg && (
                            <button
                              className={headerAreaView === 'first' ? 'on' : ''}
                              onClick={() => setHfView('first')}
                            >
                              {t('appFirstPage')}
                            </button>
                          )}
                          {(titlePg || evenOddHf) && (
                            <button
                              className={headerAreaView === 'default' ? 'on' : ''}
                              onClick={() => setHfView('default')}
                            >
                              {evenOddHf ? t('appOddPage') : t('appDefaultPage')}
                            </button>
                          )}
                          {evenOddHf && (
                            <button
                              className={headerAreaView === 'even' ? 'on' : ''}
                              onClick={() => setHfView('even')}
                            >
                              {t('appEvenPage')}
                            </button>
                          )}
                        </div>
                      )}
                      {(multiHf ||
                        (hfViewTouched && effHfView !== 'default') ||
                        shownHeader?.text ||
                        shownHeader?.paras?.length ||
                        hfImagesOf('header')?.length) && (
                        <HeaderFooterArea
                          kind="header"
                          value={shownHeader ?? { text: '' }}
                          images={hfImagesOf('header')}
                          readOnly={isProtected || readMode}
                          onCommit={(next) => commitHf('header', next)}
                          pageTotal={pageInfo.total}
                        />
                      )}
                      <EditorContent editor={editor} />
                      {footnotes.some((n) => !gapNoteIds.has(n.id)) && (
                        <div className="page-notes">
                          {/* footnotes already shown per page in page gaps aren't repeated at the end (last page's footnotes still live here) */}
                          {footnotes
                            .filter((n) => !gapNoteIds.has(n.id))
                            .map((n) => (
                              <div key={`f${n.id}`} className="page-note">
                                <sup>{footnotes.indexOf(n) + 1}</sup>
                                <span className="page-note-text">{n.text}</span>
                                <button
                                  className="page-note-btn"
                                  title={t('appEditFootnote')}
                                  onClick={() => editNote('footnote', n.id)}
                                >
                                  ✎
                                </button>
                                <button
                                  className="page-note-btn"
                                  title={t('appDeleteFootnote')}
                                  onClick={() => deleteNote('footnote', n.id)}
                                >
                                  ×
                                </button>
                              </div>
                            ))}
                        </div>
                      )}
                      {endnotes.length > 0 && (
                        // endnotes live in their own document-end area, roman-numbered — never mixed into the footnote block
                        <div className="page-notes page-endnotes">
                          <div className="page-notes-label">{t('appEndnotesLabel')}</div>
                          {endnotes.map((n, i) => (
                            <div key={`e${n.id}`} className="page-note">
                              <sup>{toRoman(i + 1)}</sup>
                              <span className="page-note-text">{n.text}</span>
                              <button
                                className="page-note-btn"
                                title={t('appEditEndnote')}
                                onClick={() => editNote('endnote', n.id)}
                              >
                                ✎
                              </button>
                              <button
                                className="page-note-btn"
                                title={t('appDeleteEndnote')}
                                onClick={() => deleteNote('endnote', n.id)}
                              >
                                ×
                              </button>
                            </div>
                          ))}
                        </div>
                      )}
                      {(multiHf ||
                        (hfViewTouched && effHfView !== 'default') ||
                        shownFooter?.text ||
                        shownFooter?.pageNumber ||
                        shownFooter?.paras?.length ||
                        hfImagesOf('footer')?.length) && (
                        <HeaderFooterArea
                          kind="footer"
                          value={shownFooter ?? { text: '' }}
                          images={hfImagesOf('footer')}
                          readOnly={isProtected || readMode}
                          onCommit={(next) => commitHf('footer', next)}
                          pageNo={lastPageNo?.text ?? pageInfo.total}
                          pageTotal={pageInfo.total}
                        />
                      )}
                      {(inkAnnotations.length > 0 || inkTool !== 'select') && !readMode && (
                        <InkOverlay
                          tool={isProtected ? 'select' : inkTool}
                          color={inkTool === 'highlighter' ? inkHighlighter.color : inkPen.color}
                          width={inkTool === 'highlighter' ? inkHighlighter.width : inkPen.width}
                          zoom={zoom}
                          annotations={inkAnnotations}
                          onAdd={addInk}
                          onRemove={removeInks}
                        />
                      )}
                    </div>
                  </div>
                ) : (
                  <div className="start-screen start-booting">{t('appStartOpening')}</div>
                )}
              </main>
              {doc && splitView && (
                <div className="split-pane">
                  <div className="split-pane-bar">
                    <span>{t('appSplitPaneLabel')}</span>
                    <button
                      className="split-pane-close"
                      title={t('appRemoveSplit')}
                      onClick={() => setSplitView(false)}
                    >
                      ×
                    </button>
                  </div>
                  <div className="split-pane-scroll">
                    <div className={docZoomClass} style={docZoomStyle}>
                      <div className="page-wrap">
                        <div
                          className="doc-page ProseMirror split-doc"
                          dangerouslySetInnerHTML={{ __html: splitHtml }}
                        />
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
            {readMode && (
              <button className="read-exit" onClick={() => setReadMode(false)}>
                {t('appExitReadMode')}
              </button>
            )}
            {doc && showComments && (
              <CommentsPanel
                comments={comments}
                docNode={editor.state.doc}
                composing={commentComposing}
                onSubmitNew={submitNewComment}
                onReply={replyToComment}
                onResolve={resolveComment}
                onCancelNew={cancelNewComment}
                onDelete={deleteComment}
                onClose={closeCommentsPanel}
              />
            )}
            {doc && compareResult && (
              <ComparePanel
                otherName={compareResult.otherName}
                entries={compareResult.entries}
                onClose={() => setCompareResult(null)}
              />
            )}
          </div>

          <footer className="status-bar">
            <div className="status-left">
              {doc && (
                <>
                  <span className="status-item">
                    {t('appPageOf', { current: pageInfo.current, total: pageInfo.total })}
                  </span>
                  <button
                    className="status-item status-wordcount"
                    title={t('appWordCountTitle')}
                    onClick={openStats}
                  >
                    {t('appWordCountN', { n: wordCount })}
                  </button>
                </>
              )}
              {!doc && t('appReady')}
              {status && <span className="status-msg"> — {status}</span>}
            </div>
            <div className="status-right">
              <button
                className="zoom-btn"
                onClick={() => setZoom((z) => Math.max(50, Math.round(z) - 10))}
              >
                −
              </button>
              <input
                className="zoom-slider"
                type="range"
                min={50}
                max={200}
                step={10}
                value={Math.round(zoom)}
                onChange={(e) => setZoom(Number(e.target.value))}
              />
              <button
                className="zoom-btn"
                onClick={() => setZoom((z) => Math.min(200, Math.round(z) + 10))}
              >
                +
              </button>
              <span className="zoom-value">{Math.round(zoom)}%</span>
            </div>
          </footer>
        </div>
      </div>

      {showLinkModal && <LinkInsertModal editor={editor} onClose={() => setShowLinkModal(false)} />}
      {showEquationModal && editor && (
        <EquationModal editor={editor} onClose={() => setShowEquationModal(false)} />
      )}
      {eqEditTarget && editor && (
        <EquationModal
          editor={editor}
          editTarget={eqEditTarget}
          onClose={() => setEqEditTarget(null)}
        />
      )}

      {doc && ctxMenu && (
        <EditorContextMenu
          editor={editor}
          menu={ctxMenu}
          onClose={() => setCtxMenu(null)}
          onFontDialog={() => setShowFontDialog(true)}
          onParagraphDialog={() => setShowParaDialog(true)}
          onLink={() => setShowLinkModal(true)}
          onNewComment={startNewComment}
          onAiPreset={(text) => {
            setShowAi(true)
            setAiPreset({ text, nonce: Date.now(), autoRun: true })
          }}
          onRestartNumbering={restartNumbering}
          onContinueNumbering={continueNumbering}
          onUpdateFields={updateFields}
        />
      )}
      {doc && showFontDialog && (
        <FontDialog editor={editor} onClose={() => setShowFontDialog(false)} />
      )}
      {doc && showParaDialog && (
        <ParagraphDialog editor={editor} onClose={() => setShowParaDialog(false)} />
      )}

      {doc && notePrompt && (
        <PromptModal
          title={
            notePrompt.id !== undefined
              ? notePrompt.kind === 'footnote'
                ? t('appEditFootnote')
                : t('appEditEndnote')
              : notePrompt.kind === 'footnote'
                ? t('appInsertFootnote')
                : t('appInsertEndnote')
          }
          placeholder={
            notePrompt.kind === 'footnote'
              ? t('appFootnotePlaceholder')
              : t('appEndnotePlaceholder')
          }
          initial={
            notePrompt.id !== undefined
              ? ((notePrompt.kind === 'footnote' ? footnotes : endnotes).find(
                  (n) => n.id === notePrompt.id,
                )?.text ?? '')
              : ''
          }
          multiline
          onSubmit={submitNote}
          onClose={() => setNotePrompt(null)}
        />
      )}

      {doc && showPagePreview && section && (
        <PaginationPreview
          section={section}
          sections={sections}
          hfParts={doc.parsed.hfParts ?? {}}
          colFlow={viewMode === 'print' ? colFlow : null}
          zoom={zoom}
          hf={{
            header,
            footer,
            ...hfVariants,
            titlePg,
            evenOddHf,
            images: {
              header: doc.parsed.headerImages ?? undefined,
              footer: doc.parsed.footerImages ?? undefined,
              headerFirst: doc.parsed.headerFirst?.images,
              footerFirst: doc.parsed.footerFirst?.images,
              headerEven: doc.parsed.headerEven?.images,
              footerEven: doc.parsed.footerEven?.images,
            },
          }}
          watermark={watermark}
          pageColor={pageColor}
          blockMetaOf={blockMetaOf}
          pageFootnotesOf={pageFootnotesOf}
          endnoteItems={endnoteItems}
          sectionHfOverride={sectionHfOverride}
          clearPageGaps={() => {
            if (editor) setPageGaps(editor.view, [])
          }}
          onExportPdf={() => void exportPdf()}
          onClose={() => setShowPagePreview(false)}
        />
      )}

      {stats && (
        <div
          className="modal-backdrop"
          onMouseDown={(e) => e.target === e.currentTarget && setStats(null)}
        >
          <div className="modal">
            <h2>{t('appWordCountTitle')}</h2>
            <table className="stats-table">
              <tbody>
                <tr>
                  <td>{t('appStatPages')}</td>
                  <td>{stats.pages}</td>
                </tr>
                <tr>
                  <td>{t('appStatWords')}</td>
                  <td>{stats.words}</td>
                </tr>
                <tr>
                  <td>{t('appStatAsianChars')}</td>
                  <td>{stats.asianChars}</td>
                </tr>
                <tr>
                  <td>{t('appStatNonAsianWords')}</td>
                  <td>{stats.nonAsianWords}</td>
                </tr>
                <tr>
                  <td>{t('appStatCharsNoSpaces')}</td>
                  <td>{stats.charsNoSpace}</td>
                </tr>
                <tr>
                  <td>{t('appStatCharsWithSpaces')}</td>
                  <td>{stats.charsWithSpace}</td>
                </tr>
                <tr>
                  <td>{t('appStatParagraphs')}</td>
                  <td>{stats.paragraphs}</td>
                </tr>
                <tr>
                  <td>{t('appStatLines')}</td>
                  <td>{stats.lines}</td>
                </tr>
              </tbody>
            </table>
            <div className="modal-actions">
              <button className="btn-primary" onClick={() => setStats(null)}>
                {t('appClose')}
              </button>
            </div>
          </div>
        </div>
      )}

      {protectModal && (
        <div
          className="modal-backdrop"
          onMouseDown={(e) => e.target === e.currentTarget && setProtectModal(null)}
        >
          <div className="modal">
            <h2>
              {protectModal.mode === 'set' ? t('appRestrictEditing') : t('appUnrestrictEditing')}
            </h2>
            <label className="pgnum-row">
              {protectModal.mode === 'set' ? t('appProtectPwdOptional') : t('appProtectPwdEnter')}
              <input
                type="password"
                autoFocus
                value={protectModal.value}
                onChange={(e) =>
                  setProtectModal({ ...protectModal, value: e.target.value, error: undefined })
                }
                onKeyDown={(e) => {
                  if (e.key === 'Enter') void submitProtectModal()
                }}
              />
            </label>
            {protectModal.error && (
              <p className="pgnum-hint" style={{ color: '#c62828' }}>
                {protectModal.error}
              </p>
            )}
            {protectModal.mode === 'set' && <p className="pgnum-hint">{t('appProtectPwdHint')}</p>}
            <div className="modal-actions">
              <button onClick={() => setProtectModal(null)}>{t('appCancel')}</button>
              <button className="btn-primary" onClick={() => void submitProtectModal()}>
                {t('appOk')}
              </button>
            </div>
          </div>
        </div>
      )}
      {pgNumModal && (
        <div
          className="modal-backdrop"
          onMouseDown={(e) => e.target === e.currentTarget && setPgNumModal(null)}
        >
          <div className="modal">
            <h2>{t('appPageNumFormatTitle')}</h2>
            <label className="pgnum-row">
              {t('appNumberFormat')}
              <select
                value={pgNumModal.fmt}
                onChange={(e) => setPgNumModal({ ...pgNumModal, fmt: e.target.value })}
              >
                <option value="decimal">1, 2, 3, …</option>
                <option value="numberInDash">- 1 -, - 2 -, - 3 -, …</option>
                <option value="lowerLetter">a, b, c, …</option>
                <option value="upperLetter">A, B, C, …</option>
                <option value="lowerRoman">i, ii, iii, …</option>
                <option value="upperRoman">I, II, III, …</option>
                <option value="chineseCounting">一, 二, 三, …</option>
              </select>
            </label>
            <label className="pgnum-row">
              {t('appStartAt')}
              <input
                type="number"
                min={0}
                placeholder={t('appContinueFromPrev')}
                value={pgNumModal.start}
                onChange={(e) => setPgNumModal({ ...pgNumModal, start: e.target.value })}
              />
            </label>
            <p className="pgnum-hint">
              {t('appPgNumHintBlank')}
              {sections.length > 1
                ? t('appPgNumAppliesTo', { n: Math.min(activeSection, sections.length - 1) + 1 })
                : ''}
            </p>
            <div className="modal-actions">
              <button onClick={() => setPgNumModal(null)}>{t('appCancel')}</button>
              <button className="btn-primary" onClick={applyPgNumFormat}>
                {t('appOk')}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
