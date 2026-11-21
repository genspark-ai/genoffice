import { FloatWatch } from './editor/float-watch'
import { attachCommentUndoSync, resetCommentUndoStore } from './editor/comment-undo-sync'
import { scriptFontHtml } from './editor/script-fonts'
import { storyEscLeavesView } from './editor/story-view'
import { carryDocSeen } from './ai/tools'
import {
  DOC_CSS_COMMITTED_EVENT,
  WRAP_WIDTHS_COMMITTED_EVENT,
  setStrictKinsoku,
} from './editor/cjk-punct-shrink'
import { justifyShrinkPluginKey } from './editor/justify-shrink'
import {
  docHyphenationSig,
  hyphenationConfigOf,
  hyphenationDictFor,
  hyphenationReady,
  hyphenationRelayout,
  hyphenationSettingsOf,
  hyphenationZoneApplies,
  loadHyphenation,
  onDocHyphenationChange,
  setHyphenation,
  type HyphenationSettings,
} from './editor/hyphenation'
import { DEFAULT_TAB_STOP_CHANGED_EVENT, tabReflowMeta } from './editor/decoration-extensions'
import { deletedAnything, touchesNode, touchedText } from './editor/touched-blocks'
import {
  createElement,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useReducer,
  useRef,
  useState,
} from 'react'
import type { CSSProperties, MouseEvent as ReactMouseEvent, SetStateAction } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import type { Editor } from '@tiptap/core'
import { handleDocsControl, type ControlRequest } from './control'
import { DOMParser as PmDOMParser, type Mark as PmMark, Slice as PmSlice } from '@tiptap/pm/model'
import { NodeSelection, TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import {
  Dropdown,
  ImageViewer,
  aiPanelInitiallyOpen,
  rememberAiPanelOpen,
  createZoomWheelClassifier,
  useAutoSavePref,
  useAiSidebarEnabled,
} from '@genoffice/ui'
import { commentAnchors, commentIdsAt } from './editor/comments'
import { isCommandLetter } from './command-letter'
import {
  PAGE_FIELD_INSTR,
  setFieldInstr,
  toggleAllFieldCodes,
  type FieldRange,
} from './editor/field-codes'
import { refreshTableFormulas } from './editor/table-formula'
import { linkTarget } from './editor/link-actions'
import { TOP_ANCHOR, anchorPos } from './editor/link-targets'
import { FieldDialog } from './components/FieldDialog'
import { createLabelSheet } from './mail-merge/label-sheet'
import { layoutSection } from './mail-merge/layout'
import { formatFieldNumber, formatFieldText } from './editor/field-switches'
import { markdownPasteHtml } from './editor/markdown-paste'
import { pasteTextSlice, singleCellPasteText } from './editor/paste-text'
import {
  imageFilesFromDataTransfer,
  insertImageFilesAtCoords,
  isImageFileDrag,
} from './editor/image-drop'
import {
  caretFontSlots,
  fillMissingRunFonts,
  isForeignPasteHtml,
  pushDownWebInlineStyles,
} from './editor/paste-web-html'
import {
  beginForeignPaste,
  caretParaFormat,
  caretStyleMark,
  consumeForeignPaste,
  defaultPasteMode,
  lastForeignPasteMode,
  mergeFormattingFragment,
  stashPastePayload,
  takeTextReroute,
} from './editor/paste-options'
import { PasteOptionsChip } from './components/PasteOptionsChip'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  DEFAULT_SECTION,
  applyPageNumType,
  customLevelFromNumberingLevel,
  applySectionSettings,
  applySectionStartType,
  applyTitlePg,
  readSections,
  verifyProtectionPassword,
  type Block,
  type CommentInfo,
  type CustomNumberingLevel,
  type NumberingDef,
  type PageNumberChapter,
  type DocProtection,
  type PropertyChanges,
  type WriteProtection,
  mergeFieldName,
  nextNoteId,
  placeholderText,
  type HeaderFooter,
  type HfPictureEdit,
  type HfImage,
  type NoteInfo,
  type NoteProps,
  type NoteSeparator,
  type TextboxDisplay,
  type ParsedDocFull,
  type SectionInfo,
  type SectionSettings,
  type CitationStyleId,
  type SourceInfo,
  pendingHeadingLevel,
  pictureWatermarkPreviewImage,
  textWatermarkPreviewImage,
  previewFontSettings,
  type StyleInfo,
  type StyleUpsert,
  type DefaultFonts,
  type ThemeColors,
  type ThemeFonts,
  type PictureWatermarkSpec,
  type Watermark,
  type WatermarkSpec,
  type CommentImage,
  type MathPr,
  DEFAULT_MATH_PR,
} from '@genoffice/docx-engine'
import type {
  AiDocContent,
  AiSettings,
  MergeDataSource,
  MergeOutputMode,
  OpenDocxResult,
} from '../shared/ipc'
import { AI_PROVIDERS } from '../shared/ipc'
import {
  buildMailMergeFiles,
  type MailMergeRequest,
  type MailMergeResponse,
} from './mail-merge/output'
import { ZoteroDocumentController } from './zotero/controller'
import { AiPanel, aiRevisionAuthor } from './ai/AiPanel'
import type { AiCommentsAccess, AiDocExtras, AiHeaderFooterAccess } from './ai/tools'
import type { AiStyleInfo } from './ai/style-ops'
import {
  applyResolvedPageSetup,
  describeSection,
  mergePendingPgNum,
  type AiPageSetupAccess,
} from './ai/page-setup'
import {
  dropRemovedPendingSections,
  patchPendingSectPr,
  pendingSectionAt,
  rewritePendingSectPrs,
  sectionIndexAtBlock,
} from './ai/pending-sections'
import {
  noteBlockSections,
  protectedNoteMarkBlock,
  renumber,
  type AiNotesAccess,
} from './ai/note-ops'
import { applyHfText, hfEditText } from './editor/hf-text'
import {
  insertHfField,
  liveHfFieldOf,
  insertHfInstrField,
  insertHfPicture,
  mountHfEditor,
  type HfEditorHandle,
} from './editor/hf-editor'
import type { HfAction, HfEditingInfo } from './components/ribbon-hf-tab'
import {
  hfCommitTarget,
  hfLinked,
  hfTextCommit,
  ownHf,
  resolveHf,
  withHfLink,
  withHfPictureEdit,
  type HfSectionState,
} from './hf-sections'
import { installHfFloatEdit, tagHfFloatEl } from './editor/hf-float-edit'
import { hfFooterLine, hfLayoutResolved, hfPhantomSpec, hfWithPhantom } from './hf-phantom'
import { hfVariantHeightsPx, sectionHfHeights } from './hf-heights'
import { makeGapNotesEl, measureNoteHeightDom, noteRowMetrics } from './editor/note-dom'
import { AiAskPopover } from './components/AiAskPopover'
import { EDIT_QUEUE_MAX, selectionForAnchor, type DocsEditQueueItem } from './ai/edit-queue'
import { addQueueAnchor, clearQueueAnchors, removeQueueAnchors } from './editor/ai-queue-anchors'
import { countWords, documentTextForWordCount, selectedWordCount, textStats } from './word-count'
import { CommentsPanel } from './components/CommentsPanel'
import { EquationModal } from './components/EquationModal'
import { HeaderFooterArea } from './components/HeaderFooterArea'
import { PageFootnotes, PageEndnotes } from './components/PageNoteAreas'
import {
  endnoteHomesOf,
  groupSectionEndnotes,
  setSectionEndnotes,
  type EndnoteHomes,
  type SectionEndnoteGroup,
} from './editor/section-endnotes'
import { sectionNotePropsOf, setNoteProps, type NoteKind } from './note-format'
import { PaginationPreview } from './components/PaginationPreview'
import { imageParaPaginationMeta, paraPaginationMeta, tableWidowOff } from './editor/para-flags'
import {
  comparePlanToBlocks,
  planByElement,
  planDiffHistogram,
  planPropsOf,
  type PlanOf,
  type PlanDiff,
} from './pagination-plan'
import { PrintDialog } from './components/PrintDialog'
import {
  hfVariantOf,
  hfVariantReservedPx,
  bumpLineSampleFontEpoch,
  carryStreamedSamples,
  endnotesAnchorY,
  createLineRectsCache,
  anchorElement,
  lineStartAnchor,
  liveSections,
  nextLineAnchor,
  measureBlocks,
  alignJoinedParagraphs,
  docGridPitchPt,
  type LineAnchor,
  hfPartWidthPx,
  headingPagesOf,
  pageCutSignature,
  pageLeadsOf,
  applyGapLedger,
  pagePaperEdges,
  pageReaderShadow,
  type FlowCarry,
  LEAD_ROW_TOL_PX,
  outerTableRows,
  type PageCut,
  pageNumbers,
  type SliceOutputs,
  tableRowFlags,
  type BlockBox,
  type BlockMeta,
  type PageNoteItem,
  pageAt,
  pageOfY,
  singleCutCell,
  columnLayoutSpecs,
  widthPassGate,
  newWidthPassState,
  resetWidthPassHistory,
  vAlignShiftSpecs,
  verticalTextSpecs,
  blockInlineExtraPx,
  sectionVertical,
  sectionPageStartSpecs,
  wrapWidthSig,
  sectionTopMarginSpecs,
  sectionWidthSpecs,
  mirrorIndentSpecs,
  mirrorMarginSpecs,
  paperWidthPx,
  pageLeftPx,
  sectionGridPitchPt,
  sectionGridPitchSpecs,
  docCharSpacePt,
  docCharCellPt,
  sectionCharCellSpecs,
  sectionCharSpaceSpecs,
  singleByteCharSpacePt,
  type ColumnBlockPlacement,
  sectionBidi,
  sectionColGeom,
  sectionColumns,
  sectionGeoms,
  sectionPageBox,
  canvasContentTopPx,
  effectiveTopPx,
  effectiveBottomPx,
  singleVariantGeomOf,
  formatPageNumber,
  sectionPageCounts,
  visiblePageCount,
  type SectionGeom,
  type SectionHfHeights,
  isResumePage,
  type PageSlice,
  type PassResume,
  blockGapRects,
} from './pagination'
import { prepareFlow, sliceFlow, type PassContext } from './pagination-pass'
import {
  GAP_BAND,
  alignGapHfStrips,
  alignTableGapFills,
  clearFloatShifts,
  floatFlowChangedMeta,
  insideFloatTable,
  setFloatVShifts,
  floatCarriesOf,
  floatShiftsOf,
  setOversizeClips,
  setPageGaps,
  stampGapFlowTops,
  canvasZoomOf,
  rowFillsSig,
  setRowFills,
  syncAnchorBands,
  syncCutOverlays,
  syncNestedTableGaps,
  nestedRowAtCut,
  nestedRowSplit,
  type NestedTableGap,
  NESTED_ROW_TOL_PX,
  tableGridCols,
  syncFloatShifts,
  syncPageBorders,
  syncPageSheets,
  clampCellBoxTops,
  clampCellImageTops,
  pinCellBoxesToPage,
  pinPageRelXPictures,
  pinParagraphRelVPictures,
  pageBorderStyleOf,
  type PageGapSpec,
  type GapMetrics,
  LayoutBatch,
  pageFramesFromGaps,
  gapLedgerOf,
  PAGE_TURN_SEL,
  isPageTurn,
  type PageFrame,
} from './editor/pagination-gaps'
import { syncWrapObstacles } from './editor/wrap-obstacles'
import { newObstacleShadow, obstaclePlanDiffs } from './editor/obstacle-plan'
import { MultipageGrid, snapshotCanvas, type CanvasSnapshot } from './components/MultipageGrid'
import { layoutMultipage, multipageColumns, multipageZoomForColumns } from './multipage'
import { setColumnLayout } from './editor/column-layout'
import { pageMargins } from './page-margins'
import { syncLineNumbers } from './editor/line-numbers'
import { syncColumnRules } from './editor/column-rules'
import { formattingMarksPluginKey, sectionBreakMap } from './editor/formatting-marks'
import { draftAutoBreaks, syncDraftBreaks } from './editor/draft-view'
import {
  COMMENT_AREA_W,
  clearMarginAnnotations,
  syncMarginAnnotations,
  type MarginHandlers,
} from './editor/margin-annotations'
import {
  bumpHfProbeFontEpoch,
  hfHasPageField,
  hfHasVisibleContent,
  hfPageNumberOnly,
  hfWithoutPageMarks,
  hfStripGeom,
  hfBodyObstacles,
  makeGapHfEl,
  hfImageHangsOnPara,
  makeHfFloatImgEl,
  type HfFloatBox,
} from './editor/hf-dom'
import {
  estimateFootnoteHeight,
  resolveNoteStyle,
  noteRunStyle,
  footnoteLineHeightPx,
  noteSeparatorPx,
  noteSeparatorRow,
  setNoteSeparatorRows,
  takeSettledMetricFaces,
  textHasCjk,
} from './line-metrics'
import { saveUntilPersisted } from './save-until-persisted'
import { SPELLCHECK_KEY, spellcheckEnabled } from './spellcheck-pref'
import { cachedByDoc } from './doc-cache'
import { useShallowStable, useStableCallbacks } from './use-stable'
import { FindPanel } from './components/FindPanel'
import { ZoomDialog } from './components/ZoomDialog'
import { TrackChangesOptionsDialog } from './components/TrackChangesOptionsDialog'
import { displayForPrint, HighlightChangesDialog } from './components/HighlightChangesDialog'
import { DocumentDialog, gutterMoved } from './components/DocumentDialog'
import {
  pageSetupChange,
  patchSections,
  recordPageSetupChange,
  replayedPageSetup,
  type PageSetupPatch,
  type PageSetupState,
} from './editor/page-setup-history'
import { ChangeCaseDialog } from './components/ChangeCaseDialog'
import { DateTimeDialog } from './components/DateTimeDialog'
import { ColumnsDialog } from './components/ColumnsDialog'
import { SymbolDialog } from './components/SymbolDialog'
import { FootnoteDialog, type NoteSettings } from './components/FootnoteDialog'
import { fontSymbolPayload } from './editor/font-symbol'
import { IconsDialog } from './components/IconsDialog'
import { prepareLibraryIcons } from './editor/insert-icons'
import { PageNumbersDialog } from './components/PageNumbersDialog'
import { TabsDialog } from './components/TabsDialog'
import { BulletsNumberingDialog } from './components/BulletsNumberingDialog'
import { NO_LIST_STYLE_ID, listStyleEntries } from './list-styles'
import { TranslationLanguageDialog } from './components/TranslationLanguageDialog'
import { IndexTablesDialog } from './components/IndexTablesDialog'
import {
  captionLabels,
  INDEX_TABLES_COMMAND,
  indexBlockRanges,
  indexSectionEdges,
  indexTablesAtSelection,
  markCitation,
  tableSpecPayload,
} from './commands/reference-commands'
import { CITATION_CATEGORY_LABELS, MarkCitationDialog } from './components/MarkCitationDialog'
import { citationRunsOf } from './editor/citations'
import { MarkIndexEntryDialog } from './components/MarkIndexEntryDialog'
import { collectBookmarks } from './editor/cross-reference'
import {
  indexEntrySwitches,
  indexEntryTerm,
  insertFieldMarks,
  markIndexEntries,
  occurrenceEnds,
} from './editor/index-entry'
import { HyphenationDialog, ManualHyphenationDialog } from './components/HyphenationDialog'
import { refreshLimitDefaults } from './editor/equation'
import { mathDefaultsCss, setMathLimits } from './editor/math-defaults'
import { MATH_WRAP_CHANGED_EVENT, setMathWrapRules } from './editor/math-wrap'
import { RevealFormattingDialog, revealedStyleName } from './components/RevealFormattingDialog'
import { AutoTextDialog } from './components/AutoTextDialog'
import { LanguageDialog } from './components/LanguageDialog'
import { setMergePreview } from './editor/merge-preview'
import { DEFAULT_CITATION_STYLE, nextFieldId } from './editor/source-citations'
import type { MergePrompt } from './mail-merge/template'
import { SelectionPane } from './components/SelectionPane'
import { CitationsPane } from './components/CitationsPane'
import { selectChartForData } from './editor/page-objects'
import {
  languageSlot,
  noProofOf,
  parseLangAttr,
  proofingLanguageOf,
  setProofingLanguage,
} from './editor/proofing-language'
import {
  autoTextEntries,
  deleteAutoText,
  insertAutoText,
  saveAutoText,
  selectionClip,
  selectionSnippet,
  type AutoTextEntry,
} from './editor/auto-text'
import { PhoneticGuideDialog } from './components/PhoneticGuideDialog'
import { EncloseCharactersDialog } from './components/EncloseCharactersDialog'
import { CombineCharactersDialog } from './components/CombineCharactersDialog'
import {
  applyEqLayout,
  eqLayoutFontOf,
  eqLayoutTargetOf,
  type EqLayoutKind,
  type EqLayoutTarget,
} from './editor/eq-layout'
import { FitTextDialog } from './components/FitTextDialog'
import { TextEffectsDialog } from './components/TextEffectsDialog'
import { pickAndInsertFile } from './editor/insert-file'
import { pickAndAutoMark } from './editor/index-automark'
import { applyTextFx, textFillOf, textFxOf } from './editor/text-fx'
import { fitTextOf, setFitText, textStretch, toggleEaLayout } from './editor/asian-layout'
import { applyRuby, rubyTargetOf, type RubyTarget } from './editor/phonetic-guide'
import { PropertiesDialog } from './components/PropertiesDialog'
import type { ChartType } from './components/chart-types'
import {
  BordersShadingDialog,
  FontSubstitutionDialog,
  sectionWithPageBorder,
} from './components/BordersShadingDialog'
import { RULE_FIELD_INSTR, ruleNeedsList, type MergeRule } from './mail-merge/rules'
import { checkMissingFonts, collectDocFonts, docFontsCoveredBy } from './font-check'
import { fontSubstituteOf, setFontSubstitute } from './font-substitutes'
import { convertFontsInEditor } from './editor/convert-fonts'
import {
  fontsBeingRestored,
  replayedFontConversions,
  withFontsConverted,
  withFontsReverted,
  type FontConversions,
} from './editor/font-conversion-history'
import {
  AltTextDialog,
  EquationOptionsDialog,
  LineNumbersDialog,
  NumberDialog,
  ParagraphSpacingDialog,
  ReduceFileSizeDialog,
  InsertFileRangeDialog,
  TemplateGalleryDialog,
  RenameDialog,
  SizePositionDialog,
  SortDialog,
  TextDirectionDialog,
  moveErrorKey,
} from './components/MenuDialogs'
import {
  applyPictureGeometry,
  selectedPictureAlt,
  selectedPictureGeometry,
  setSelectedPictureAlt,
} from './editor/picture-props'
import { saveDesignDefault, saveFontDefault, savePageDefault } from './design-default'
import { charCellSpacingCss, gridBaseSizePt } from './char-grid'
import { sortScope, sortSelection, type SortRefusal } from './editor/sort-paragraphs'
import {
  FilterRecipientsDialog,
  FindRecipientDialog,
  RecipientsDialog,
  RuleDialog,
} from './components/RecipientDialogs'
import {
  EMPTY_RECIPIENT_QUERY,
  activeRecords,
  findRecipient,
  type RecipientQuery,
} from './mail-merge/recipients'
import { ribbonLangKey } from './components/translate-langs'
import { setTranslationLanguage, translationLanguage } from './translate-lang'
import { redactTextBetween, type RedactableNode } from './ai/redact-view'
import { parseTabStops, tabStopsAttr } from './components/tab-stops'
import { MailMergePanel } from './components/MailMergePanel'
import { MERGE_RANGE_ALL, type MergeRange } from './mail-merge/range'
import { DropCapDialog, dropCapValueOf } from './components/DropCapDialog'
import { dropCapFontOf } from './editor/drop-cap'
import {
  datePictureOf,
  formatDatePicture,
  toFullWidth,
  wantsFullWidth,
} from './editor/date-picture'
import { PasteSpecialDialog } from './components/PasteSpecialDialog'
import {
  loadTrackChangesOptions,
  movesByAuthor,
  saveTrackChangesOptions,
  trackChangesOptionsCss,
  trackChangesOptionsState,
  type TrackChangesOptions,
} from './editor/track-changes-options'
import type { PageLayout } from './editor/goto'
import { Ribbon } from './components/Ribbon'
import { BookmarkModal, CaptionModal, CrossRefModal } from './components/ReferenceDialogs'
import { fieldLinkBookmark, refreshNoteRefFields } from './editor/note-ref-fields'
import { xmlOfNode } from './editor/toc-update'
import { setAutoCaptionBlocks } from './editor/auto-caption'
import { watermarkPresetsOf } from './components/ribbon-design-tab'
import { WatermarkDialog, watermarkPageBoxOf } from './components/WatermarkDialog'
import { CropDialog, CutoutDialog } from './components/PictureDialogs'
import { replaceSelectedPictureBytes } from './editor/picture-bytes'
import {
  AdjustListIndentsDialog,
  DefineBulletDialog,
  DefineMultilevelDialog,
  DefineNumberFormatDialog,
  SetNumberingValueDialog,
  type ListDialogKind,
} from './components/ListDialogs'
import {
  documentListPresets,
  numberPresetLevels,
  customizeLibraryCard,
  listLibrary,
  recentListPresets,
  rememberListPreset,
} from './list-presets'
import { bordersShadingTarget, computeFormatState } from './components/ribbon-format-state'
import {
  IconDraftView,
  IconOutlineView,
  IconStoryView,
  IconPrintLayout,
  IconReadMode,
  IconRedo,
  IconSave,
  IconUndo,
  IconWebLayout,
} from './components/icons'
import { ToastHost } from './components/toast'
import {
  imageSizeOf,
  AI_REWRITE_ACK_KEY,
  ChartInsertModal,
  LinkInsertModal,
  TableInsertModal,
  applyParagraphStyle,
  clearParagraphFormatting,
  applyGalleryStyle,
  clearAllFormatting,
  insertBlankPageAt,
  insertImageFromDataUrl,
  insertImageViaDialog,
  insertPageBreakAt,
  setParaAttrs,
  activeParaAttrs,
  insertTableAt,
  type RevisionDisplayMode,
  type ViewMode,
} from './components/ribbon-tabs'
import { CompareDocumentsDialog } from './components/CompareDocumentsDialog'
import {
  EditorContextMenu,
  FontDialog,
  ParagraphDialog,
  type ContextMenuState,
} from './components/ContextMenu'
import { RedactDialog } from './editor/RedactDialog'
import {
  CellMarginsDialog,
  DeleteCellsDialog,
  InsertCellsDialog,
  SplitCellsDialog,
  type TableDialogKind,
} from './components/TableDialogs'
import {
  ConvertTableToTextDialog,
  ConvertTextToTableDialog,
  FormulaDialog,
} from './components/TableConvertDialogs'
import { TablePropertiesDialog } from './components/TablePropertiesDialog'
import { showToast } from './components/toast-bus'
import { editableOrToast, ensureCaretInTable, keepCellSelection } from './editor/table-ops'
import { PromptModal } from './components/PromptModal'
import { ShortcutsDialog } from './components/ShortcutsDialog'
import { CustomizeKeyboardDialog } from './components/CustomizeKeyboardDialog'
import { setKeyLayerRunner } from './key-layer'
import { SpellingDialog } from './components/SpellingDialog'
import { WordCountDialog, type DocStats } from './components/WordCountDialog'
import { AutoCorrectDialog } from './components/AutoCorrectDialog'
import { PreferencesDialog } from './components/PreferencesDialog'
import { applyCase, nextCaseMode, selectionText } from './editor/case-transform'
import { stepHangingIndent, stepParagraphIndent } from './editor/indent'
import { PasswordDialog } from './components/PasswordDialog'
import { ProtectDialog, type ProtectDialogResult } from './components/ProtectDialog'
import { t, useI18n, type StringKey } from './i18n/locale'
import {
  getActiveSubEditor,
  setActiveSubEditor,
  subscribeSubEditorState,
} from './editor/active-editor'
import { installSelectionBar } from './editor/selection-bar'
import { collectHeadings } from './editor/headings'
import { applyTocPageDisplays, backfillTocPageDom } from './editor/toc-refresh'
import { renumberSeqFields, unreliableSeqLabelsOf } from './editor/field-results'
import { setSelectionAlign } from './editor/direction'
import {
  firstListNumId,
  markupShown,
  runMenuCommand,
  canRepeatCommand,
  type CommandHost,
  type PictureDialogKind,
} from './commands/menu-commands'
import { FALLBACK_STYLES, defaultParagraphStyleId } from './style-gallery'
import { NONE, type ActionCommandId } from '../shared/commands'
import { menuCommandState } from './commands/menu-state'
import { textDirectionContext } from './commands/text-direction'
import { useEditTarget } from './commands/edit-target'
import { onEditableChange, setEditorEditable } from './editor/editable'

import {
  editorExtensions,
  findFloatImageAt,
  runPictureAtEvent,
  resolvedCommentsPluginKey,
  revisionDisplayState,
} from './editor/extensions'
import { resolveDarkPage, writeDarkPagePref } from './dark-page-pref'
import { useDocThemeIsDark, useUiThemeIsDark } from './ui-theme'
import { type InkAnnotation, type InkTool } from './editor/ink'
import { loadInkPresets, saveInkPresets, type InkPreset } from './editor/ink-presets'
import { InkOverlay } from './components/InkOverlay'
import {
  TRACK_IGNORE,
  collectRevisions,
  selectRevision,
  type RevisionRange,
  type TrackChangesStorage,
} from './editor/revisions'
import {
  DEFAULT_REVISION_FILTER,
  authorInkIndex,
  buildReviewEntries,
  filterRevisions,
  revisionAuthorCss,
  revisionAuthorsInkOrder,
  revisionsInSelection,
  storyAuthorshipCss,
  type ReviewEntry,
  type RevisionFilter,
  type RevisionScope,
} from './editor/revision-filters'
import { ReviewingPane } from './components/ReviewingPane'
import {
  collectHfRevisions,
  collectNoteRevisions,
  collectTextboxRevisions,
  liveStoryRevision,
  ownHfSlots,
  resolveHfValue,
  resolveNotes,
  resolveTextboxes,
  type StoryRevision,
} from './editor/story-revisions'
import {
  BALLOON_MODE_KEY,
  balloonKindOf,
  balloonRevKinds,
  loadBalloonMode,
  markupPrintsIn,
  noteRunsInView,
  revisionHiddenIn,
  type BalloonMode,
} from './editor/revision-view'
import { setInlineComments } from './editor/balloon-markers'
import { NavPane, type NavTab } from './components/NavPane'
import { loadLeftPane, saveLeftPane, type LeftPane } from './components/pane-width'
import {
  commentsView,
  iconClickTarget,
  loadCommentsMode,
  saveCommentsMode,
  type CommentsMode,
} from './components/comments-mode'
import { StylesPane } from './components/StylesPane'
import { useStoredFlag } from './stored-flag'
import { Ruler, rulerDims } from './components/Ruler'
import { RulerBar, rulerBarHeight } from './components/RulerBar'
import { VerticalRuler } from './components/VerticalRuler'
import { docBodyFont, docHasCjk, docLineFactor, docStyleCss, docThemeCss } from './doc-style-css'
import { fontAliasAnswers, onFontAliasesChanged } from './system-fonts'
import { isDocDirty, runGuardedCandidate, runGuardedDocumentAction } from './doc-dirty'
import {
  EMPTY_HF_VARIANTS,
  restingHfAreaVariant,
  type DocState,
  type HfVariantKey,
  type HfVariantsState,
  type HfView,
  EMPTY_PENDING_NUMBERING,
  pendingNumberingDirty,
  type PendingNumbering,
} from './doc-state'
import {
  applyAiDocContent as applyAiDocContentImpl,
  buildDocBytes as buildDocBytesImpl,
  exportPdf as exportPdfImpl,
  exportHtml as exportHtmlImpl,
  exportMarkdown as exportMarkdownImpl,
  exportImages as exportImagesImpl,
  loadFile as loadFileImpl,
  saveAsTemplate as saveAsTemplateImpl,
  newFile as newFileImpl,
  printDoc as printDocImpl,
  save as saveImpl,
  suggestedSaveAsName,
  writeRecoveryCopy as writeRecoveryCopyImpl,
  type FileActionContext,
  type PendingPdfExport,
} from './file-actions'
import { reduceFileSize } from './reduce-file-size'
import { selectedPictureSources } from './reduce-pictures'
import { BlockIndex } from './pagination-index'
import { isPhasedContentPending } from './phased-content'
import { sameSectionInfos } from './pagination-sections'
import {
  blocksKeepSlicing,
  firstChangedTopLevelIndex,
  singleInlineEditIndex,
  type FastPassBlock,
} from './editor/pagination-fast-path'
import { FollowUpChain } from './pagination-loop'
import { tabLayoutDiag } from './editor/decoration-extensions'
import { layoutHealth } from './editor/layout-health'

/** min gap between whole-document pagination passes while a phased open streams its tail */
const STREAMING_PASS_GAP_MS = 1500
/** passes wait at least this many times their own duration while streaming */
const STREAMING_PASS_DUTY = 4
const STREAMING_PASS_GAP_MAX_MS = 15_000
/** after an edit, the pass waits this many times its own last duration (300 ms floor) */
const EDIT_PASS_DUTY = 3
const EDIT_PASS_DEBOUNCE_MAX_MS = 2000
/** a pass re-running for its own late-applied widths / suppression: nobody is
 *  typing, so it only waits for the previous pass's DOM writes to commit */
const FOLLOW_UP_PASS_DELAY_MS = 100
const EXPORT_PASS_RETRY_MS = 500
const WORD_COUNT_THROTTLE_MS = 400
/** the refresh makes two whole-document walks (text extraction + revision
 *  scan); past this duration the document is big enough that walking it
 *  every 400 ms stole real typing time (issue #2781), so the cadence
 *  stretches the way the pagination scheduler's does */
const WORD_COUNT_SLOW_WALK_MS = 25
const WORD_COUNT_SLOW_INTERVAL_MS = 2000
/** consecutive follow-up passes a pass may schedule for itself */
import { runHeadlessDocumentExport } from './headless-export'
import { installMcpBridge } from './mcp-bridge'
import {
  clampDocsZoom,
  DOCS_ZOOM_MAX,
  DOCS_ZOOM_MIN,
  nextDocsZoom,
  stepDocsZoom,
  type ZoomFitMode,
} from './wheel-zoom'
import {
  allocateListNumId as allocateListNumIdImpl,
  addPictureBullet,
  changeListLevel as changeListLevelImpl,
  continueNumbering as continueNumberingImpl,
  currentListTarget as currentListTargetImpl,
  currentListValue,
  editListLevel as editListLevelImpl,
  hasPreviousList,
  setNumberingValue as setNumberingValueImpl,
  createCustomListDef as createCustomListDefImpl,
  listStyleNum as listStyleNumImpl,
  addListStyle as addListStyleImpl,
  deleteListStyle as deleteListStyleImpl,
  listStylesInUse,
  listNumberingState,
  restartNumbering as restartNumberingImpl,
  type NumberingContext,
} from './numbering-actions'
import {
  addInk as addInkImpl,
  addCommentAt as addCommentAtImpl,
  cancelNewComment as cancelNewCommentImpl,
  clearInks as clearInksImpl,
  commentThreadAtCaret as commentThreadAtCaretImpl,
  deleteAllComments as deleteAllCommentsImpl,
  pickCompareFile,
  runCompare,
  deleteComment as deleteCommentImpl,
  deleteCommentAtCaret as deleteCommentAtCaretImpl,
  convertNotes as convertNotesImpl,
  deleteNote as deleteNoteImpl,
  editComment as editCommentImpl,
  gotoComment as gotoCommentImpl,
  handleRevision as handleRevisionImpl,
  handleRevisionRange as handleRevisionRangeImpl,
  removeInks as removeInksImpl,
  replyToComment as replyToCommentImpl,
  resolveComment as resolveCommentImpl,
  startNewComment as startNewCommentImpl,
  submitNewComment as submitNewCommentImpl,
  submitNote as submitNoteImpl,
  type NotePrompt,
  type ReviewContext,
} from './review-actions'
import { authorName, onAuthorNameChanged } from './author-name'

const IS_MAC = navigator.platform.toLowerCase().includes('mac')

const twipsToPx = (twips: number) => (twips / 1440) * 96

/** tiny stable string hash for decoration keys */
function hashStr(str: string): number {
  let h = 0
  for (let i = 0; i < str.length; i++) h = (h * 31 + str.charCodeAt(i)) | 0
  return h
}

const EMPTY_BLOCKS: Block[] = []
const NO_SECT_ENDNOTES: SectionEndnoteGroup[] = []

const SORT_REFUSALS: Record<SortRefusal, StringKey> = {
  needsParagraphs: 'ribbonSortNeedsParagraphs',
  mergedCells: 'ribbonSortMergedCells',
}

const sortDialogScope = (state: EditorState) => {
  const scope = sortScope(state)
  return typeof scope === 'string' ? null : scope
}

// O(doc) derivations cached by PM doc reference: caret moves and unrelated
// state updates reuse the last result instead of re-walking the whole document
const wordCountOfDoc = cachedByDoc((d) => countWords(documentTextForWordCount(d)))
const revisionsOfDoc = cachedByDoc((d) => collectRevisions(d))

/**
 * Document position of a measured line-start DOM anchor. posAtDOM works from the DOM
 * tree, unlike posAtCoords whose viewport hit-testing returns degenerate positions for
 * lines scrolled off-screen (which misplaced page-break markers into table bodies,
 * inflating the canvas table by a phantom anonymous row and skewing pagination).
 */
function posFromAnchor(view: Editor['view'], anchor: LineAnchor): number | undefined {
  try {
    // element anchors (picture-only lines): address the position before the
    // element via its parent + child index (offset semantics inside an atom
    // leaf's own DOM are undefined)
    const pos =
      anchor.node instanceof Element
        ? anchor.node.parentNode
          ? view.posAtDOM(
              anchor.node.parentNode,
              Array.prototype.indexOf.call(anchor.node.parentNode.childNodes, anchor.node),
            )
          : -1
        : view.posAtDOM(anchor.node, anchor.charOffset)
    return pos >= 0 ? pos : undefined
  } catch {
    return undefined
  }
}

/** Clean pasted Word/web HTML: mso conditional comments, <o:p>, and unwrapping <li><p>x</p></li>;
 *  foreign fragments also get block-level font/size/color pushed down onto
 *  spans so whole-paragraph web copies keep their formatting (r181) */
function cleanPastedHtml(html: string): string {
  const cleaned = unwrapSingleCellTable(
    html
      .replace(/<!--\[if[\s\S]*?<!\[endif\]-->/g, '')
      .replace(/<o:p>[\s\S]*?<\/o:p>/g, '')
      .replace(/<li([^>]*)>\s*<p[^>]*>([\s\S]*?)<\/p>\s*<\/li>/g, '<li$1>$2</li>'),
  )
  return isForeignPasteHtml(html) ? pushDownWebInlineStyles(cleaned) : cleaned
}

/** Plain text for a text-mode paste whose clipboard carries no text/plain
 *  flavor (some apps write HTML only) — the fragment's text content with
 *  block boundaries as line breaks. */
function htmlFallbackText(html: string): string {
  if (!html) return ''
  try {
    const body = new window.DOMParser().parseFromString(html, 'text/html').body
    for (const block of body.querySelectorAll('p,div,li,tr,h1,h2,h3,h4,h5,h6,br')) {
      block.after(body.ownerDocument.createTextNode('\n'))
    }
    return (body.textContent ?? '').replace(/\n{2,}/g, '\n').trim()
  } catch {
    return ''
  }
}

// The per-paste foreign-HTML handshake (mode selection, chip payload) lives
// in editor/paste-options.ts: the paste DOM handler and the ribbon Paste
// button arm it (the clipboard is parsed BEFORE handlePaste runs, so the
// mode must be decided when the paste event arrives), transformPasted and
// the handlePaste lanes consume it. Our own clipboard HTML keeps its exact
// run marks, plain-text pastes go through pasteTextSlice, and drops never
// arm anything.

/**
 * Word parity: pasting a copy of a SINGLE spreadsheet cell inserts its
 * content as text, not a 1x1 table (a lone cell copied from Sheets/Excel
 * pasted into Docs as a table and flipped the ribbon into table tools). Only a payload whose entire content is one table with one
 * cell unwraps; anything else — multi-cell tables, tables mixed with prose —
 * stays intact.
 */
function unwrapSingleCellTable(html: string): string {
  if (!/<table/i.test(html)) return html
  try {
    const doc = new window.DOMParser().parseFromString(html, 'text/html')
    const body = doc.body
    const tables = body.querySelectorAll('table')
    if (tables.length !== 1) return html
    const table = tables[0]!
    // The table must be the only real content of the paste. Text equality
    // covers prose ANYWHERE outside the table — including inside wrappers
    // that also contain it — and Element.textContent excludes HTML comments,
    // so Excel's StartFragment markers don't block the unwrap.
    if ((body.textContent ?? '').trim() !== (table.textContent ?? '').trim()) return html
    // text-less content outside the table (images, separators) also keeps it
    if (
      [...body.querySelectorAll('img,svg,video,hr')].some((element) => !table.contains(element))
    ) {
      return html
    }
    const cells = table.querySelectorAll('td,th')
    if (cells.length !== 1) return html
    if (cells[0]!.querySelector('table')) return html
    return cells[0]!.innerHTML
  } catch {
    return html
  }
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

const DEFAULT_SETTINGS: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((p) => [
      p.id,
      { apiKey: '', model: p.defaultModel, baseUrl: p.needsBaseUrl ? '' : undefined },
    ]),
  ) as AiSettings['providers'],
}

/** heading level after a define_style patch: null outlineLevel drops heading status, a number sets it, absent keeps the current */
/** An encrypted file is not a document yet: the replace guard waits until the
 *  password prompt has decrypted one (submitDocPwd), so a cancelled or wrong
 *  password never saves or discards the document that is still open. */
function replacesDocument(result: NonNullable<OpenDocxResult>): boolean {
  return !('needsPassword' in result)
}

const warnedPlanClasses = new Set<string>()

/** I9 shadow: plan vs scrape for the measured blocks; disagreements are reported, never applied */
function shadowPlanDiff(
  planOf: PlanOf,
  blocks: BlockBox[],
  metaOf: (docxIndex: number) => BlockMeta | undefined,
): PlanDiff[] {
  const diffs = comparePlanToBlocks(blocks, planOf, metaOf)
  if (import.meta.env.DEV) {
    for (const d of diffs) {
      if (warnedPlanClasses.has(d.cls)) continue
      warnedPlanClasses.add(d.cls)
      console.warn(`[pagination] plan/scrape disagree on ${d.cls}`, d)
    }
  }
  return diffs
}

/** the language a new phonetic guide is tagged with, by UI language, when the document names none */
const RUBY_LANGS: Record<string, string> = {
  zh: 'zh-CN',
  'zh-TW': 'zh-TW',
  ja: 'ja-JP',
  ko: 'ko-KR',
}

export function App() {
  // subscribe to language switches for re-render; strings all go through module-level t, so memoized callbacks never capture stale closures
  const { lang } = useI18n()
  const [doc, setDoc] = useState<DocState | null>(null)
  /** a phased open is still streaming the document tail: editor stays read-only */
  const [docLoading, setDocLoading] = useState(false)
  /** true until the pending-open / new-blank boot checks settle; the start screen stays hidden meanwhile */
  const bootPendingRef = useRef<Promise<[OpenDocxResult, boolean, AiDocContent | null]> | null>(
    null,
  )
  const bootHandledRef = useRef(false)
  /** password prompt for an ECMA-376 encrypted docx; submit retries via openDocxDecrypt */
  const [docPwdPrompt, setDocPwdPrompt] = useState<{
    path: string
    name: string
    value: string
    /** i18n key of the inline failure line ('' = none) */
    errorKey: '' | 'appDocPwdWrong' | 'appDocPwdUnsupported'
    busy: boolean
  } | null>(null)
  /** Review > Protect: the combined Word-style Protect Document dialog */
  const [showProtectDialog, setShowProtectDialog] = useState(false)
  /** prompt for a document with a password to modify (w:writeProtection): enter it or open read-only */
  const [modifyPwdPrompt, setModifyPwdPrompt] = useState<{
    value: string
    errorKey: '' | 'appDocPwdWrong'
  } | null>(null)
  const [_recent, setRecent] = useState<string[]>([])
  const [settings, setSettings] = useState<AiSettings>(DEFAULT_SETTINGS)
  const [showAi, setShowAi] = useState(() => aiPanelInitiallyOpen('aidocs.showAi'))
  // Settings › "show the AI sidebar" off hides the panel and its rail the way Word
  // treats the Copilot pane: an explicit Ask-AI action (ribbon, context menu,
  // shortcut, preset) still opens the panel for this document, and closing it
  // returns to hidden. The dock stays mounted either way, so panel state and
  // in-flight runs survive; the hook closes it when the switch itself goes off.
  const aiSidebarEnabled = useAiSidebarEnabled(setShowAi)
  const [spellcheck, setSpellcheck] = useState(spellcheckEnabled)
  const [largeDocSpellOff, setLargeDocSpellOff] = useState(false)
  const spellcheckActive = spellcheck && !largeDocSpellOff
  const spellcheckActiveRef = useRef(spellcheckActive)
  spellcheckActiveRef.current = spellcheckActive
  /** Increments on every open/new document: AiPanel remounts by key to reset the conversation and history (save path changes don't bump it, so the session continues) */
  const [aiPanelKey, setAiPanelKey] = useState(0)
  const [ribbonTabRequest, setRibbonTabRequest] = useState<{ tab: string; nonce: number } | null>(
    null,
  )
  const [status, setStatus] = useState('')
  const [zoom, setZoom] = useState(100)
  const scrollContainerRef = useRef<HTMLElement>(null)
  /** bumps after every pagination pass: the canvas' page gaps are in place */
  const [layoutRev, setLayoutRev] = useState(0)
  /** Multiple Pages grid: page index under the live editor, null while single-column */
  const multipageActiveRef = useRef<number | null>(null)
  const mpFramesRef = useRef<PageFrame[]>([])
  // Word-style dark page (editor/dark-page.ts): the shell's document-page-theme
  // setting (genoffice#1811) decides which theme the page follows; a light
  // page theme never shows a dark page, in the dark one it is on by default
  // until the user switches it off, and that choice sticks (dark-page-pref.ts)
  const themeDark = useUiThemeIsDark()
  const docThemeDark = useDocThemeIsDark()
  const [darkPage, setDarkPage] = useState(() => resolveDarkPage(docThemeDark))
  useEffect(() => {
    setDarkPage(resolveDarkPage(docThemeDark))
  }, [docThemeDark])
  // for toggle-by-one in the menu path, whose closure is not re-created per render
  const darkPageRef = useRef(darkPage)
  darkPageRef.current = darkPage
  /** View ▸ Dark Mode (menu and ribbon): remember the choice, then apply it */
  const updateDarkPage = useCallback(
    (next: boolean) => {
      if (docThemeDark) writeDarkPagePref(next)
      setDarkPage(next)
    },
    [docThemeDark],
  )
  const [section, setSection] = useState<SectionSettings | null>(null)
  /** All sections (readSections): pagination/preview use per-section geometry; layout edits apply to the cursor's section */
  const [sections, setSections] = useState<SectionInfo[]>([])
  const [sectionDirty, setSectionDirty] = useState(false)
  /** Index of the cursor's section (maintained by selectionUpdate) */
  const [activeSection, setActiveSection] = useState(0)
  // displayed number of the caret's page: a mirrored document's ruler follows its parity
  const [caretPageNo, setCaretPageNo] = useState(1)
  const locateCaretPageRef = useRef<() => void>(() => {})
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
  const [mirrorMargins, setMirrorMargins] = useState(false)
  const [mirrorMarginsDirty, setMirrorMarginsDirty] = useState(false)
  const pageSetupStateRef = useRef<PageSetupState | null>(null)
  /** a page setup command is running: the state it started from, and whether it also edited the document */
  const pageSetupUndoRef = useRef<{ before: PageSetupState; edited: boolean } | null>(null)
  /** call before a command changes page setup: the change becomes one undo step once state settles */
  const armPageSetupUndo = () => {
    if (pageSetupUndoRef.current || !pageSetupStateRef.current) return
    pageSetupUndoRef.current = { before: pageSetupStateRef.current, edited: false }
    forceRender()
  }
  const [autoHyphenation, setAutoHyphenation] = useState(false)
  const [autoHyphenationDirty, setAutoHyphenationDirty] = useState(false)
  const [hyphenationEdit, setHyphenationEdit] = useState<Partial<HyphenationSettings> | null>(null)
  /** a Manual… run is going through the document: the page shows where words would break */
  const [manualHyphenation, setManualHyphenation] = useState(false)
  const [showSpelling, setShowSpelling] = useState(false)
  const manualHyphenationRef = useRef(false)
  const [defaultTabStopEdit, setDefaultTabStopEdit] = useState<number | null>(null)
  /** settings.xml w:gutterAtTop as set in this session; null when untouched */
  const [gutterAtTopEdit, setGutterAtTopEdit] = useState<boolean | null>(null)
  /** Convert Permanently of Font Substitution: what the save has to write and what earlier saves wrote; null when untouched */
  const [fontConversions, setFontConversions] = useState<FontConversions | null>(null)
  /** settings.xml m:mathPr as set in this session; null when untouched */
  const [mathPrEdit, setMathPrEdit] = useState<MathPr | null>(null)
  /** footnote / endnote properties set in this session (only the fields changed); null when untouched */
  const [notePropsEdit, setNotePropsEdit] = useState<NoteSettings | null>(null)
  /** bumped when the marks were renumbered outside an edit of the note lists */
  const [noteRenumbered, setNoteRenumbered] = useState(0)
  const noteSettings = useMemo<NoteSettings>(
    () => ({
      footnote: { ...doc?.parsed.footnoteProps, ...notePropsEdit?.footnote },
      endnote: { ...doc?.parsed.endnoteProps, ...notePropsEdit?.endnote },
    }),
    [doc, notePropsEdit],
  )
  const notesBeneathText = noteSettings.footnote.pos === 'beneathText'
  /** note numbering in effect per section: settings.xml and this session's document-wide edit, then the section's own */
  const documentNoteProps = useMemo<NoteSettings>(
    () => ({
      footnote: { ...doc?.parsed.settingsNoteProps?.footnote, ...notePropsEdit?.footnote },
      endnote: { ...doc?.parsed.settingsNoteProps?.endnote, ...notePropsEdit?.endnote },
    }),
    [doc, notePropsEdit],
  )
  const sectionNoteProps = useMemo(
    () =>
      sectionNotePropsOf(
        documentNoteProps,
        // the last section is written from its own state
        sections.length > 0
          ? sections.map((s, i) => (i === sections.length - 1 && section) || s.settings)
          : section
            ? [section]
            : [],
      ),
    [documentNoteProps, sections, section],
  )
  /** settings.xml w:characterSpacingControl as set in this session; null when untouched */
  const [compressPunctEdit, setCompressPunctEdit] = useState<boolean | null>(null)
  /** Track moves set in this session; the document carries it (settings.xml w:doNotTrackMoves) */
  const [trackMovesEdit, setTrackMovesEdit] = useState<boolean | null>(null)
  const trackMoves = trackMovesEdit ?? doc?.parsed.doNotTrackMoves !== true
  /** settings.xml w:strictFirstAndLastChars as set in this session; null when untouched */
  const [strictKinsokuEdit, setStrictKinsokuEdit] = useState<boolean | null>(null)
  // read where the document CSS is rebuilt outside a render (font and style refreshes)
  const autoHyphenationRef = useRef(autoHyphenation)
  autoHyphenationRef.current = autoHyphenation
  // a typed line grid set in the Document dialog: the document CSS is built from the
  // file, so it is told what the sections say now (for the document it was set on)
  const liveTypedGridRef = useRef<{ parsed: ParsedDocFull; typed: boolean } | null>(null)
  const liveDocCss = useCallback(
    (parsed: ParsedDocFull): string =>
      docStyleCss(
        {
          ...parsed,
          autoHyphenation: autoHyphenationRef.current || manualHyphenationRef.current,
        },
        liveTypedGridRef.current?.parsed === fileCtxRef.current.doc?.parsed
          ? { typedGrid: liveTypedGridRef.current?.typed }
          : undefined,
      ),
    [],
  )
  const [pageInfo, setPageInfo] = useState({ current: 1, total: 1 })
  // last page's number for the document-end footer: text for the page marker, num for
  // even/odd parity (section restarts / pageNumberFmt make both differ from the physical count)
  const [lastPageNo, setLastPageNo] = useState<{ text: string; num: number } | null>(null)
  const [firstPageNo, setFirstPageNo] = useState<string | null>(null)
  /** SECTIONPAGES of the first and last pages (the canvas's own header and footer strips) */
  const [edgeSectionPages, setEdgeSectionPages] = useState<{ first: number; last: number } | null>(
    null,
  )
  /** Page-number-format dialog + pending final-section pgNumType write (non-final sections rewrite sectPr via pgNumDirtySections) */
  const [pgNumModal, setPgNumModal] = useState<{
    fmt: string
    start: string
    restart: boolean
    /** the heading level whose number leads the page number; '' = none */
    chapStyle: string
    chapSep: string
  } | null>(null)
  const [pgNumEdit, setPgNumEdit] = useState<{
    fmt?: string
    start?: number
    chapter?: PageNumberChapter | null
  } | null>(null)
  const [pgNumDirtySections, setPgNumDirtySections] = useState<number[]>([])
  /** Pending numbering definitions to append (saved via SaveOptions.numbering; restart numbering / new list definitions) */
  const [pendingNumbering, setPendingNumbering] =
    useState<PendingNumbering>(EMPTY_PENDING_NUMBERING)
  const numberingDirty = pendingNumberingDirty(pendingNumbering)
  /** Header/footer edits for non-final sections (key = `${lastBlockIndex}:${kind}`), saved via SaveOptions.sectionHf */
  const [sectionHfEdits, setSectionHfEdits] = useState<Record<string, HeaderFooter>>({})
  /** Link to Previous switched on for `${sectionIdx}:${kind}`: the section's own default reference goes on save */
  const [hfLinks, setHfLinks] = useState<Record<string, true>>({})
  /** the strip whose editor is open (drives the contextual Header & Footer tab) */
  const [hfEditing, setHfEditing] = useState<{
    kind: 'header' | 'footer'
    section: number
    variant: HfView
    /** absolute page index of the strip (formatted numbers repeat across restarted sections) */
    pageIdx: number
    /** section whose part the strip showed when the editor opened */
    owner: number
  } | null>(null)
  const hfEditingRef = useRef(hfEditing)
  hfEditingRef.current = hfEditing
  /** an edge strip commits to the variant its editor opened with, not the view a toggle switched to meanwhile */
  const edgeHfTarget = (kind: 'header' | 'footer', view: HfView) => {
    const open = hfEditingRef.current
    return open?.kind === kind
      ? { variant: open.variant, owner: open.owner }
      : { variant: view, owner: undefined }
  }
  /** bumps on document changes so derived section state follows deleted break paragraphs */
  const [docVersion, setDocVersion] = useState(0)
  const hfEditorRef = useRef<HfEditorHandle | null>(null)
  const [hfAreaEditRequest, setHfAreaEditRequest] = useState<{
    kind: 'header' | 'footer'
    nonce: number
  } | null>(null)
  const multiHf = sections.length > 1
  /** Variant an on-canvas edge area shows/edits: like Word, it follows its page — the header area is page 1 and the footer area the last page */
  const areaView = (kind: 'header' | 'footer'): HfView =>
    restingHfAreaVariant(kind, {
      titlePg,
      evenOddHf,
      pageCount: pageInfo.total,
      ...(lastPageNo ? { lastPageNo: lastPageNo.num } : {}),
    })
  const headerAreaView = areaView('header')
  const footerAreaView = areaView('footer')

  /** the AI header/footer tool edits the strip the matching edge area shows */
  const commitHf = (kind: 'header' | 'footer', next: HeaderFooter, viewOverride?: HfView) => {
    commitHfAt(
      kind === 'header' ? 0 : lastSectionIdx,
      kind,
      viewOverride ?? (kind === 'header' ? headerAreaView : footerAreaView),
      next,
    )
  }
  /** "Different first page" toggle (ribbon checkbox); page 1's header/footer become the first-page variant, as in Word */
  const toggleTitlePg = (on: boolean) => {
    hfEditorRef.current?.exit()
    armPageSetupUndo()
    setTitlePg(on)
    setTitlePgDirty(true)
    setStatus(on ? t('appTitlePgOn') : t('appTitlePgOff'))
  }
  const [showMarks, setShowMarks] = useState(false)
  // Word keeps View ▸ Ruler on across documents and launches
  const [showRuler, setShowRuler] = useStoredFlag('aidocs.showRuler')
  const [leftPane, setLeftPane] = useState<LeftPane>(loadLeftPane)
  const showNav = leftPane === 'nav'
  const setShowNav = useCallback((on: boolean) => setLeftPane(on ? 'nav' : null), [])
  const closeNav = useCallback(() => setLeftPane(null), [])
  const [navTab, setNavTab] = useState<NavTab>('headings')
  const [navTabRequest, setNavTabRequest] = useState<{ tab: NavTab; nonce: number } | null>(null)
  const clearNavTabRequest = useCallback(() => setNavTabRequest(null), [])
  const showThumbnails = showNav && navTab === 'pages'
  // Review → Reviewing Pane: the outermost right-edge panel (Word's vertical pane)
  const [showReviewPane, setShowReviewPane] = useState(false)
  const [viewMode, setViewMode] = useState<ViewMode>('print')
  const [readMode, setReadMode] = useState(false)
  // Story Focus: dim everything but the sentence the caret sits in (the
  // story-view bar's toggle; the plugin draws the dim ranges)
  const [storyFocus, setStoryFocus] = useState(false)
  // the view Story View returns to when it closes (Leave Story, Esc): the one
  // the document was in before entering
  const [storyReturnView, setStoryReturnView] = useState<ViewMode>('print')
  // Read mode is a print-layout reading, so a view pick and Read Mode are exclusive
  const switchView = (mode: ViewMode) => {
    setViewMode(mode)
    setReadMode(false)
  }
  const setReadModeExclusive = (v: boolean) => {
    if (v) setViewMode('print')
    setReadMode(v)
  }
  const [showGrid, setShowGrid] = useState(false)
  const [splitView, setSplitView] = useState(false)
  const [showPagePreview, setShowPagePreview] = useState(false)
  const [splitHtml, setSplitHtml] = useState('')
  const [showFind, setShowFind] = useState(false)
  const [imageDragOver, setImageDragOver] = useState(false)
  const [findFocusInput, setFindFocusInput] = useState(0)
  const [findFocusReplace, setFindFocusReplace] = useState(0)
  const [findFocusGoTo, setFindFocusGoTo] = useState(0)
  const [findStep, setFindStep] = useState<{ dir: 1 | -1; nonce: number } | null>(null)
  // last pagination pass, read by the Go To tab (page / section / line targets)
  const pageLayoutRef = useRef<PageLayout | null>(null)
  /** live section list of the last pass (header obstacles of the next pass) */
  const lastSecListRef = useRef<SectionInfo[] | null>(null)
  const cutsNow = useCallback(() => pageLayoutRef.current?.cuts, [])
  /** runs the debounced pagination pass now (set by the scheduler effect);
   *  `fresh` runs one even when none is pending */
  const flushPaginationRef = useRef<(fresh?: boolean) => void>(() => {})
  /** the current pass's layout, after any pending pass: readers that need
   *  the document as it stands (TOC numbers, the preview snapshot) */
  const layoutNow = useCallback((fresh = false): PageLayout | null => {
    flushPaginationRef.current(fresh)
    return pageLayoutRef.current
  }, [])
  const [showShortcuts, setShowShortcuts] = useState(false)
  const [showCustomizeKeyboard, setShowCustomizeKeyboard] = useState(false)
  const [showAutoCorrect, setShowAutoCorrect] = useState(false)
  const [autoTextList, setAutoTextList] = useState(autoTextEntries)
  const [showPreferences, setShowPreferences] = useState(false)
  const commandHostRef = useRef<CommandHost | null>(null)
  const runCommand = useCallback((command: ActionCommandId, payload?: string) => {
    if (commandHostRef.current) runMenuCommand(commandHostRef.current, command, payload)
  }, [])
  const ribbonActionsRef = useRef<{
    stepFontSize?: (dir: 1 | -1) => void
    nudgeFontSize?: (dir: 1 | -1) => void
    toggleRibbon?: () => void
  }>({})
  const [ribbonCollapsed, setRibbonCollapsed] = useState(false)
  // Show Comments (Word's top-right toggle): Contextual cards beside the page or the List panel
  const [showComments, setShowComments] = useState(true)
  const [commentsMode, setCommentsModeState] = useState<CommentsMode>(loadCommentsMode)
  const setCommentsMode = useCallback((mode: CommentsMode) => {
    saveCommentsMode(mode)
    setCommentsModeState(mode)
    setShowComments(true)
  }, [])
  // thread popped up from its anchor icon while Show Comments is off
  const [commentPopup, setCommentPopup] = useState<string | null>(null)
  // Contextual mode opens the panel for a composer or a focus the margin cannot host; it stays until closed
  const [commentPanelOpen, setCommentPanelOpen] = useState(false)
  const [showStylesPane, setShowStylesPane] = useState(false)
  const [showSelectionPane, setShowSelectionPane] = useState(false)
  const [showCitationsPane, setShowCitationsPane] = useState(false)
  const [commentFocus, setCommentFocus] = useState<{
    id: string
    nonce: number
    reply?: boolean
  } | null>(null)
  /** Style definitions pending write-back (key = styleId), saved via SaveOptions.styleUpserts */
  const [defaultFonts, setDefaultFonts] = useState<DefaultFonts>()
  const fontSettingsVersionRef = useRef(0)
  const fontSettingsPendingRef = useRef<Promise<void> | null>(null)
  const [styleUpserts, setStyleUpserts] = useState<Record<string, StyleUpsert>>({})
  const [comments, setCommentsState] = useState<CommentInfo[]>([])
  // Synchronous mirror of the comments state: an agent turn can run several
  // reply_comment calls with no render in between, and each id allocation
  // must see the previous write (stale reads minted duplicate comment ids).
  const commentsLiveRef = useRef<CommentInfo[]>([])
  const setComments = useCallback((action: SetStateAction<CommentInfo[]>) => {
    commentsLiveRef.current =
      typeof action === 'function' ? action(commentsLiveRef.current) : action
    setCommentsState(commentsLiveRef.current)
  }, [])
  const [commentsDirty, setCommentsDirty] = useState(false)
  const [watermark, setWatermark] = useState<string | null>(null)
  const [watermarkDirty, setWatermarkDirty] = useState(false)
  const [watermarkStyle, setWatermarkStyle] = useState<Omit<WatermarkSpec, 'text'> | null>(null)
  const [watermarkPicture, setWatermarkPicture] = useState<PictureWatermarkSpec | null>(null)
  // the pending picture watermark as the header image it will parse back as,
  // sized against the last section's margin box like the save path
  const pendingWatermarkImage = useMemo(() => {
    if (!watermarkDirty) return null
    if (watermarkPicture) {
      const s = sections[sections.length - 1]?.settings ?? section
      const box = s
        ? {
            widthPt: (s.pageWidth - s.marginLeft - s.marginRight) / 20,
            heightPt: (s.pageHeight - s.marginTop - s.marginBottom) / 20,
          }
        : null
      return pictureWatermarkPreviewImage(watermarkPicture, box)
    }
    if (watermark) return textWatermarkPreviewImage({ text: watermark, ...(watermarkStyle ?? {}) })
    return null
  }, [watermarkDirty, watermarkPicture, watermark, watermarkStyle, sections, section])
  /** Design tab / AI set_watermark: pending stores the save path drains */
  const applyWatermark = useCallback((spec: Watermark | null) => {
    const wm = typeof spec === 'string' ? { text: spec } : spec
    if (wm && 'image' in wm) {
      setWatermark(null)
      setWatermarkStyle(null)
      setWatermarkPicture(wm)
    } else if (wm) {
      const { text, ...style } = wm
      setWatermark(text)
      setWatermarkStyle(Object.keys(style).length > 0 ? style : null)
      setWatermarkPicture(null)
    } else {
      setWatermark(null)
      setWatermarkStyle(null)
      setWatermarkPicture(null)
    }
    setWatermarkDirty(true)
  }, [])
  const [inkAnnotations, setInkAnnotations] = useState<InkAnnotation[]>([])
  const [inksDirty, setInksDirty] = useState(false)
  const [inkTool, setInkTool] = useState<InkTool>('select')
  const [inkPresets, setInkPresetsState] = useState<InkPreset[]>(loadInkPresets)
  const [inkActiveId, setInkActiveId] = useState(() => inkPresets[0].id)
  const setInkPresets = useCallback((next: InkPreset[]) => {
    setInkPresetsState(next)
    saveInkPresets(next)
  }, [])
  const inkActive = inkPresets.find((p) => p.id === inkActiveId) ?? inkPresets[0]
  const [footnotes, setFootnotes] = useState<NoteInfo[]>([])
  /** Footnote ids already shown in canvas page gaps (the end-of-document list skips them to avoid duplication) */
  const [gapNoteIds, setGapNoteIds] = useState<Set<string>>(new Set())
  const [endnotes, setEndnotes] = useState<NoteInfo[]>([])
  /** Endnote-area anchor: measured flow-end Y (layout px from the page-wrap top); null until measured */
  const [endnotesAreaTop, setEndnotesAreaTop] = useState<number | null>(null)
  const [notesDirty, setNotesDirty] = useState(false)
  const [sources, setSources] = useState<SourceInfo[]>([])
  const [sourcesDirty, setSourcesDirty] = useState(false)
  const [citationStyle, setCitationStyle] = useState<CitationStyleId | null>(null)
  const [zoteroDocumentData, setZoteroDocumentDataState] = useState('')
  const [zoteroDocumentDataDirty, setZoteroDocumentDataDirty] = useState(false)
  const zoteroDocumentDataRef = useRef('')
  const setZoteroDocumentData = useCallback((value: string) => {
    zoteroDocumentDataRef.current = value
    setZoteroDocumentDataState(value)
  }, [])
  const zoteroNoteFields =
    footnotes.some((note) => note.zoteroField) || endnotes.some((note) => note.zoteroField)
  const [zoteroInstalled, setZoteroInstalled] = useState(false)
  useEffect(() => {
    let live = true
    const check = () =>
      void window.desktop.zoteroAvailable?.().then((found) => live && setZoteroInstalled(found))
    check()
    // Zotero installed or started while the document is open
    window.addEventListener('focus', check)
    return () => {
      live = false
      window.removeEventListener('focus', check)
    }
  }, [])
  const [themeFonts, setThemeFonts] = useState<ThemeFonts | null>(null)
  const [themeFontsDirty, setThemeFontsDirty] = useState(false)
  const [themeColors, setThemeColors] = useState<ThemeColors | null>(null)
  const [themeColorsDirty, setThemeColorsDirty] = useState(false)
  const [commentComposing, setCommentComposing] = useState(false)
  const [trackChanges, setTrackChanges] = useState(false)
  // only the document toggle (ribbon / status bar / shortcut) maps to w:trackRevisions
  const [trackRevisionsDirty, setTrackRevisionsDirty] = useState(false)
  const setTrackChangesByUser = useCallback((on: boolean) => {
    setTrackChanges(on)
    setTrackRevisionsDirty(true)
  }, [])
  const toggleTrackChangesByUser = useCallback(() => {
    setTrackChanges((v) => !v)
    setTrackRevisionsDirty(true)
  }, [])
  const [revisionDisplay, setRevisionDisplay] = useState<RevisionDisplayMode>('all')
  /** the user picked a Markup view this session: opening another document keeps it */
  const revisionDisplayRef = useRef(revisionDisplay)
  revisionDisplayRef.current = revisionDisplay
  const [revisionFilter, setRevisionFilter] = useState<RevisionFilter>(DEFAULT_REVISION_FILTER)
  const revisionFilterRef = useRef(revisionFilter)
  revisionFilterRef.current = revisionFilter
  /** Show Markup → Balloons; an application preference like Word's */
  const [balloonMode, setBalloonMode] = useState<BalloonMode>(loadBalloonMode)
  useEffect(() => {
    localStorage.setItem(BALLOON_MODE_KEY, balloonMode)
  }, [balloonMode])
  /** Track Changes Options (Word's Advanced Options): screen markup styling + recorder switches */
  const [tcPrefs, setTcPrefs] = useState<TrackChangesOptions>(loadTrackChangesOptions)
  const tcOptions = useMemo(() => ({ ...tcPrefs, trackMoves }), [tcPrefs, trackMoves])
  const setTcOptions = (next: TrackChangesOptions) => {
    setTcPrefs(next)
    if (next.trackMoves !== trackMoves) setTrackMovesEdit(next.trackMoves)
  }
  const [showTcOptions, setShowTcOptions] = useState(false)
  const [showHighlightChanges, setShowHighlightChanges] = useState(false)
  const tcOptionsStyle = useMemo(() => trackChangesOptionsCss(tcOptions), [tcOptions])
  // the original view restores old formatting via decorations: sync the mode and trigger one repaint
  const balloonsActive = markupPrintsIn(revisionDisplay, viewMode)
  const commentCardsPossible =
    balloonsActive && balloonMode !== 'inline' && revisionDisplay !== 'simple'
  const commentsUi = commentsView(
    showComments,
    commentsMode,
    commentCardsPossible,
    commentPanelOpen,
  )
  useEffect(() => {
    if (commentComposing) setCommentPanelOpen(true)
  }, [commentComposing])
  useEffect(() => {
    if (commentFocus && !commentCardsPossible) setCommentPanelOpen(true)
  }, [commentFocus, commentCardsPossible])
  useEffect(() => {
    revisionDisplayState.mode = revisionDisplay
    revisionDisplayState.balloons = balloonMode
    revisionDisplayState.balloonsActive = balloonsActive
    if (editor) editor.view.dispatch(editor.state.tr.setMeta('addToHistory', false))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [revisionDisplay, balloonMode, balloonsActive])
  // section breaks on tracked-deleted paragraph marks don't break in Word's
  // markup views (only the Original view restores them)
  const delSectBreaks = useMemo(
    () =>
      revisionDisplay === 'original'
        ? undefined
        : new Set(
            (doc?.parsed.blocks ?? [])
              .filter((b) => b.paraMarkDel && b.docxIndex != null)
              .map((b) => b.docxIndex as number),
          ),
    [doc, revisionDisplay],
  )
  const [protection, setProtection] = useState<DocProtection | null>(null)
  const [protectionDirty, setProtectionDirty] = useState(false)
  const [writeProtection, setWriteProtection] = useState<WriteProtection | null>(null)
  const [writeProtectionDirty, setWriteProtectionDirty] = useState(false)
  const [removePersonalInfo, setRemovePersonalInfo] = useState(false)
  const [removePersonalInfoDirty, setRemovePersonalInfoDirty] = useState(false)
  /** modify password entered (or set by the user this session) — false = write-locked, document read-only */
  const [modifyUnlocked, setModifyUnlocked] = useState(true)
  const [compareDialog, setCompareDialog] = useState<'compare' | 'combine' | null>(null)
  const [autoSave, setAutoSave] = useAutoSavePref('aidocs.autoSave', window.desktop)
  // tab closed but this renderer kept alive (shell freeze workaround): go inert
  const [tornDown, setTornDown] = useState(false)
  const [aiPreset, setAiPreset] = useState<{
    text: string
    nonce: number
    autoRun?: boolean
  } | null>(null)
  // selection-scoped AI edit queue (anchors live as editor decorations)
  const [editQueue, setEditQueue] = useState<DocsEditQueueItem[]>([])
  const editQueueRef = useRef(editQueue)
  editQueueRef.current = editQueue
  const queueSeqRef = useRef(0)
  // opening/creating a document drops every anchor with setContent; the queue
  // must not leak the previous file's items (they would sit orphaned at the cap)
  useEffect(() => {
    setEditQueue([])
  }, [aiPanelKey])
  const [docCss, setDocCss] = useState('')
  // Live CJK-ness of the body while editing; overrides docCss's --doc-line-factor
  const [liveDocCjk, setLiveDocCjk] = useState<boolean | null>(null)
  // the same answer as a ref, so the per-edit gate can skip a whole-document
  // scan without waiting for React (null = the current document is unscanned)
  const liveDocCjkRef = useRef<boolean | null>(null)
  // header/footer push-down re-measures after the doc-scoped <style> elements
  // commit: the DOM probe (hfReservedHeightPx) keys on the mounted CSS content,
  // so a value computed in the render pass that introduced new CSS is replaced
  // by a post-commit measure instead of going stale
  const [hfMeasureEpoch, setHfMeasureEpoch] = useState(0)
  // docCss already carries the parse-time factor: the live override only
  // exists once the body's CJK-ness differs from what parsing saw. Keying the
  // commit event on it (not on liveDocCjk) spares the shrink measurers a
  // second whole-document pass when the first live recompute agrees.
  const parsedDocCjk = useMemo(() => (doc ? docHasCjk(doc.parsed) : false), [doc])
  const liveLineFactor =
    doc && liveDocCjk != null && liveDocCjk !== parsedDocCjk
      ? docLineFactor(doc.parsed, liveDocCjk)
      : null
  useEffect(() => {
    setHfMeasureEpoch((e) => e + 1)
    // measurement views that read computed alignment/spacing re-run once the
    // stylesheet is in the DOM (setContent measured against the previous one)
    document.dispatchEvent(new Event(DOC_CSS_COMMITTED_EVENT))
  }, [docCss, liveLineFactor, themeFonts, themeColors, revisionDisplay])
  // header and footer strips break by the document's hyphenation rules: other rules, other heights
  useEffect(() => onDocHyphenationChange(() => setHfMeasureEpoch((e) => e + 1)), [])
  const [stats, setStats] = useState<{ body: DocStats; notes: DocStats } | null>(null)
  const [showLinkModal, setShowLinkModal] = useState(false)
  const [fieldEdit, setFieldEdit] = useState<FieldRange | null>(null)
  const [showTableModal, setShowTableModal] = useState(false)
  const [showZoomDialog, setShowZoomDialog] = useState(false)
  const [showEquationModal, setShowEquationModal] = useState(false)
  /** the Insert Chart dialog and the type it opens on; null while closed */
  const [chartModal, setChartModal] = useState<{ type?: ChartType; pick: number } | null>(null)
  const [showBookmarkModal, setShowBookmarkModal] = useState(false)
  const [showCrossRefModal, setShowCrossRefModal] = useState(false)
  const [showCaptionModal, setShowCaptionModal] = useState(false)
  useEffect(() => setAutoCaptionBlocks(() => fileCtxRef.current.doc?.parsed.blocks ?? []), [])
  const [showDocumentDialog, setShowDocumentDialog] = useState(false)
  const [showChangeCase, setShowChangeCase] = useState(false)
  const [showDateTime, setShowDateTime] = useState(false)
  const [showDropCap, setShowDropCap] = useState(false)
  const [showWatermarkDialog, setShowWatermarkDialog] = useState(false)
  const [showColumnsDialog, setShowColumnsDialog] = useState(false)
  const [showSymbolDialog, setShowSymbolDialog] = useState(false)
  const [showIconsDialog, setShowIconsDialog] = useState(false)
  const [showPageNumbers, setShowPageNumbers] = useState(false)
  const [showBulletsNumbering, setShowBulletsNumbering] = useState(false)
  const [showTranslationLanguage, setShowTranslationLanguage] = useState(false)
  const [showIndexTables, setShowIndexTables] = useState(false)
  const [showHyphenation, setShowHyphenation] = useState(false)
  const [showEquationOptions, setShowEquationOptions] = useState(false)
  const [showFootnoteDialog, setShowFootnoteDialog] = useState(false)
  const [showRevealFormatting, setShowRevealFormatting] = useState(false)
  const [showRecipients, setShowRecipients] = useState(false)
  const [showFilterRecipients, setShowFilterRecipients] = useState(false)
  const [showFindRecipient, setShowFindRecipient] = useState(false)
  const [showProperties, setShowProperties] = useState(false)
  const [showRename, setShowRename] = useState(false)
  const [showBordersShading, setShowBordersShading] = useState(false)
  const [showFontSubstitution, setShowFontSubstitution] = useState(false)
  const [mergeRule, setMergeRule] = useState<MergeRule | null>(null)
  const [showNumber, setShowNumber] = useState(false)
  const [showSort, setShowSort] = useState(false)
  const [showTextDirection, setShowTextDirection] = useState(false)
  const [showAltText, setShowAltText] = useState(false)
  const [showLineNumbers, setShowLineNumbers] = useState(false)
  const [showReduceFileSize, setShowReduceFileSize] = useState(false)
  const [reduceSelection, setReduceSelection] = useState<string[]>([])
  const [showTemplateGallery, setShowTemplateGallery] = useState(false)
  const [insertFileRange, setInsertFileRange] = useState<{
    fileName: string
    bookmarks: string[]
    answer: (range: string | null | undefined) => void
  } | null>(null)
  const [showParagraphSpacing, setShowParagraphSpacing] = useState(false)
  const [showSizePosition, setShowSizePosition] = useState(false)
  // Word's Focus: the page alone, still editable; Esc brings the window back
  const [focusMode, setFocusMode] = useState(false)
  useEffect(() => {
    if (!focusMode) return
    const onKey = (e: KeyboardEvent) => {
      // an open dialog takes the Escape for itself
      if (e.key === 'Escape' && !document.querySelector('.modal-backdrop')) setFocusMode(false)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [focusMode])
  const [propertyChanges, setPropertyChanges] = useState<PropertyChanges | null>(null)
  const [showInsertField, setShowInsertField] = useState(false)
  const [pasteSpecialTarget, setPasteSpecialTarget] = useState<Editor | null>(null)
  const [tabsTarget, setTabsTarget] = useState<Editor | null>(null)
  const [fitTextTarget, setFitTextTarget] = useState<Editor | null>(null)
  const [textEffectsTarget, setTextEffectsTarget] = useState<Editor | null>(null)
  const [languageTarget, setLanguageTarget] = useState<Editor | null>(null)
  const [markCitationAt, setMarkCitationAt] = useState<{
    text: string
    from: number
    at: number
  } | null>(null)
  const [markEntryAt, setMarkEntryAt] = useState<{ text: string; at: number } | null>(null)
  const [autoText, setAutoText] = useState<{
    mode: 'new' | 'insert'
    suggestedName: string
    entries: AutoTextEntry[]
  } | null>(null)
  const [eqLayoutDialog, setEqLayoutDialog] = useState<{
    editor: Editor
    kind: EqLayoutKind
    target: EqLayoutTarget
  } | null>(null)
  const [phoneticGuide, setPhoneticGuide] = useState<{ editor: Editor; target: RubyTarget } | null>(
    null,
  )
  const [pictureDialog, setPictureDialog] = useState<{
    kind: PictureDialogKind
    dataUrl: string
  } | null>(null)
  const [eqEditTarget, setEqEditTarget] = useState<{
    pos: number
    latex: string
    kind: 'inline' | 'block'
  } | null>(null)
  const [ctxMenu, setCtxMenu] = useState<ContextMenuState | null>(null)
  const ctxSeq = useRef(0)
  useEffect(() => {
    window.desktop.armContextMenu?.()
    return window.desktop.onContextMenuRequest?.((req) => {
      setCtxMenu((prev) =>
        prev && prev.seq === req.seq
          ? {
              ...prev,
              spell: req.misspelledWord
                ? { word: req.misspelledWord, suggestions: req.suggestions }
                : null,
            }
          : prev,
      )
    })
  }, [])
  const [viewImage, setViewImage] = useState<string | null>(null)
  useEffect(() => window.desktop.onViewImage?.((src) => setViewImage(src)), [])
  const saveImageAs = useCallback((src: string) => void window.desktop.saveImageAs(src), [])
  const [showFontDialog, setShowFontDialog] = useState(false)
  const [showParaDialog, setShowParaDialog] = useState(false)
  /** label the reader gives a span they are withholding from the model */
  const [redactTarget, setRedactTarget] = useState<{
    from: number
    to: number
    seed: string
  } | null>(null)
  /** Word's table dialogs, shared by the ribbon, the right-click menu and the Table menu */
  const [tableDialog, setTableDialog] = useState<TableDialogKind | null>(null)
  const [tableGridlines, setTableGridlines] = useState(
    () => localStorage.getItem('aidocs.tableGridlines') === '1',
  )
  const toggleTableGridlines = useCallback(
    () =>
      setTableGridlines((on) => {
        localStorage.setItem('aidocs.tableGridlines', on ? '0' : '1')
        return !on
      }),
    [],
  )
  // section content width the table commands / dialogs fit the grid into (same formula as the ribbon)
  const tableSection = sections[activeSection]?.settings ?? section
  /** the default paragraph style's spacing: what was set in this session, else what the file says */
  const defaultStyleSpacing = () => {
    const styles = ribbonStylesRef.current
    const styleId = defaultParagraphStyleId(styles) ?? 'Normal'
    const pending = styleUpserts[styleId]?.pPr
    const info = styles?.get(styleId)?.display
    return {
      before: pending?.spaceBeforeTwips ?? info?.spaceBeforeTwips,
      after: pending?.spaceAfterTwips ?? info?.spaceAfterTwips,
      line: pending?.lineSpacing ?? info?.lineSpacing,
      // a pending spacing is always a multiple; the file's may be an exact height
      multiple: pending?.lineSpacing != null || (info?.lineRule ?? 'auto') === 'auto',
    }
  }
  const defaultParagraphSpacing = () => {
    const { before, after, line, multiple } = defaultStyleSpacing()
    return { beforeTwips: before ?? 0, afterTwips: after ?? 0, line: multiple && line ? line : 1 }
  }
  // the last section's first-page choice is the live toggle; an earlier section's is what the file says
  const activeTitlePg =
    activeSection >= sections.length - 1 ? titlePg : (sections[activeSection]?.titlePg ?? false)
  const tableSectionWidthPx = tableSection
    ? Math.max(
        1,
        (tableSection.pageWidth - tableSection.marginLeft - tableSection.marginRight) / 15,
      )
    : 624
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
  const zoteroControllerRef = useRef<ZoteroDocumentController | null>(null)
  const editor = useEditor({
    extensions: editorExtensions,
    content: { type: 'doc', content: [{ type: 'docParagraph' }] },
    editorProps: {
      // Word checks spelling as you type by default; user toggle in Review → Spelling.
      // useEditor re-applies these options on every render, so the value must
      // be the effective one or a re-render undoes the large-document override
      attributes: { class: 'doc-page', spellcheck: spellcheckActive ? 'true' : 'false' },
      // Word/web HTML cleanup: strip mso comments and <o:p>, unwrap <li><p>x</p></li>
      // (docListItem is an inline container; block-level p would shatter the list)
      transformPastedHTML: cleanPastedHtml,
      // plain-text paste takes the insertion point's formatting like typing
      // (Word Keep Text Only) — the default parser misses storedMarks (r172)
      clipboardTextParser: (text, $context, _plain, view) => pasteTextSlice(text, $context, view),
      handleDOMEvents: {
        paste: (_view, event) => {
          const html = event.clipboardData?.getData('text/html') ?? ''
          const armed = beginForeignPaste(html)
          stashPastePayload(
            armed
              ? {
                  html,
                  text: event.clipboardData?.getData('text/plain') ?? '',
                  mode: defaultPasteMode(),
                }
              : null,
          )
          return false
        },
      },
      // foreign-HTML paste, by the armed mode: 'source' fills runs that still
      // have no concrete font after the parse (generic-only family chains
      // like `sans-serif`) with the insertion point's font slots (r181);
      // 'merge' formats the whole fragment like typing at the insertion
      // point, keeping emphasis and structure; 'text' passes through — the
      // handlePaste reroute below sends it down the plain-text lane
      transformPasted: (slice, view) => {
        const mode = consumeForeignPaste()
        if (!mode || mode === 'text') return slice
        const schema = view.state.schema
        const markType = schema.marks.docTextStyle
        if (!markType) return slice
        const marks = view.state.storedMarks ?? view.state.selection.$from.marks()
        if (mode === 'merge') {
          return new PmSlice(
            mergeFormattingFragment(
              slice.content,
              schema,
              caretStyleMark(marks, schema),
              caretParaFormat(view.state.selection.$from, schema),
            ),
            slice.openStart,
            slice.openEnd,
          )
        }
        const slots = caretFontSlots(marks, markType)
        if (!slots) return slice
        return new PmSlice(
          fillMissingRunFonts(slice.content, slots, schema),
          slice.openStart,
          slice.openEnd,
        )
      },
      // image files dragged from the OS (or a browser) land as inline pictures
      // at the drop point, like Word; other drops keep the default handling
      handleDrop: (_view, event) => {
        const ed = editorRef.current
        const files = imageFilesFromDataTransfer(event.dataTransfer)
        if (!ed || files.length === 0) return false
        return insertImageFilesAtCoords(ed, files, { left: event.clientX, top: event.clientY })
      },
      // clipboard images (screenshots/copied images): become inline images; mixed content
      // with HTML still uses default parsing (text in the HTML takes priority)
      handlePaste: (view, event) => {
        const data = event.clipboardData
        if (!data) return false
        const html = data.getData('text/html')
        // text-mode foreign paste (default setting or chip choice): the whole
        // payload goes down the plain-text lane, formatting like typing
        if (takeTextReroute()) {
          const plain = data.getData('text/plain') || htmlFallbackText(html)
          if (plain) {
            const slice = pasteTextSlice(plain, view.state.selection.$from, view)
            view.dispatch(view.state.tr.replaceSelection(slice).scrollIntoView())
            return true
          }
        }
        // a lone unformatted spreadsheet cell is a text paste (the r146 unwrap
        // turns it into text) and takes the insertion point's formatting like
        // typing (r172) — the HTML lanes below keep no marks for it and would
        // land it in the theme font (r176: Sheets cell → Docs pasted as Aptos)
        if (html) {
          const cellText = singleCellPasteText(html)
          if (cellText !== null) {
            const slice = pasteTextSlice(cellText, view.state.selection.$from, view)
            view.dispatch(view.state.tr.replaceSelection(slice).scrollIntoView())
            return true
          }
        }
        // web-copied images (r139): "Copy image" in a browser puts a bitmap +
        // an <img src="http..."> HTML fragment + often the URL as text/plain
        // on the clipboard. Detect image-only HTML so the bitmap wins over
        // text parsing, and so a bitmap-less copy can fetch the referenced
        // image instead of pasting nothing (our parse rules are data:-only).
        let htmlImgSrcs: string[] = []
        let htmlHasText = false
        if (html) {
          try {
            const imgDom = new window.DOMParser().parseFromString(html, 'text/html')
            htmlImgSrcs = [...imgDom.querySelectorAll('img')]
              .map((img) => img.getAttribute('src') ?? '')
              .filter(Boolean)
            htmlHasText = (imgDom.body.textContent ?? '').trim().length > 0
          } catch {
            /* unparseable html: treat as not-an-image copy */
          }
        }
        // remote-only: in-app data: copies carry data-image-meta and must keep
        // going through the DocProtected parse rules (size/align/wrap round-trip
        // preserved); only web copies divert to bitmap/fetch handling
        const imageOnlyHtml =
          htmlImgSrcs.length > 0 &&
          !htmlHasText &&
          htmlImgSrcs.every((src) => /^https?:\/\//i.test(src))
        // pasting block HTML onto an empty paragraph: replace wholesale (ProseMirror's default
        // first-block merge would demote headings to body text; Word keeps the source block format on empty paragraphs)
        const { $from, empty: selEmpty } = view.state.selection
        if (
          !imageOnlyHtml &&
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
              // this lane re-parses and dispatches itself, bypassing
              // transformPasted — apply the armed paste mode to foreign
              // fragments here too (r181; the empty paragraph's pilcrow
              // memory lives in storedMarks, caret-marks.ts)
              let content = parsed.content
              if (isForeignPasteHtml(html)) {
                const schema = view.state.schema
                const markType = schema.marks.docTextStyle
                const marks = view.state.storedMarks ?? $from.marks()
                if (lastForeignPasteMode() === 'merge') {
                  content = mergeFormattingFragment(
                    content,
                    schema,
                    caretStyleMark(marks, schema),
                    caretParaFormat($from, schema),
                  )
                } else {
                  const slots = markType ? caretFontSlots(marks, markType) : null
                  if (slots) content = fillMissingRunFonts(content, slots, schema)
                }
              }
              view.dispatch(view.state.tr.replaceWith($from.before(1), $from.after(1), content))
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
        // image-only copies keep the bitmap even when text/plain carries the
        // source URL (browser "Copy image" does that) — Word pastes the picture
        if (imageFile && (!text.trim() || imageOnlyHtml)) {
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
        // bitmap-less web image copy: fetch the referenced image(s) through the
        // hardened main-process fetcher (SSRF-guarded, CDN headers) and insert
        if (!imageFile && imageOnlyHtml) {
          void (async () => {
            for (const src of htmlImgSrcs.slice(0, 10)) {
              const ed = editorRef.current
              if (!ed) return
              const fetched = await window.desktop.fetchImage(src)
              if (fetched) {
                await insertImageFromDataUrl(
                  ed,
                  `data:${fetched.mime};base64,${fetched.base64}`,
                  'Image (pasted)',
                )
              }
            }
          })()
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
  useEffect(() => {
    saveTrackChangesOptions(tcOptions)
    trackChangesOptionsState.current = tcOptions
    if (!editor) return
    editor.storage.trackChanges.trackMoves = tcOptions.trackMoves
    editor.storage.trackChanges.trackFormatting = tcOptions.trackFormatting
    editor.view.dispatch(editor.state.tr.setMeta('addToHistory', false))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tcOptions, editor])

  // every header/footer consumer (strips, gaps, heights, preview, AI, save) reads
  // through one resolver, so pending edits, Link to Previous and inheritance agree
  // the section list as Word sees it right now: a deleted break paragraph merges its
  // section into the next, pending breaks are already in `sections`. Derived from the
  // document itself (not the paint-time measurement), so a strip edit lands in the right
  // slot even before the canvas has re-measured
  const hfSectionsMemoRef = useRef<SectionInfo[] | null>(null)
  const hfSections = useMemo(() => {
    if (!editor || sections.length <= 1) return sections
    const present: Array<{ docxIndex: number }> = []
    editor.state.doc.forEach((node) => {
      const di = node.attrs?.docxIndex
      if (typeof di === 'number') present.push({ docxIndex: di })
    })
    const next = liveSections(sections, present as BlockBox[], undefined, delSectBreaks)
    // every transaction re-runs this memo, and a streaming open appends a
    // chunk per frame: a fresh array per chunk re-ran the pagination effect
    // (a synchronous whole-document pass each) for every chunk. Unchanged
    // content keeps the previous identity.
    const prev = hfSectionsMemoRef.current
    if (prev && sameSectionInfos(prev, next)) return prev
    hfSectionsMemoRef.current = next
    return next
    // docVersion: the doc identity is what changes, and it is not a React value
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editor, sections, delSectBreaks, docVersion])
  const hfState = useMemo<HfSectionState>(
    () => ({
      sections: hfSections,
      edits: sectionHfEdits,
      links: hfLinks,
      docStrips: { header, footer, variants: hfVariants },
      parsed: doc?.parsed ?? null,
    }),
    [hfSections, sectionHfEdits, hfLinks, header, footer, hfVariants, doc],
  )
  const hfResolveAt = useCallback(
    (si: number, kind: 'header' | 'footer', variant: HfView) =>
      resolveHf(hfState, si, kind, variant),
    [hfState],
  )
  const hfValueAt = useCallback(
    (si: number, kind: 'header' | 'footer', variant: HfView) =>
      resolveHf(hfState, si, kind, variant).value,
    [hfState],
  )
  const hfLinkedAt = (si: number, kind: 'header' | 'footer', variant: HfView) =>
    hfLinked(hfState, si, kind, variant)
  const lastSectionIdx = Math.max(0, hfSections.length - 1)
  const shownHeader = hfValueAt(0, 'header', headerAreaView)
  const shownFooter = hfValueAt(lastSectionIdx, 'footer', footerAreaView)
  /** images of the strip an edge area shows (logos etc., display-only; floating shapes render separately) */
  const hfImagesOf = (kind: 'header' | 'footer') =>
    hfResolveAt(
      kind === 'header' ? 0 : lastSectionIdx,
      kind,
      kind === 'header' ? headerAreaView : footerAreaView,
    ).images?.filter((img) => !img.floating)
  /**
   * Write an edit made on section `si`'s strip to the part that owns it. `ownerAtOpen`
   * (the owner when the editor opened) refuses a write whose owner moved to an earlier
   * section meanwhile: Link to Previous under an open editor must not overwrite the
   * strip it reconnects to.
   */
  const commitHfAt = (
    si: number,
    kind: 'header' | 'footer',
    variant: HfView,
    next: HeaderFooter,
    ownerAtOpen?: number,
  ): boolean => {
    const target = hfCommitTarget(hfState, si, kind, variant, ownerAtOpen)
    if (!target) return false
    // functional: two commits in one turn (Remove Page Numbers on both strips, an
    // editor exit racing an AI set) must both land
    setSectionHfEdits((m) => ({ ...m, [target.key]: hfTextCommit(m[target.key], next) }))
    return true
  }
  /** a floating header picture moved or resized on section `si`'s strip
   *  (hf-float-edit); false when the strip has no part to take it */
  const commitHfPictureAt = (
    si: number,
    kind: 'header' | 'footer',
    variant: HfView,
    drawing: number,
    edit: HfPictureEdit,
  ): boolean => {
    if (!hfCommitTarget(hfState, si, kind, variant)) return false
    setSectionHfEdits(
      (m) =>
        withHfPictureEdit({ ...hfState, edits: m }, si, kind, variant, drawing, edit)?.edits ?? m,
    )
    return true
  }
  /** pending changes outside the body, refreshed with the revision badge (see storyRevisions memo) */
  const storyRevisionsRef = useRef<StoryRevision[]>([])
  const hfStateRef = useRef(hfState)
  hfStateRef.current = hfState
  // live mirrors: an AI turn applies several decisions between renders
  const footnotesLiveRef = useRef(footnotes)
  footnotesLiveRef.current = footnotes
  const endnotesLiveRef = useRef(endnotes)
  endnotesLiveRef.current = endnotes
  /** accept / reject changes in notes, header / footer parts and textboxes; the count applied */
  const applyStoryRevisions = (
    targets: StoryRevision[],
    mode: 'accept' | 'reject',
    all = false,
  ): number => {
    const live = storyRevisionsRef.current
    const picked = targets
      .map((t) => liveStoryRevision(live, t))
      .filter((r): r is StoryRevision => !!r)
    if (picked.length === 0 && !all) return 0
    const locs = (kind: StoryRevision['story']['kind']) =>
      picked.filter((r) => r.story.kind === kind).map((r) => r.story)
    const nextNotes = { footnote: footnotesLiveRef.current, endnote: endnotesLiveRef.current }
    for (const kind of ['footnote', 'endnote'] as const) {
      const mine = locs(kind)
      if (mine.length === 0 && !all) continue
      // Accept All also clears what the models do not show (paragraph marks): every entry replays
      nextNotes[kind] = all
        ? resolveNotes(nextNotes[kind], mine, mode, true).map((n) =>
            n.revisionsResolved?.some((r) => !r.filter)
              ? n
              : { ...n, revisionsResolved: [...(n.revisionsResolved ?? []), { mode }] },
          )
        : resolveNotes(nextNotes[kind], mine, mode)
      ;(kind === 'footnote' ? setFootnotes : setEndnotes)(nextNotes[kind])
      ;(kind === 'footnote' ? footnotesLiveRef : endnotesLiveRef).current = nextNotes[kind]
      setNotesDirty(true)
    }
    let state = hfStateRef.current
    const slots = new Map<string, StoryRevision[]>()
    for (const r of picked) {
      if (!r.story.hf) continue
      const key = `${r.story.hf.section}:${r.story.hf.kind}:${r.story.hf.variant}`
      slots.set(key, [...(slots.get(key) ?? []), r])
    }
    const hfTargets = all
      ? ownHfSlots(state).map((slot) => ({
          ...slot,
          revs: slots.get(`${slot.si}:${slot.kind}:${slot.variant}`) ?? [],
        }))
      : [...slots.values()].map((revs) => {
          const { section: si, kind, variant } = revs[0].story.hf!
          return { si, kind, variant, value: ownHf(state, si, kind, variant)?.value, revs }
        })
    for (const slot of hfTargets) {
      const { si, kind, variant, revs } = slot
      let value = slot.value
      if (!value) continue
      let pending = Object.values(state.edits).includes(value)
      // a strip editor open on this part would commit its pre-resolution runs on
      // exit: drain what was typed into the value first, then drop the editor
      const open = hfEditingRef.current
      const handle = hfEditorRef.current
      if (handle && open && open.kind === kind && open.owner === si) {
        const typed = handle.changed()
        handle.discard()
        if (typed) {
          value = typed
          pending = true
        }
      }
      const next = resolveHfValue(
        value,
        revs.map((r) => r.story),
        mode,
        pending,
        all,
      )
      const target = hfCommitTarget(state, si, kind, variant)
      if (!target) continue
      commitHfAt(si, kind, variant, next)
      state = { ...state, edits: { ...state.edits, [target.key]: next } }
      hfStateRef.current = state
    }
    if (editor) {
      const tr = editor.state.tr.setMeta(TRACK_IGNORE, true)
      const byPos = new Map<number, StoryRevision[]>()
      for (const r of picked) {
        if (r.story.textbox)
          byPos.set(r.story.textbox.pos, [...(byPos.get(r.story.textbox.pos) ?? []), r])
      }
      for (const [pos, revs] of byPos) {
        const node = tr.doc.nodeAt(pos)
        const boxes = node?.attrs?.textboxes as TextboxDisplay[] | null | undefined
        if (!node || !boxes) continue
        tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          textboxes: resolveTextboxes(
            boxes,
            revs.map((r) => r.story),
            mode,
          ),
        })
      }
      if (tr.docChanged) editor.view.dispatch(tr)
      // the list a follow-up in the same turn reads must already show the outcome
      storyRevisionsRef.current = [
        ...collectHfRevisions(state),
        ...collectTextboxRevisions(editor.state.doc),
        ...collectNoteRevisions(editor.state.doc, 'footnote', nextNotes.footnote),
        ...collectNoteRevisions(editor.state.doc, 'endnote', nextNotes.endnote),
      ]
    }
    dirtyRef.current = true
    return picked.length
  }
  const setHfLink = (si: number, kind: 'header' | 'footer', variant: HfView, on: boolean) => {
    if (
      on &&
      !window.confirm(t(kind === 'header' ? 'appHfLinkConfirmHeader' : 'appHfLinkConfirmFooter'))
    )
      return
    // the open editor holds the strip being discarded: it must not commit it anywhere
    if (on) hfEditorRef.current?.discard()
    if (!withHfLink(hfState, si, kind, variant, on)) return
    setSectionHfEdits(
      (m) => withHfLink({ ...hfState, edits: m }, si, kind, variant, on)?.edits ?? m,
    )
    setHfLinks((l) => withHfLink({ ...hfState, links: l }, si, kind, variant, on)?.links ?? l)
  }
  const setHfDistance = (si: number, kind: 'header' | 'footer', twips: number) => {
    // `si` indexes the live list; the saved list (what the save rewrites) is matched by its break paragraph
    const live = hfSections[si]
    const saved = live ? sections.findIndex((s) => s.lastBlockIndex === live.lastBlockIndex) : -1
    const cur = (saved >= 0 ? sections[saved].settings : undefined) ?? section
    if (!cur) return
    armPageSetupUndo()
    const next = { ...cur, [kind === 'header' ? 'headerDist' : 'footerDist']: twips }
    if (saved >= 0) editSectionSettings((_s, i) => (i === saved ? next : null))
    if (sections.length <= 1 || saved < 0 || saved === sections.length - 1) {
      setSection(next)
      setSectionDirty(true)
    } else setSectionsDirty((d) => (d.includes(saved) ? d : [...d, saved]))
  }

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
    const loadSettings = () => void window.desktop.getAiSettings().then(setSettings)
    loadSettings()
    return window.desktop.onAiSettingsChanged?.(loadSettings)
  }, [])

  useEffect(() => {
    rememberAiPanelOpen('aidocs.showAi', showAi)
  }, [showAi])

  useEffect(() => {
    saveLeftPane(leftPane)
  }, [leftPane])

  const spellcheckWasOn = useRef(spellcheckActive)
  const respellKickBusy = useRef(false)
  // bumped by the context menu after Add to Dictionary / a language change:
  // the session checker is updated but existing squiggles only refresh on typing
  const [respellNonce, setRespellNonce] = useState(0)
  const respellNonceSeen = useRef(0)
  // spell-diag trace: squiggle losses in the field are intermittent and
  // platform-bound, so the toggle/kick lifecycle leaves a line in
  // userData/spell-diag.log for support to collect (never breaks the app)
  const spellDiag = (line: string) => {
    try {
      window.desktop.spellDiag?.(line)
    } catch {
      /* diagnostics only */
    }
  }
  useEffect(() => {
    let cancelRespellKick = () => {}
    localStorage.setItem(SPELLCHECK_KEY, spellcheck ? '1' : '0')
    spellDiag(
      `toggle spellcheck=${spellcheckActive} platform=${navigator.platform} composing=${editor.view.composing}`,
    )
    // setOptions (not a direct DOM write): ProseMirror re-applies editorProps
    // attributes on every updateState, so only the props route sticks
    editor.setOptions({
      editorProps: {
        ...editor.options.editorProps,
        attributes: {
          ...(editor.options.editorProps.attributes as Record<string, string>),
          spellcheck: spellcheckActive ? 'true' : 'false',
        },
      },
    })
    // Blink only respells an editable as a consequence of real (trusted)
    // typing of a word-committing character inside it. A retest ruled
    // everything else out pixel by pixel: attribute flips, focus cycles (the
    // first attempt at this fix), script or execCommand edits, fresh DOM
    // nodes, session spellchecker kicks, synthetic clicks and arrow keys —
    // even a typed zero-width space — existing typos stay unmarked. So on the
    // re-enable transition, have the main process type one trusted space at
    // the caret and remove it again by script (a trusted Backspace would work
    // too, but its deletion re-suppresses the caret paragraph's markers).
    // ProseMirror must not see any of it: its DOM observer is paused (no
    // transaction, no history entry, no dirty flag) and a capture-phase key
    // shield keeps its keymap and other key handlers out of the round trip —
    // the same shield also tells us when the keystroke has actually landed.
    const respellRequested = respellNonce !== respellNonceSeen.current
    if (spellcheckActive && (!spellcheckWasOn.current || respellRequested)) {
      let attempts = 0
      // cancel a pending retry/rAF if the effect re-runs (another off/on) or the
      // component unmounts — two overlapping kick chains would each pause the DOM
      // observer and install the key shield, swallowing keystrokes twice
      let cancelled = false
      let rafId = 0
      let retryTimer: ReturnType<typeof setTimeout> | undefined
      cancelRespellKick = () => {
        cancelled = true
        if (rafId) cancelAnimationFrame(rafId)
        if (retryTimer !== undefined) clearTimeout(retryTimer)
      }
      const runKick = () => {
        if (cancelled) return
        const view = editor.view
        const dom = view.dom
        if (!dom.isConnected) return spellDiag('kick abandoned: editor gone')
        if (!spellcheckActiveRef.current) return spellDiag('kick canceled: toggled off again')
        // a live IME composition (French dead keys, CJK input) or an in-flight
        // kick used to make the re-enable return silently, so existing typos
        // were never re-marked (one "the toggle worked only once" path).
        // Retry for a few seconds instead of dropping the kick.
        if (view.composing || respellKickBusy.current) {
          spellDiag(
            `kick deferred composing=${view.composing} busy=${respellKickBusy.current} attempt=${attempts}`,
          )
          if (attempts++ < 8) retryTimer = setTimeout(runKick, 600)
          else spellDiag('kick gave up after retries')
          return
        }
        respellKickBusy.current = true
        spellDiag('kick start')
        const sub = getActiveSubEditor()
        const prev = document.activeElement as HTMLElement | null
        // the caret can sit pages away from the viewport (last edit position);
        // focusing and typing there scrolls it into view and the toggle
        // "jumps to another page". Pin the scroller
        // for the whole round trip: the reset runs inside the scroll event,
        // before paint, so no jump ever shows.
        const scroller = scrollContainerRef.current
        const pinTop = scroller?.scrollTop ?? 0
        const pinLeft = scroller?.scrollLeft ?? 0
        const pinScroll = () => {
          if (!scroller) return
          if (scroller.scrollTop !== pinTop) scroller.scrollTop = pinTop
          if (scroller.scrollLeft !== pinLeft) scroller.scrollLeft = pinLeft
        }
        scroller?.addEventListener('scroll', pinScroll, true)
        view.focus() // puts the DOM caret where the state says it is
        const sel = window.getSelection()
        // typing over a range would replace it — kick from a caret instead;
        // view.focus() below restores the real selection from PM state after
        if (sel && !sel.isCollapsed) sel.collapseToEnd()
        // the kick inserts exactly one space at the caret; snapshot the caret
        // text node so the scrub can restore it byte-identically
        const caretNode = sel?.anchorNode
        const caretText = caretNode instanceof Text ? caretNode : null
        const caretData = caretText?.data ?? null
        let sawKick = false
        let kickDown = false
        let kickPress = false
        // shield: PM's keymap must not see the kick, and the user's own keys
        // must not mutate the DOM while the observer is paused (the dwell
        // below makes that window user-noticeable on Windows) — swallow them.
        // Exactly one space keydown and the keypress it produces pass (the
        // kick); a later user space would land behind PM's back too
        const shield = (e: KeyboardEvent) => {
          let pass = false
          if (e.key === ' ') {
            if (e.type === 'keydown' && !kickDown) pass = kickDown = true
            else if (e.type === 'keypress' && !kickPress) pass = kickPress = true
          }
          if (pass && e.type === 'keypress') sawKick = true
          else if (!pass) e.preventDefault()
          e.stopPropagation()
        }
        document.addEventListener('keydown', shield, true)
        document.addEventListener('keypress', shield, true)
        const observer = (
          view as unknown as { domObserver: { stop: () => void; start: () => void } }
        ).domObserver
        observer.stop()
        const scrub = () => {
          // caret in a text node: remove exactly the one inserted character.
          // deleteData keeps the node's other spell markers alive — a whole
          // `data` reassignment would wipe the paragraph's fresh squiggles.
          if (caretText && caretData !== null && caretText.data !== caretData) {
            const now = caretText.data
            if (now.length === caretData.length + 1) {
              let i = 0
              while (i < caretData.length && now[i] === caretData[i]) i++
              caretText.deleteData(i, 1)
            }
            if (caretText.data !== caretData) caretText.data = caretData
            return
          }
          // the space landed elsewhere: Blink canonicalizes the insertion point
          // (end of a link, mark boundary, empty paragraph) into a sibling or
          // fresh text node — the caret it left sits right after the space
          if (!sawKick) return
          const after = window.getSelection()
          const node = after?.anchorNode
          if (!(node instanceof Text) || !after || after.anchorOffset === 0) return
          if (node === caretText) return
          const ch = node.data[after.anchorOffset - 1]
          if (ch === ' ' || ch === '\u00a0') node.deleteData(after.anchorOffset - 1, 1)
        }
        void window.desktop
          .respellKick()
          .catch(() => undefined)
          .then(async () => {
            try {
              // the IPC can resolve before the input pipeline delivers the
              // keystroke — wait for the shield to see it (or give up quietly:
              // a kick that never landed left nothing to scrub)
              const deadline = Date.now() + 800
              while (!sawKick && Date.now() < deadline) {
                await new Promise((r) => setTimeout(r, 30))
              }
              // Blink respells from an idle task and drops the answer if the text
              // changed meanwhile: scrub behind the task and the checker's reply
              if (sawKick) {
                await new Promise<void>((r) => {
                  if (typeof requestIdleCallback === 'function')
                    requestIdleCallback(() => r(), { timeout: 400 })
                  else setTimeout(r, 250)
                })
                await new Promise((r) => setTimeout(r, /win/i.test(navigator.platform) ? 900 : 150))
              }
              scrub()
            } finally {
              // a throw anywhere above must not leave the DOM observer stopped,
              // the capture-phase key shield swallowing the user's keystrokes,
              // or the busy latch held — any of those outlives the toggle and
              // reads as "spellcheck died for good, even new typing").
              observer.start()
              document.removeEventListener('keydown', shield, true)
              document.removeEventListener('keypress', shield, true)
              respellKickBusy.current = false
              spellDiag(`kick done landed=${sawKick}`)
              // the kick already focused the main editor and collapsed the
              // range, so the selection/textbox-routing/focus resync must run
              // whenever the editor is still live — including a still-mounted
              // re-toggle (cancelled) — or the user is left typing in the body
              // with a dropped selection. Only a detached/destroyed view (where
              // focus would throw) is skipped.
              if (dom.isConnected) {
                view.focus() // resync the DOM selection from PM state
                // ribbon toggle during textbox editing: keep the routing and
                // hand the keyboard back where it was — typing must not replace
                // the box
                if (sub) {
                  setActiveSubEditor(sub)
                  if (prev && prev !== dom) prev.focus({ preventScroll: true })
                }
                // unpin after the focus resync so its selection scroll is caught too
                pinScroll()
              }
              // always drop the pin listener, even on a torn-down round trip
              scroller?.removeEventListener('scroll', pinScroll, true)
            }
          })
      }
      rafId = requestAnimationFrame(runKick)
    }
    spellcheckWasOn.current = spellcheckActive
    respellNonceSeen.current = respellNonce
    return () => cancelRespellKick()
  }, [editor, spellcheck, spellcheckActive, respellNonce])

  // Ctrl/⌘+wheel zoom: Chromium delivers a trackpad pinch as ctrlKey wheel
  // events with small fractional deltas (kept continuous), while a mouse notch
  // steps ten percentage points. Must be non-passive to preventDefault.
  // The viewport point under the cursor is stashed for the zoom layout effect
  // below so the content there stays put (public #238); set only when the zoom
  // actually changes, else a clamped tick would leave a stale anchor behind.
  const zoomAnchorRef = useRef<{ vx: number; vy: number } | null>(null)
  useEffect(() => {
    const zoomWheel = createZoomWheelClassifier()
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return
      if (!(e.target as HTMLElement | null)?.closest?.('.editor-scroll')) return
      e.preventDefault()
      const intent = zoomWheel.feed(e, e.timeStamp)
      if (!intent) return
      const rect = scrollContainerRef.current?.getBoundingClientRect()
      const anchor = rect ? { vx: e.clientX - rect.left, vy: e.clientY - rect.top } : null
      setZoom((z) => {
        const next = nextDocsZoom(z, intent, e.deltaY)
        if (next !== z) zoomAnchorRef.current = anchor
        return next
      })
    }
    window.addEventListener('wheel', onWheel, { passive: false })
    return () => window.removeEventListener('wheel', onWheel)
  }, [])

  // Keep the document point under the anchor (cursor for wheel zoom, viewport
  // center otherwise) fixed across a zoom change. Runs after the zoomed layout
  // is committed, so the page box is read post-zoom and the pre-zoom box is
  // derived from it: CSS zoom scales the box linearly, the top edge is a fixed
  // scroller padding, and `.doc-zoom` is margin-auto centered until it overflows.
  // The pre-zoom scroll offsets come from the last scroll event: a zoom-out can
  // shrink the overflow and clamp them before this effect gets to read them.
  const scrollPosRef = useRef({ left: 0, top: 0 })
  const prevZoomRef = useRef(zoom)

  // ---- protection enforcement (Review > Protect Document) ----
  const editRestriction = protection?.enforced ? protection.edit : null
  /** modify password set but not entered: honor-system write lock, document read-only */
  const writeLocked = !!writeProtection?.hash && !modifyUnlocked
  /** body is read-only (readOnly/forms/comments restriction or write lock) */
  const isProtected =
    writeLocked ||
    editRestriction === 'readOnly' ||
    editRestriction === 'forms' ||
    editRestriction === 'comments'
  /** comments restriction: body read-only but adding comments stays allowed */
  const commentsAllowed = !writeLocked && editRestriction === 'comments'
  /** trackedChanges restriction: editing allowed, revision recording forced on */
  const trackChangesForced = !writeLocked && editRestriction === 'trackedChanges'

  // the trackedChanges restriction keeps the recorder on (the ribbon toggle is disabled)
  useEffect(() => {
    if (trackChangesForced && !trackChanges) setTrackChanges(true)
  }, [trackChangesForced, trackChanges])

  useEffect(() => (editor ? onEditableChange(editor, forceRender) : undefined), [editor])

  // Read Mode / Protect Document: the document becomes read-only; Esc leaves Read Mode.
  // A phased open stays editable (the append boundary is guarded by
  // StreamingTailGuardExtension; saves wait for the full content).
  useEffect(() => {
    if (!editor) return
    setEditorEditable(editor, !readMode && !isProtected)
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
    storage.author = authorName()
    return onAuthorNameChanged(() => {
      storage.author = authorName()
    })
  }, [editor, trackChanges])

  // forms protection locks the body but the form's own checkboxes stay
  // clickable; Read Mode and a still-loading document keep them read-only
  useEffect(() => {
    if (!editor) return
    editor.storage.checkboxToggle.formsFill =
      !writeLocked && editRestriction === 'forms' && !readMode && !docLoading
  }, [editor, writeLocked, editRestriction, readMode, docLoading])

  // outline view (View ▸ Outline View) folds: the plugin draws the ⊕/⊖ handles
  // and hides folded heading rows only while the view mode is outline; leaving
  // the view clears the collapsed set (view state, not document state)
  useEffect(() => {
    if (!editor) return
    editor.commands.setOutlineViewActive(viewMode === 'outline')
  }, [editor, viewMode])

  // story view (View ▸ Story View): the plugin draws the margin markers and
  // keeps typed text off the document's hidden formatting, only while the view
  // mode is story; Focus dims everything but the caret's sentence, inside the
  // view only
  useEffect(() => {
    if (!editor) return
    editor.commands.setStoryViewActive(viewMode === 'story')
    editor.commands.setStoryViewFocus(storyFocus && viewMode === 'story')
  }, [editor, viewMode, storyFocus])

  // entering Story View records the view it came from, so Leave Story (and Esc)
  // returns the document to where it was
  const viewBeforeStoryRef = useRef<ViewMode>('print')
  useEffect(() => {
    if (viewMode === 'story') {
      if (viewBeforeStoryRef.current !== 'story') setStoryReturnView(viewBeforeStoryRef.current)
      viewBeforeStoryRef.current = 'story'
    } else {
      viewBeforeStoryRef.current = viewMode
    }
  }, [viewMode])

  // Backstop for paths that bypass the view buttons (read-only opens land in read mode)
  useEffect(() => {
    if (readMode && viewMode !== 'print') setViewMode('print')
  }, [readMode, viewMode])

  // the preview reads the canvas as pages: Draft gives way in the render that mounts it
  const draftBreaksRef = useRef<HTMLDivElement>(null)
  const draftHeldRef = useRef(false)
  const viewModeRef = useRef(viewMode)
  viewModeRef.current = viewMode
  const showPreview = useCallback((show: boolean) => {
    if (show && viewModeRef.current === 'draft') {
      draftHeldRef.current = true
      setViewMode('print')
    }
    setShowPagePreview(show)
  }, [])
  useEffect(() => {
    if (showPagePreview || !draftHeldRef.current) return
    draftHeldRef.current = false
    if (viewMode === 'print') setViewMode('draft')
  }, [showPagePreview, viewMode])

  // window title follows the document, so the OS window list and Switch Window show file names
  useEffect(() => {
    document.title = doc ? doc.fileName : 'GenOffice Docs'
  }, [doc])

  useEffect(() => window.desktop.onTeardown?.(() => setTornDown(true)), [])

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

  // MCP bridge: let an external agent drive this visible editor. Commands arrive
  // from the shell main process and run against the live ctx (refs refresh per render).
  useEffect(() => {
    if (tornDown || !editor) return
    return installMcpBridge({ getCtx: () => fileCtxRef.current })
  }, [tornDown, editor])

  // Recompute the document-level line-height factor while editing:
  // docStyleCss decides it once at parse time, so typing CJK into a blank document
  // kept the Western factor until save/reopen (whole page ~8% shorter than the file).
  useEffect(() => {
    setLiveDocCjk(null)
    liveDocCjkRef.current = null
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
      liveDocCjkRef.current = has
      setLiveDocCjk(has)
    }
    let timer = 0
    // the whole-document scan only runs when the answer could have flipped:
    // CJK text arriving, content being deleted (the last of it could go), or
    // an unknown step kind. A rescan after every pause made long documents
    // lag (issue #2781); the fast paths answer from the touched text alone.
    const onUpdate = ({
      transaction,
      appendedTransactions,
    }: {
      transaction: Transaction
      appendedTransactions: Transaction[]
    }) => {
      window.clearTimeout(timer)
      const batch = [transaction, ...appendedTransactions]
      if (batch.some((tr) => textHasCjk(touchedText(tr) ?? ''))) {
        liveDocCjkRef.current = true
        setLiveDocCjk(true)
        return
      }
      if (liveDocCjkRef.current === true && !batch.some(deletedAnything)) return
      if (liveDocCjkRef.current === false && batch.every((tr) => touchedText(tr) !== null)) return
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
    const sync = () => setSplitHtml(scriptFontHtml(editor.getHTML()))
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
  const pendingMixedExportRef = useRef<PendingPdfExport | false>(false)
  // deferrals bump this so the effect re-arms even when the preview is already open
  const [pendingExportTick, setPendingExportTick] = useState(0)
  // Print dialog (Word-style preview + range); the pagination preview is its print source
  const [showPrintDialog, setShowPrintDialog] = useState(false)
  const showPrintDialogRef = useRef(false)
  showPrintDialogRef.current = showPrintDialog
  // Chromium is capturing the live .pv-page sheets: Escape / Close must not unmount them
  const [exportLocksPreview, setExportLocksPreview] = useState(false)
  // the dialog auto-opened the pagination preview: close it again with the dialog
  const printAutoOpenedPreviewRef = useRef(false)

  /** App state bundle for the extracted file actions (file-actions.ts); refreshed every render. */
  const fileCtxRef = useRef<FileActionContext>(null as unknown as FileActionContext)
  const seenAliasSeqRef = useRef(0)
  // style-pane previews patch the parsed styles; a font-settle rebuild must keep them
  const liveStylesRef = useRef<{ parsed: ParsedDocFull; styles: Map<string, StyleInfo> } | null>(
    null,
  )
  fileCtxRef.current = {
    editor,
    doc,
    resetCommentUndoStore,
    dirtyRef,
    saveInFlightRef,
    saveIncompleteRef,
    pendingMixedExportRef,
    bumpPendingExportTick: () => setPendingExportTick((n) => n + 1),
    printAutoOpenedPreviewRef,
    setShowPrintDialog,
    setStatus,
    setRecent,
    setShowAi,
    setDoc,
    setAiPanelKey,
    setDocCss,
    setDocLoading,
    setReadMode,
    setLargeDocSpellOff,
    setShowPagePreview: showPreview,
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
    propertyChanges,
    setPropertyChanges,
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
    pageCount: pageInfo.total,
    hfLinks,
    setHfLinks,
    titlePg,
    titlePgDirty,
    evenOddHf,
    evenOddHfDirty,
    setTitlePg,
    setTitlePgDirty,
    setEvenOddHf,
    setEvenOddHfDirty,
    mirrorMargins,
    mirrorMarginsDirty,
    setMirrorMargins,
    setMirrorMarginsDirty,
    autoHyphenation,
    autoHyphenationDirty,
    setAutoHyphenation,
    setAutoHyphenationDirty,
    hyphenationEdit,
    setHyphenationEdit,
    endManualHyphenation: () => {
      manualHyphenationRef.current = false
      setManualHyphenation(false)
    },
    defaultTabStopEdit,
    setDefaultTabStopEdit,
    gutterAtTopEdit,
    setGutterAtTopEdit,
    fontConversions,
    setFontConversions,
    compressPunctEdit,
    setCompressPunctEdit,
    trackMovesEdit,
    setTrackMovesEdit,
    strictKinsokuEdit,
    setStrictKinsokuEdit,
    mathPrEdit,
    setMathPrEdit,
    notePropsEdit,
    setNotePropsEdit,
    pgNumEdit,
    pgNumDirtySections,
    setPgNumEdit,
    setPgNumDirtySections,
    pendingNumbering,
    numberingDirty,
    setPendingNumbering,
    styleUpserts,
    setStyleUpserts,
    defaultFonts,
    setDefaultFonts,
    fontSettingsVersionRef,
    settleFontSettings: async () => {
      await fontSettingsPendingRef.current?.catch(() => {})
      return fileCtxRef.current
    },
    comments,
    commentsDirty,
    setComments,
    setCommentsDirty,
    setShowComments,
    setCommentPanelOpen,
    setCommentPopup,
    setCommentComposing,
    watermark,
    watermarkDirty,
    setWatermark,
    setWatermarkDirty,
    watermarkStyle,
    setWatermarkStyle,
    watermarkPicture,
    setWatermarkPicture,
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
    citationStyle,
    setCitationStyle,
    zoteroDocumentData,
    zoteroDocumentDataDirty,
    setZoteroDocumentData,
    setZoteroDocumentDataDirty,
    themeFonts,
    themeFontsDirty,
    themeColors,
    themeColorsDirty,
    setThemeFonts,
    setThemeFontsDirty,
    setThemeColors,
    setThemeColorsDirty,
    trackChanges,
    setTrackChanges,
    trackRevisionsDirty,
    setTrackRevisionsDirty,
    protection,
    protectionDirty,
    setProtection,
    setProtectionDirty,
    writeProtection,
    writeProtectionDirty,
    setWriteProtection,
    setWriteProtectionDirty,
    removePersonalInfo,
    removePersonalInfoDirty,
    setRemovePersonalInfo,
    setRemovePersonalInfoDirty,
    onWriteProtectionLoaded: (wp) => {
      setModifyUnlocked(!wp?.hash)
      setModifyPwdPrompt(wp?.hash ? { value: '', errorKey: '' } : null)
    },
    promptDocxPassword: (info) =>
      setDocPwdPrompt({ path: info.path, name: info.name, value: '', errorKey: '', busy: false }),
  }

  const loadFile = useCallback(async (result: OpenDocxResult) => {
    const outcome = await loadFileImpl(fileCtxRef.current, result)
    // a failed open with no document yet (boot, or the open that superseded the
    // boot open) must land on blank instead of the "Opening…" screen
    if (outcome === 'failed' && !fileCtxRef.current.doc) await newFileImpl(fileCtxRef.current)
    return outcome
  }, [])
  const resetFile = useCallback(() => newFileImpl(fileCtxRef.current), [])
  const commitOpenedFile = useCallback(
    async (result: OpenDocxResult) => {
      if (!result) return
      await loadFile(result)
    },
    [loadFile],
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
      void commitOpenedFile(result)
    })
    // With no pending file the window lands directly in the editor on a blank document
    // (the AI panel carries the generate-from-prompt flow). StrictMode runs the mount
    // effect twice but the pending queues can only be consumed once, so the consume
    // Promise lives in a ref and its result is processed only once.
    bootPendingRef.current ??= Promise.all([
      window.desktop.consumePendingOpenDocx(),
      // Still consume the one-shot new-blank flag so it doesn't leak into the next open
      window.desktop.consumeNewBlankDoc(),
      window.desktop.consumeAiDocContent(),
    ])
    void bootPendingRef.current
      .then(async ([pending, , aiContent]) => {
        if (bootHandledRef.current) return
        bootHandledRef.current = true
        // A failed open (corrupt file etc.) falls back to a blank document —
        // otherwise the tab shows "Opening…" forever with only a status-bar
        // line explaining why (github.com/genspark-ai/genoffice issue #102).
        // 'password': the prompt is up; its cancel path lands on blank instead.
        const outcome = pending ? await loadFile(pending) : 'canceled'
        if (outcome === 'canceled') await resetFile()
        if (aiContent && !pending) {
          // fileCtxRef refreshes per render: wait until resetFile's setDoc landed
          for (let i = 0; i < 100 && !fileCtxRef.current.doc; i++) {
            await new Promise((resolve) => setTimeout(resolve, 20))
          }
          await applyAiDocContentImpl(fileCtxRef.current, aiContent)
        }
      })
      // Open failures also land on a blank document, or the tab stays at "Opening…" forever
      .catch(() => {
        if (bootHandledRef.current) return
        bootHandledRef.current = true
        void resetFile().catch(() => {})
      })
    return unsubscribe
  }, [commitOpenedFile, editor, loadFile, resetFile])

  const openFile = useCallback(
    () =>
      runGuardedCandidate(
        () => window.desktop.confirmDocumentReplace(),
        () => window.desktop.openDocx(),
        commitOpenedFile,
        replacesDocument,
      ),
    [commitOpenedFile],
  )

  /** `fromGallery` names a template of the templates folder; without it the user picks a file */
  const newFromTemplate = useCallback(
    (fromGallery?: string) =>
      runGuardedCandidate(
        () => window.desktop.confirmDocumentReplace(),
        () => window.desktop.openTemplate(fromGallery),
        commitOpenedFile,
        replacesDocument,
      ),
    [commitOpenedFile],
  )

  /** new document from the built-in blank template (AI can then generate into it) */
  const newFile = useCallback(
    () => runGuardedDocumentAction(() => window.desktop.confirmDocumentReplace(), resetFile),
    [resetFile],
  )

  const openRecent = useCallback(
    (path: string) =>
      runGuardedCandidate(
        () => window.desktop.confirmDocumentReplace(),
        () => window.desktop.openDocxPath(path),
        commitOpenedFile,
        replacesDocument,
      ),
    [commitOpenedFile],
  )

  /** decrypt-and-open retry loop for the password prompt (wrong password stays in the dialog) */
  const submitDocPwd = async () => {
    if (!docPwdPrompt || docPwdPrompt.busy || !docPwdPrompt.value) return
    setDocPwdPrompt({ ...docPwdPrompt, busy: true, errorKey: '' })
    const res = await window.desktop.openDocxDecrypt(docPwdPrompt.path, docPwdPrompt.value)
    if (res.ok) {
      // Open deferred the replace guard for the encrypted file (see
      // replacesDocument); it runs once, here, now that a real document exists.
      // Declining keeps both the prompt and the current document.
      if (!(await window.desktop.confirmDocumentReplace())) {
        setDocPwdPrompt((current) => (current ? { ...current, busy: false } : current))
        return
      }
      setDocPwdPrompt(null)
      await commitOpenedFile(res.result)
      return
    }
    setDocPwdPrompt({
      ...docPwdPrompt,
      value: res.reason === 'wrong-password' ? '' : docPwdPrompt.value,
      busy: false,
      errorKey: res.reason === 'wrong-password' ? 'appDocPwdWrong' : 'appDocPwdUnsupported',
    })
  }

  const cancelDocPwd = () => {
    setDocPwdPrompt(null)
    // canceling a boot-time open leaves no document: land on blank, not "Opening…"
    if (!fileCtxRef.current.doc) void resetFile()
  }

  /** apply the diff the Protect Document dialog produced (undefined field = unchanged) */
  const applyProtectDialog = async (result: ProtectDialogResult) => {
    setShowProtectDialog(false)
    let changed = false
    if (result.openPassword !== undefined) {
      const cur = fileCtxRef.current.doc
      const res = await window.desktop.setDocPassword(cur?.filePath ?? null, result.openPassword)
      if (res.ok) {
        setDoc((d) => (d ? { ...d, encrypted: !!result.openPassword } : d))
        // the on-disk file only changes on the next save
        dirtyRef.current = true
        changed = true
      }
    }
    if (result.writeProtection !== undefined) {
      setWriteProtection(result.writeProtection)
      setWriteProtectionDirty(true)
      // the user set (or removed) the modify password themselves: never lock them out
      setModifyUnlocked(true)
      dirtyRef.current = true
      changed = true
    }
    if (result.protection !== undefined) {
      setProtection(result.protection)
      setProtectionDirty(true)
      dirtyRef.current = true
      changed = true
    }
    if (result.removePersonalInfo !== undefined) {
      setRemovePersonalInfo(result.removePersonalInfo)
      setRemovePersonalInfoDirty(true)
      dirtyRef.current = true
      changed = true
    }
    if (changed) setStatus(t('appProtectUpdated'))
  }

  /** modify-password prompt (write-protected document): verify, or fall back to read-only */
  const submitModifyPwd = async () => {
    if (!modifyPwdPrompt || !writeProtection) return
    if (await verifyProtectionPassword(modifyPwdPrompt.value, writeProtection)) {
      setModifyUnlocked(true)
      setModifyPwdPrompt(null)
    } else {
      setModifyPwdPrompt({ value: '', errorKey: 'appDocPwdWrong' })
    }
  }

  const save = useCallback(
    (saveAs: boolean, auto = false) => saveImpl(fileCtxRef.current, saveAs, auto),
    [],
  )

  // ---- Mailings: the data source and the merge ----

  /** the recipient list the user picked; null until they pick one */
  const [mergeSource, setMergeSource] = useState<MergeDataSource | null>(null)
  const [mergePanel, setMergePanel] = useState<{ mode: MergeOutputMode; pick: number } | null>(null)
  const [mergeRecordIndex, setMergeRecordIndex] = useState(0)
  const [mergeRange, setMergeRange] = useState<MergeRange>(MERGE_RANGE_ALL)
  /** rows of the list the user unticked, the conditions the rest must meet and their order */
  const [mergeExcluded, setMergeExcluded] = useState<ReadonlySet<number>>(() => new Set())
  const [mergeQuery, setMergeQuery] = useState<RecipientQuery>(EMPTY_RECIPIENT_QUERY)
  // a new list starts whole, at its first record; the old index may be past its end
  useEffect(() => {
    setMergeRecordIndex(0)
    setMergeRange(MERGE_RANGE_ALL)
    setMergeExcluded(new Set())
    setMergeQuery(EMPTY_RECIPIENT_QUERY)
  }, [mergeSource])
  /** the list as the preview and the merge take it */
  const mergeRecipients = useMemo(
    () =>
      mergeSource && {
        ...mergeSource,
        records: activeRecords(mergeSource.records, mergeExcluded, mergeQuery),
      },
    [mergeSource, mergeExcluded, mergeQuery],
  )
  const mergeRecipientCount = mergeRecipients?.records.length ?? 0
  // Merge All Windows has a window to bring back only while one stands outside the tab strip
  const [windowsDetached, setWindowsDetached] = useState(false)
  useEffect(() => {
    let live = true
    void window.desktop.detachedWindowCount?.().then((n) => live && setWindowsDetached(n > 0))
    const off = window.desktop.onDetachedWindowsChanged?.((n) => setWindowsDetached(n > 0))
    return () => {
      live = false
      off?.()
    }
  }, [])
  /** Preview Results: the merge fields show the current recipient instead of their names */
  const [mergePreview, setMergePreviewOn] = useState(false)
  const previewRecord = (mergePreview && mergeRecipients?.records[mergeRecordIndex]) || null
  useEffect(() => {
    if (editor && !editor.isDestroyed) setMergePreview(editor, previewRecord)
  }, [editor, previewRecord])
  useEffect(
    () => setMergeRecordIndex((index) => Math.max(Math.min(index, mergeRecipientCount - 1), 0)),
    [mergeRecipientCount],
  )

  /** open the file dialog and read the recipient list the user chose */
  const pickMergeDataSource = useCallback(async () => {
    const result = await window.desktop.pickMergeDataSource()
    if (!result) return
    if (!result.ok) {
      setStatus(result.error)
      return
    }
    setMergeSource(result.source)
    setStatus(
      t('appMergeSourceLoaded', {
        name: result.source.name,
        n: result.source.records.length,
      }),
    )
  }, [])

  /** a question an Ask or Fill-in rule puts during the merge, and where its answer goes */
  const [mergeQuestion, setMergeQuestion] = useState<{
    prompt: MergePrompt
    total: number
    answer: (text: string) => void
  } | null>(null)

  /** build the merged documents and write them where the user says */
  const mergeMailings = useCallback(
    async (request: MailMergeRequest): Promise<MailMergeResponse> => {
      try {
        const files = await buildMailMergeFiles(
          fileCtxRef.current,
          request,
          (prompt) =>
            new Promise<string>((answer) =>
              setMergeQuestion({
                prompt,
                total: (request.firstRecord ?? 1) + request.records.length - 1,
                answer,
              }),
            ),
        )
        if (files.length === 0) return { ok: false, error: t('appMergeNothingToWrite') }
        const written = await window.desktop.writeMergeOutput({
          mode: request.mode,
          files: files.map((file) => ({
            name: file.name,
            data: file.bytes.buffer.slice(
              file.bytes.byteOffset,
              file.bytes.byteOffset + file.bytes.byteLength,
            ) as ArrayBuffer,
          })),
        })
        if (written.ok) {
          setStatus(t('appMergeWritten', { n: written.count ?? files.length }))
        }
        return written
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
    [],
  )

  /**
   * The reader asked the model for a name on a document that already has one.
   *
   * The stem is a proposal for the Save As dialog, never a rename: the dialog
   * is where the reader confirms, and moving a file that may be open elsewhere
   * or synced is not ours to do quietly.
   */
  const nameWithAi = useCallback(async () => {
    const proposed = await suggestedSaveAsName(fileCtxRef.current)
    if (!proposed) return
    await saveImpl(fileCtxRef.current, true, false, proposed)
  }, [])

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

  // an unsaved break undone or deleted takes its modelled section with it
  useEffect(() => {
    const live = editor && dropRemovedPendingSections(editor.state.doc, sections)
    if (!live) return
    const gone = new Set(live.dropped)
    const shift = (d: number[]) =>
      d.filter((i) => !gone.has(i)).map((i) => i - live.dropped.filter((g) => g < i).length)
    setSections(live.sections)
    setSectionsDirty(shift)
    setPgNumDirtySections(shift)
    // docVersion: the doc identity is what changes, and it is not a React value
  }, [editor, sections, docVersion])

  /** Layout-tab settings per section; one behind an unsaved break has them written into its break paragraph */
  const editSectionSettings = (
    settingsOf: (s: SectionInfo, i: number) => SectionSettings | null,
  ) => {
    const patched = editor
      ? rewritePendingSectPrs(editor, sections, (s, i) => {
          const next = settingsOf(s, i)
          return next && applySectionSettings(s.sectPrXml, next)
        })
      : sections
    setSections((prev) =>
      prev.map((s, i) => {
        const next = settingsOf(s, i)
        return next ? { ...s, settings: next, sectPrXml: patched[i]?.sectPrXml ?? s.sectPrXml } : s
      }),
    )
  }

  /**
   * Insert a section break: the new break paragraph takes a copy of the current
   * section's sectPr (content before the break keeps the original
   * section's settings); "continuous" is written to the following section's (the
   * original current section sectPr's) w:type.
   */
  /**
   * A section's sectPr as save would write it now: unsaved Layout-tab / AI edits
   * applied on top of the parsed XML (a break copies this, not the stale file bytes).
   */
  const effectiveSectPrXml = (idx: number): string | null => {
    const sec = sections[idx]
    if (!sec) return null
    if (sec.pendingBreak) return sec.sectPrXml
    const isLast = idx === sections.length - 1
    let xml = sec.sectPrXml
    if (isLast ? sectionDirty : sectionsDirty.includes(idx)) {
      xml = applySectionSettings(xml, (isLast && section) || sec.settings)
      xml = applyTitlePg(xml, isLast ? titlePg : sec.titlePg)
    } else if (isLast && titlePgDirty) {
      xml = applyTitlePg(xml, titlePg)
    }
    if (isLast ? pgNumEdit !== null : pgNumDirtySections.includes(idx)) {
      const pg = isLast
        ? pgNumEdit
        : {
            fmt: sec.pageNumberFmt,
            start: sec.pageNumberStart,
            chapter: sec.pageNumberChapter ?? null,
          }
      xml = applyPageNumType(xml, pg?.fmt, pg?.start, pg?.chapter)
    }
    return xml
  }

  const insertSectionBreak = useCallback(
    (
      type: SectionInfo['startType'],
      afterBlockIndex?: number,
      // the AI tool's own section mirror (fresh between renders) and the sectPr copy it chose
      ai?: { sectPr: string; sections: SectionInfo[] },
    ): boolean => {
      if (!editor || !doc) return false
      const pmDoc = editor.state.doc
      const live = ai?.sections ?? sections
      // an explicit block (AI tool) targets that block's section instead of the cursor's
      const targetSection =
        afterBlockIndex === undefined
          ? activeSection
          : sectionIndexAtBlock(pmDoc, live, Math.max(afterBlockIndex, 0))
      const target = live[targetSection]
      const sectPrCopy =
        ai?.sectPr ||
        effectiveSectPrXml(targetSection) ||
        applySectionSettings(
          '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>',
          section ?? sections[0]?.settings ?? DEFAULT_SECTION,
        )
      const genXml = `<w:p><w:pPr>${sectPrCopy}</w:pPr></w:p>`
      // an unsaved section's sectPr lives in its own break paragraph: write the
      // type there before the copy goes in (the copy carries the same XML)
      const pendingXml = target?.pendingBreak ? applySectionStartType(target.sectPrXml, type) : null
      if (pendingXml) patchPendingSectPr(editor, live, targetSection, pendingXml)
      const { $head } = editor.state.selection
      let pos = $head.depth > 0 ? $head.after(1) : $head.pos
      if (afterBlockIndex !== undefined) {
        pos = 0
        for (let i = 0; i <= afterBlockIndex && i < pmDoc.childCount; i++)
          pos += pmDoc.child(i).nodeSize
      }
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
      // modelled now, not when the save re-reads the file: a save that races
      // further edits skips the reload, and a whole-document page setup would
      // miss the section. A ribbon break in a never-saved document stays out
      // of the model: it takes effect on the first save.
      const breakAfter =
        afterBlockIndex ?? ($head.depth > 0 ? $head.index(0) : pmDoc.childCount - 1)
      const created =
        target && (ai || doc.filePath)
          ? pendingSectionAt(pmDoc, target, breakAfter, sectPrCopy)
          : null
      // the section after the break is terminated by the "original current section sectPr",
      // whose w:type decides how the new section starts: always write the chosen type back
      // (including a nextPage reset, so a leftover continuous from the original section doesn't linger)
      if (pendingXml) {
        setSections((prev) =>
          prev.map((s, i) =>
            i === targetSection ? { ...s, startType: type, sectPrXml: pendingXml } : s,
          ),
        )
      } else if (live.length === 0 || targetSection === live.length - 1) {
        setTrailingStartType(type)
      } else {
        setSections((prev) =>
          prev.map((s, i) => (i === targetSection ? { ...s, startType: type } : s)),
        )
        setSectionsDirty((d) => (d.includes(targetSection) ? d : [...d, targetSection]))
      }
      if (created) {
        setSections((prev) => {
          const next = [...prev]
          next.splice(targetSection, 0, created)
          if (next[targetSection + 1]) {
            next[targetSection + 1] = {
              ...next[targetSection + 1]!,
              firstBlockIndex: created.lastBlockIndex + 1,
              startType: type,
            }
          }
          return next
        })
        const shift = (d: number[]) => d.map((i) => (i >= targetSection ? i + 1 : i))
        setSectionsDirty(shift)
        setPgNumDirtySections(shift)
      }
      const labels: Record<SectionInfo['startType'], string> = {
        nextPage: t('appBreakNextPage'),
        continuous: t('appBreakContinuous'),
        evenPage: t('appBreakEvenPage'),
        oddPage: t('appBreakOddPage'),
        // parse-only start type (single-column: acts like next page); the UI never inserts it
        nextColumn: t('appBreakNextPage'),
      }
      if (doc.filePath) {
        pendingSectionSaveRef.current = true
        setStatus(t('appSectionBreakInserted', { type: labels[type] }))
      } else {
        setStatus(t('appSectionBreakPending'))
      }
      return created !== null
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps -- effectiveSectPrXml is rebuilt per render from the state listed here
    [
      editor,
      doc,
      sections,
      activeSection,
      section,
      sectionDirty,
      sectionsDirty,
      titlePg,
      titlePgDirty,
      pgNumEdit,
      pgNumDirtySections,
    ],
  )
  // "This point forward" needs the section after the new break modelled at once: a saved
  // document, and a caret section that is not itself waiting behind an unsaved break
  const canApplyForward =
    !!doc?.filePath && sections.length > 0 && !sections[activeSection]?.pendingBreak
  const insertSectionBreakRef = useRef(insertSectionBreak)
  insertSectionBreakRef.current = insertSectionBreak

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
    linkStyle: (styleId, numId, ilvl) =>
      setStyleUpserts((prev) => ({
        ...prev,
        [styleId]: {
          ...(prev[styleId] ?? { styleId }),
          styleId,
          pPr: { ...(prev[styleId]?.pPr ?? {}), numPr: { numId, ilvl } },
        },
      })),
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

  /** which list dialog is open (ribbon galleries and the right-click menu open them) */
  const [listDialog, setListDialog] = useState<ListDialogKind | null>(null)
  /** the library card the open define dialog customizes */
  const [listCard, setListCard] = useState<number | null>(null)
  /** the list style Bullets and Numbering opens on: one just added */
  const [addedListStyle, setAddedListStyle] = useState<string | undefined>()
  const [listRecentsEpoch, setListRecentsEpoch] = useState(0)

  // Story View: Esc leaves, back to the view the document came in from. An
  // Escape meant for something else never claims the view (storyEscLeavesView):
  // the context menu, the comment popup and the dialogs close on their own
  // Escape instead, and so does an Escape inside the find panel.
  useEffect(() => {
    if (viewMode !== 'story') return
    const onKey = (e: KeyboardEvent) => {
      if (!storyEscLeavesView(e, !!(showLinkModal || listDialog || ctxMenu || commentPopup))) return
      setViewMode(storyReturnView)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [viewMode, showLinkModal, listDialog, ctxMenu, commentPopup, storyReturnView])
  const recentLists = useMemo(
    () => ({
      bullets: recentListPresets('bullets'),
      numbers: recentListPresets('numbers'),
      multi: recentListPresets('multi'),
    }),
    // recomputed after every pick
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [listRecentsEpoch],
  )
  const documentLists = useMemo(() => {
    if (!doc) return undefined
    const defs = new Map(doc.parsed.numbering)
    for (const d of pendingNumbering.newDefs) {
      const overlay = (
        editor?.storage.listNumbering as { defs?: Map<string, NumberingDef> } | undefined
      )?.defs?.get(d.numId)
      if (overlay) defs.set(d.numId, overlay)
    }
    return documentListPresets(defs.values())
  }, [doc, pendingNumbering.newDefs, editor])
  const listStyles = useMemo(
    () =>
      listStyleEntries(
        doc?.parsed.listStyles,
        pendingNumbering.styleDefs,
        pendingNumbering.removedStyles,
      ),
    [doc, pendingNumbering.styleDefs, pendingNumbering.removedStyles],
  )
  const applyDefinedList = useCallback(
    (levels: CustomNumberingLevel[], kind: 'bullets' | 'numbers' | 'multi') => {
      if (!editor) return
      const numId = createCustomListDefImpl(numberingCtxRef.current, levels)
      if (!numId) return
      rememberListPreset(kind, levels)
      setListRecentsEpoch((n) => n + 1)
      const ilvl = editor.isActive('docListItem')
        ? Number(editor.getAttributes('docListItem').ilvl) || 0
        : 0
      editor
        .chain()
        .focus()
        .setNode('docListItem', {
          kind: levels[0]?.numFmt === 'bullet' ? 'bullet' : 'ordered',
          numId,
          ilvl,
        })
        .run()
      setListDialog(null)
    },
    [editor],
  )
  const changeListLevel = useCallback(
    (ilvl: number) => {
      if (editor) changeListLevelImpl(editor, ilvl)
    },
    [editor],
  )
  const listTarget = listDialog ? currentListTargetImpl(numberingCtxRef.current) : null

  // typed "- " / "1. " lists get real numbering definitions, like the ribbon buttons
  useEffect(() => {
    if (!editor) return
    const storage = editor.storage.autocorrect
    storage.hooks = {
      bulletNumId: () => allocateListNumId('bullet'),
      numberedNumId: (numFmt, pattern, start) => {
        const levels = numberPresetLevels(numFmt, pattern)
        if (start > 1) levels[0] = { ...levels[0], start }
        return createCustomListDef(levels)
      },
      insertTable: (cols) => insertTableAt(editor, 1, cols, true),
      autoText: autoTextEntries,
      insertAutoText: (entry) => insertAutoText(editor, entry),
      openOptions: () => {
        setAutoTextList(autoTextEntries())
        setShowAutoCorrect(true)
      },
    }
    return () => {
      storage.hooks = null
    }
  }, [editor, allocateListNumId, createCustomListDef])

  // ---- Inline fields: insertion and F9 update (cached results recomputed locally) ----

  const fieldValue = useCallback(
    (instr: string, lang?: string): string => {
      const kw = instr.trim().split(/\s+/)[0]?.toUpperCase()
      const now = new Date()
      const picture = datePictureOf(instr)
      const dated = (text: string) => (wantsFullWidth(instr) ? toFullWidth(text) : text)
      switch (kw) {
        case 'DATE':
        case 'CREATEDATE':
        case 'SAVEDATE':
          return dated(
            picture
              ? formatDatePicture(now, picture, lang)
              : now.toLocaleDateString(lang ?? 'zh-CN'),
          )
        case 'TIME':
          return dated(
            picture
              ? formatDatePicture(now, picture, lang)
              : now.toLocaleTimeString(lang ?? 'zh-CN', { hour: '2-digit', minute: '2-digit' }),
          )
        case 'NUMPAGES':
          return formatFieldNumber(instr, pageInfo.total)
        case 'PAGE':
          return formatFieldNumber(instr, pageInfo.current)
        case 'FILENAME':
          return formatFieldText(
            instr,
            (/\\p\b/i.test(instr) ? doc?.filePath : null) ?? doc?.fileName ?? '',
          )
        case 'MERGEFIELD': {
          // a merge field shows its name in guillemets until the merge fills it
          const name = mergeFieldName(instr)
          return name ? placeholderText(name) : ''
        }
        default:
          return ''
      }
    },
    [pageInfo, doc],
  )

  const insertField = useCallback(
    (instr: string, lang?: string) => {
      if (!editor) return
      // a table of contents entry shows nothing: it is a marker, as an index entry is
      const entry = /^\s*TC\s+(\S[\s\S]*?)\s*$/i.exec(instr)?.[1]
      if (entry && !getActiveSubEditor()) {
        insertFieldMarks(editor, [
          { at: editor.state.selection.to, attrs: { term: '', tc: entry } },
        ])
        setStatus(t('appFieldInserted', { instr }))
        return
      }
      const value = fieldValue(instr, lang) || ' '
      // a caret in a header, footer or text box takes the field, as it takes typing
      const sub = getActiveSubEditor()
      const target = sub?.isEditable && sub.schema.marks.instrField ? sub : editor
      // a header or footer keeps page numbers live: its own field node, not a cached result
      const live = liveHfFieldOf(instr)
      if (live && target.schema.nodes.hfField) {
        insertHfField(target, live)
        setStatus(t('appFieldInserted', { instr }))
        return
      }
      target
        .chain()
        .focus()
        .insertContent({
          type: 'text',
          text: value,
          marks: [
            {
              type: 'instrField',
              attrs: RULE_FIELD_INSTR.test(instr)
                ? { instr, fieldId: nextFieldId(target.state.doc), fieldPart: 'single' }
                : { instr },
            },
            // Word names the months and days of a date field in the language of its text
            ...(lang && target.schema.marks.docTextStyle
              ? [
                  {
                    type: 'docTextStyle',
                    attrs: { lang: JSON.stringify({ [languageSlot(lang)]: lang }) },
                  },
                ]
              : []),
          ],
        })
        .unsetMark('instrField')
        .run()
      setStatus(t('appFieldInserted', { instr }))
    },
    [editor, fieldValue],
  )

  /** F9: recompute the inline field caches (PAGE/NUMPAGES/date-time/file name, caption numbers, formulas); TOC has its own Update Table */
  const updateFields = useCallback(() => {
    if (!editor) return
    const { state, view } = editor
    const jobs: Array<{ from: number; to: number; text: string; marks: readonly PmMark[] }> = []
    let changed = 0
    state.doc.descendants((node, pos) => {
      if (!node.isText) return
      const mark = node.marks.find((m) => m.type.name === 'instrField')
      if (!mark) return
      const slots = parseLangAttr(
        node.marks.find((m) => m.type.name === 'docTextStyle')?.attrs.lang,
      )
      const next = fieldValue(
        String(mark.attrs.instr),
        slots?.val ?? slots?.eastAsia ?? slots?.bidi,
      )
      // a page field keeps the mark it has: its page here is the one in view, not its own
      const paged = PAGE_FIELD_INSTR.test(String(mark.attrs.instr))
      if (!next || (next === node.text && (paged || mark.attrs.dirty !== true))) return
      if (next !== node.text) changed++
      // any other field recomputed here carries its result: nothing is left for Word to update
      const marks = paged
        ? node.marks
        : node.marks.map((m) =>
            m === mark ? mark.type.create({ ...mark.attrs, dirty: false }) : m,
          )
      jobs.push({ from: pos, to: pos + node.nodeSize, text: next, marks })
    })
    let tr = state.tr
    for (const j of jobs.sort((a, b) => b.from - a.from)) {
      tr = tr.replaceWith(j.from, j.to, state.schema.text(j.text, [...j.marks]))
    }
    const noteRefs = refreshNoteRefFields(tr, {
      above: t('ribbonCrossRefAbove'),
      below: t('ribbonCrossRefBelow'),
    })
    const blocks = doc?.parsed.blocks ?? EMPTY_BLOCKS
    const updated =
      changed +
      noteRefs +
      refreshTableFormulas(tr) +
      renumberSeqFields(tr, blocks, {
        skip: unreliableSeqLabelsOf(blocks, state.schema),
        dropDirty: true,
      })
    if (!tr.docChanged) {
      setStatus(t('appNoFieldsToUpdate'))
      return
    }
    view.dispatch(tr)
    setStatus(updated === 0 ? t('appNoFieldsToUpdate') : t('appFieldsUpdated', { n: updated }))
  }, [editor, fieldValue, doc])

  // editorRef: lets the handlePaste closure (the useEditor config exists before the instance) reach the instance
  useEffect(() => {
    editorRef.current = editor
    zoteroControllerRef.current = null
  }, [editor])

  useEffect(
    () =>
      window.desktop.onZoteroRequest(async (request) => {
        try {
          const activeEditor = editorRef.current
          if (!activeEditor) throw new Error('No active GenOffice document')
          const controller =
            zoteroControllerRef.current ??
            new ZoteroDocumentController(activeEditor, {
              get: () => zoteroDocumentDataRef.current,
              set: (value) => {
                setZoteroDocumentData(value)
                setZoteroDocumentDataDirty(true)
              },
            })
          zoteroControllerRef.current = controller
          const result = await controller.handle(request)
          window.desktop.respondToZotero({ requestId: request.requestId, ok: true, result })
        } catch (error) {
          window.desktop.respondToZotero({
            requestId: request.requestId,
            ok: false,
            error: error instanceof Error ? error.message : String(error),
          })
        }
      }),
    [setZoteroDocumentData],
  )

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
      if (tr.docChanged) {
        const before = editor.state.doc
        editor.view.dispatch(tr)
        carryDocSeen(editor, before)
      }
    }
    // a numId-less list item can only appear inside a range a step touched,
    // so the whole-document walk runs only when that happened. Walking every
    // keystroke made long documents lag (issue #2781).
    const onFillUpdate = ({
      transaction,
      appendedTransactions,
    }: {
      transaction: Transaction
      appendedTransactions: Transaction[]
    }) => {
      const batch = [transaction, ...appendedTransactions]
      if (batch.some((tr) => touchesNode(tr, (node) => node.type.name === 'docListItem'))) fill()
    }
    editor.on('update', onFillUpdate)
    return () => {
      editor.off('update', onFillUpdate)
    }
  }, [editor, allocateListNumId])

  /** Open the page-number-format dialog (initial value = the cursor section's pgNumType) */
  const openPgNumModal = useCallback(() => {
    const sec = sections[Math.min(activeSection, sections.length - 1)]
    setPgNumModal({
      fmt: sec?.pageNumberFmt ?? 'decimal',
      start: sec?.pageNumberStart !== undefined ? String(sec.pageNumberStart) : '1',
      restart: sec?.pageNumberStart !== undefined,
      chapStyle: sec?.pageNumberChapter ? String(sec.pageNumberChapter.style) : '',
      chapSep: sec?.pageNumberChapter?.sep ?? 'hyphen',
    })
  }, [sections, activeSection])

  /** Apply the page-number format to the cursor's section: final section via SaveOptions.pgNumType, others rewrite their sectPr */
  const applyPgNumFormat = useCallback(() => {
    if (!pgNumModal) return
    const fmt = pgNumModal.fmt === 'decimal' ? undefined : pgNumModal.fmt
    const start =
      !pgNumModal.restart || pgNumModal.start.trim() === ''
        ? undefined
        : Math.max(0, parseInt(pgNumModal.start, 10) || 0)
    const chapter = pgNumModal.chapStyle
      ? { style: Number(pgNumModal.chapStyle), sep: pgNumModal.chapSep }
      : null
    const idx = sections.length > 0 ? Math.min(activeSection, sections.length - 1) : -1
    if (idx >= 0) {
      const patched = editor
        ? rewritePendingSectPrs(editor, sections, (s, i) =>
            i === idx ? applyPageNumType(s.sectPrXml, fmt, start, chapter) : null,
          )
        : sections
      setSections((prev) =>
        prev.map((s, i) =>
          i === idx
            ? {
                ...s,
                pageNumberFmt: fmt,
                pageNumberStart: start,
                pageNumberChapter: chapter ?? undefined,
                sectPrXml: patched[i]?.sectPrXml ?? s.sectPrXml,
              }
            : s,
        ),
      )
    }
    if (idx < 0 || idx === sections.length - 1) {
      setPgNumEdit({
        ...(fmt !== undefined ? { fmt } : {}),
        ...(start !== undefined ? { start } : {}),
        chapter,
      })
    } else {
      setPgNumDirtySections((d) => (d.includes(idx) ? d : [...d, idx]))
    }
    setPgNumModal(null)
    setStatus(t('appPgNumFormatSet'))
  }, [pgNumModal, sections, activeSection, editor])

  const exportPdf = useCallback(
    (outPath?: string) => exportPdfImpl(fileCtxRef.current, outPath),
    [],
  )
  const exportHtml = useCallback(
    (outPath?: string) => exportHtmlImpl(fileCtxRef.current, outPath),
    [],
  )
  const exportMarkdown = useCallback(
    (outPath?: string) => exportMarkdownImpl(fileCtxRef.current, outPath),
    [],
  )
  const exportImages = useCallback(() => exportImagesImpl(fileCtxRef.current), [])
  const printDoc = useCallback(() => printDocImpl(fileCtxRef.current), [])

  // for real-device verification: trigger export directly via CDP (same as __pageDebug)
  useEffect(() => {
    ;(window as unknown as Record<string, unknown>).__exportPdf = exportPdf
    ;(window as unknown as Record<string, unknown>).__exportImages = exportImages
    ;(window as unknown as Record<string, unknown>).__openPagePreview = () => showPreview(true)
    ;(window as unknown as Record<string, unknown>).__setRevisionDisplay = (
      mode: RevisionDisplayMode,
    ) => setRevisionDisplay(mode)
  }, [exportPdf, exportImages, showPreview])

  // Headless export mode (--headless-export): this renderer lives in a hidden
  // window whose only job is to run the File menu's PDF or HTML export against
  // a path the CLI chose, then report back so the main process can quit.
  const headlessExportStartedRef = useRef(false)
  useEffect(() => {
    if (headlessExportStartedRef.current) return
    headlessExportStartedRef.current = true
    void (async () => {
      const target = await window.desktop.consumeHeadlessExport()
      if (!target) return
      const report = await runHeadlessDocumentExport(
        target.outPath,
        // A failed open falls back to an untitled blank document (filePath
        // null); exporting that would hand the CLI a blank PDF and call it
        // success, so only a document that came from disk counts as ready.
        () => {
          const doc = fileCtxRef.current.doc
          return { opened: typeof doc?.filePath === 'string', failed: doc?.filePath === null }
        },
        target.format === 'html' ? exportHtml : target.format === 'md' ? exportMarkdown : exportPdf,
      )
      window.desktop.headlessExportDone(report)
    })()
  }, [exportPdf, exportHtml, exportMarkdown])

  useEffect(() => {
    const pending = pendingMixedExportRef.current
    if (pending === false) return
    if (!showPagePreview) {
      // preview dismissed while the export was parked: settle as canceled
      pendingMixedExportRef.current = false
      pending.resolve(false)
      setStatus(t('appExportPdfCanceled'))
      return
    }
    // The ref keeps holding `pending` while the wait runs; every settle path
    // must claim it (identity check + clear in one sync segment), so a loop
    // superseded by a newer deferral dies silently instead of double-settling
    // or exporting concurrently.
    const claim = () => {
      if (pendingMixedExportRef.current !== pending) return false
      pendingMixedExportRef.current = false
      return true
    }
    const timer = window.setTimeout(async () => {
      // preview pages mount asynchronously; the export needs at least one.
      // Never re-enter exportPdf without a page: the pageless paths would
      // defer again (unbounded loop) or repeat an already-failed direct print.
      for (let i = 0; i < 40 && !document.querySelector('.pv-page'); i++) {
        await new Promise((r) => window.setTimeout(r, 250))
        if (pendingMixedExportRef.current !== pending) return
        if (!document.querySelector('.pagination-preview')) {
          if (!claim()) return
          pending.resolve(false)
          setStatus(t('appExportPdfCanceled'))
          return
        }
      }
      if (!claim()) return
      if (!document.querySelector('.pv-page')) {
        pending.resolve(false)
        setStatus(t('appExportPdfFailed', { error: 'pagination preview produced no pages' }))
        return
      }
      setExportLocksPreview(true)
      const ok = await exportPdf(pending.outPath).catch(() => false)
      setExportLocksPreview(false)
      // a Print the user started meanwhile owns the preview now
      if (pending.closePreview && !showPrintDialogRef.current) setShowPagePreview(false)
      pending.resolve(ok)
    }, 1200)
    return () => window.clearTimeout(timer)
  }, [showPagePreview, pendingExportTick, exportPdf])

  // Highlight Changes > in printed document: the print sheet prints the live pages, so
  // the Markup view switches while the sheet is open and comes back when it closes
  const [printHighlight, setPrintHighlight] = useState<boolean | null>(null)
  const printRestoreDisplayRef = useRef<RevisionDisplayMode | null>(null)
  useEffect(() => {
    if (showPrintDialog) {
      // a rerun while the sheet is open still restores the view the screen had before it
      const screen = printRestoreDisplayRef.current ?? revisionDisplayRef.current
      const target = displayForPrint(screen, printHighlight)
      if (target === revisionDisplayRef.current) return
      printRestoreDisplayRef.current = screen
      setRevisionDisplay(target)
    } else if (printRestoreDisplayRef.current) {
      setRevisionDisplay(printRestoreDisplayRef.current)
      printRestoreDisplayRef.current = null
    }
  }, [showPrintDialog, printHighlight])

  const closePrintDialog = useCallback(() => {
    setShowPrintDialog(false)
    if (printAutoOpenedPreviewRef.current) {
      printAutoOpenedPreviewRef.current = false
      setShowPagePreview(false)
    }
  }, [])

  // ---- References: footnotes / endnotes ----

  const [notePrompt, setNotePrompt] = useState<NotePrompt | null>(null)

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
    // a getter into the live mirror, not a render-time reference: AI tool
    // calls run several review actions between renders, and each one must
    // see the previous write (setComments replaces the array)
    get comments() {
      return commentsLiveRef.current
    },
    setComments,
    setCommentsDirty,
    setCommentComposing,
    setShowComments,
    setCommentFocus,
    setInkAnnotations,
    setInksDirty,
    currentDocBytes: () => buildDocBytesImpl(fileCtxRef.current),
    get revisionFilter() {
      return revisionFilterRef.current
    },
    showRevision: (target) => showRevision(target),
    showAllMarkup: () => setRevisionDisplay('all'),
    storyRevisions: () => storyRevisionsRef.current,
    applyStoryRevisions: (targets, mode, all) => applyStoryRevisions(targets, mode, all),
    currentStoryRevision: () =>
      storyRevisionsRef.current.find((r) => r.story.id === lastStoryRef.current) ?? null,
  }

  const insertNote = useCallback((kind: 'footnote' | 'endnote') => {
    setNotePrompt({ kind })
  }, [])

  /** what the dialog opens on: the caret's section's numbering, and the document's placement (the page has one) */
  /** the section the caret's notes are numbered in (a new paragraph after a saved break is in the next one) */
  const noteSectionAtCaret = (): number => {
    const head = editor?.state.selection.$head
    const at =
      editor && head && head.depth > 0
        ? noteBlockSections(editor.state.doc, sections)[head.index(0)]
        : undefined
    return Math.min(at ?? activeSection, Math.max(sections.length - 1, 0))
  }
  const activeNoteProps = (kind: NoteKind): NoteProps => {
    const { pos: _own, ...numbering } = sectionNoteProps[noteSectionAtCaret()]?.[kind] ?? {}
    const pos = noteSettings[kind].pos
    return { ...numbering, ...(pos ? { pos } : {}) }
  }
  const renumberNotes = (props: typeof sectionNoteProps) => {
    if (!editor) return
    setNoteProps(props, sections)
    const tr = editor.state.tr
    renumber(tr, 'footnote')
    renumber(tr, 'endnote')
    // the numbers follow the setting, which is an undo step of its own
    if (tr.docChanged) editor.view.dispatch(tr.setMeta('addToHistory', false))
  }
  const noteNumberingSigOf = (props: typeof sectionNoteProps, list: readonly SectionInfo[]) =>
    JSON.stringify([
      props.map((p) =>
        [p.footnote, p.endnote].map((k) => [
          k.numFmt,
          k.numStart ?? 1,
          k.numRestart === 'eachSect',
        ]),
      ),
      list.map((s) => [s.lastBlockIndex, s.pendingBreak ? 1 : 0]),
    ])
  const noteNumberingSig = noteNumberingSigOf(sectionNoteProps, sections)
  /** the numbering the marks were last counted under; a file opens counted by the engine */
  const noteNumberingRef = useRef<{ parsed: unknown; sig: string } | null>(null)
  useEffect(() => {
    if (!editor || !doc) return
    const counted = noteNumberingRef.current
    noteNumberingRef.current = { parsed: doc.parsed, sig: noteNumberingSig }
    if (counted?.parsed !== doc.parsed) {
      const opened = readSections(doc.parsed)
      const sig = noteNumberingSigOf(
        sectionNotePropsOf(
          doc.parsed.settingsNoteProps,
          opened.map((s) => s.settings),
        ),
        opened,
      )
      if (sig === noteNumberingSig) return
    } else if (counted.sig === noteNumberingSig) return
    // a section added, removed or given other note properties: its marks count again
    renumberNotes(sectionNoteProps)
    setNoteRenumbered((n) => n + 1)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editor, doc, noteNumberingSig])

  // the Footnote and Endnote dialog's Apply: the marks renumber and redraw now, the file takes it on save.
  // A section's own properties win over the document's, so a document-wide change goes into every section too.
  const applyNoteSettings = (delta: NoteSettings, scope: 'document' | 'section') => {
    if (!editor) return
    const last = Math.max(sections.length - 1, 0)
    const at = noteSectionAtCaret()
    const hit = (i: number) => scope === 'document' || i === at
    const own = (s: SectionSettings): SectionSettings => ({
      ...s,
      ...(Object.keys(delta.footnote).length > 0
        ? { footnotePr: { ...s.footnotePr, ...delta.footnote } }
        : {}),
      ...(Object.keys(delta.endnote).length > 0
        ? { endnotePr: { ...s.endnotePr, ...delta.endnote } }
        : {}),
    })
    editSectionSettings((s, i) => (hit(i) ? own(s.settings) : null))
    if (section && hit(last)) {
      setSection(own(section))
      setSectionDirty(true)
    }
    const earlier = sections.map((_, i) => i).filter((i) => i < last && hit(i))
    if (earlier.length > 0) setSectionsDirty((d) => [...new Set([...d, ...earlier])])
    if (scope === 'document')
      setNotePropsEdit({
        footnote: { ...notePropsEdit?.footnote, ...delta.footnote },
        endnote: { ...notePropsEdit?.endnote, ...delta.endnote },
      })
    const next = sectionNoteProps.map((p, i) =>
      hit(i)
        ? {
            footnote: { ...p.footnote, ...delta.footnote },
            endnote: { ...p.endnote, ...delta.endnote },
          }
        : p,
    )
    renumberNotes(next)
    noteNumberingRef.current = doc && {
      parsed: doc.parsed,
      sig: noteNumberingSigOf(next, sections),
    }
  }

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
  /**
   * Capture the range now, while the menu still knows what was selected: the
   * dialog takes seconds to answer and a click elsewhere would otherwise move
   * the span the reader asked to withhold.
   */
  const startRedaction = useCallback(() => {
    if (!editor) return
    const { from, to } = editor.state.selection
    if (from === to) return
    setRedactTarget({
      from,
      to,
      seed: editor.state.doc.textBetween(from, to, ' ').trim().slice(0, 24),
    })
  }, [editor])

  /**
   * Stop withholding whatever the selection carries.
   *
   * The command reads the selection itself and clears the mark across it, so
   * the same selection the menu decided on is the one that gets cleaned: the
   * menu closes before this runs, but closing a menu does not move a caret.
   * No dialog, because there is nothing to ask — the words were never replaced,
   * so there is no label to revise and nothing to confirm.
   */
  const clearRedaction = useCallback(() => {
    if (!editor) return
    editor.commands.unsetRedaction()
  }, [editor])

  const submitNewComment = useCallback(
    (text: string, image?: CommentImage) => submitNewCommentImpl(reviewCtxRef.current, text, image),
    [],
  )
  const replyToComment = useCallback(
    (parentId: string, text: string) => replyToCommentImpl(reviewCtxRef.current, parentId, text),
    [],
  )
  const editComment = useCallback(
    (id: string, text: string) => editCommentImpl(reviewCtxRef.current, id, text),
    [],
  )
  const resolveComment = useCallback(
    (id: string, done: boolean) => resolveCommentImpl(reviewCtxRef.current, id, done),
    [],
  )
  const deleteComment = useCallback((id: string) => deleteCommentImpl(reviewCtxRef.current, id), [])
  const handleRevision = useCallback(
    (action: 'accept' | 'reject', scope: RevisionScope) =>
      handleRevisionImpl(reviewCtxRef.current, action, scope),
    [],
  )
  const handleRevisionRange = useCallback(
    (action: 'accept' | 'reject', target: RevisionRange) =>
      handleRevisionRangeImpl(reviewCtxRef.current, action, target),
    [],
  )
  // Word: navigating onto a change hidden by the current view shows All
  // Markup; the view switch must commit before the scroll can find the text
  const showRevision = useCallback(
    (target: RevisionRange) => {
      if (!editor) return
      const show = target.story
        ? () => showStoryRevisionRef.current(target as StoryRevision)
        : () => {
            setActiveStoryId(null)
            selectRevision(editor, target)
          }
      if (revisionHiddenIn(revisionDisplayRef.current, target)) {
        setRevisionDisplay('all')
        window.setTimeout(show, 0)
      } else show()
    },
    [editor],
  )
  /** the change outside the body Next / Previous last opened: the pane highlights it, Accept Next takes it */
  const [activeStoryId, setActiveStoryIdState] = useState<string | null>(null)
  const lastStoryRef = useRef<string | null>(null)
  const setActiveStoryId = (id: string | null) => {
    lastStoryRef.current = id
    setActiveStoryIdState(id)
  }
  // Word opens the story the change lives in: the note's reference is selected and
  // its entry flashed, the header / footer strip opens for editing, a textbox's
  // anchor paragraph is selected (the box is drawn beside it)
  const showStoryRevision = (target: StoryRevision) => {
    if (!editor) return
    const { story } = target
    if (story.hf) {
      const { section: si, kind } = story.hf
      // the edge areas are the first section's first-page header and the last
      // section's last-page footer; any other strip sits in a page gap
      const last = Math.max(0, hfStateRef.current.sections.length - 1)
      const edge = kind === 'header' ? si === 0 : si === last
      const strip = edge
        ? null
        : document.querySelector<HTMLElement>(
            `.page-gap-hf.page-hf-${kind}[data-hf-section="${si}"]`,
          )
      if (strip) startGapHfEditRef.current(strip)
      else if (edge) openHfEditorRef.current(kind)
      setActiveStoryId(story.id)
      return
    }
    if (target.from >= 0 && target.from <= editor.state.doc.content.size) {
      const tr = editor.state.tr
      tr.setSelection(TextSelection.near(tr.doc.resolve(target.from), 1))
      tr.scrollIntoView()
      tr.setMeta(TRACK_IGNORE, true)
      editor.view.dispatch(tr)
      editor.view.focus()
    }
    // after the dispatch: the selection listener clears the cursor on body moves
    setActiveStoryId(story.id)
    if (story.noteId) {
      window.setTimeout(() => {
        const row = document.querySelector<HTMLElement>(
          `.page-gap-note[data-note-id="${CSS.escape(story.noteId!)}"]`,
        )
        if (!row) return
        row.scrollIntoView({ block: 'nearest' })
        row.classList.add('rev-target')
        window.setTimeout(() => row.classList.remove('rev-target'), 1500)
      }, 50)
    }
  }
  const showStoryRevisionRef = useRef(showStoryRevision)
  showStoryRevisionRef.current = showStoryRevision
  const addInk = useCallback(
    (annotation: InkAnnotation) => addInkImpl(reviewCtxRef.current, annotation),
    [],
  )
  const removeInks = useCallback((ids: string[]) => removeInksImpl(reviewCtxRef.current, ids), [])
  const clearInks = useCallback(() => clearInksImpl(reviewCtxRef.current), [])

  // the ribbon's revision badge walks every block: paced with the word count below
  const [revisionRanges, setRevisionRanges] = useState<RevisionRange[]>([])
  // open comment threads add the markup column right of the paper: width-fit
  // zoom must count it or the canvas overflows the pane
  // (Simple Markup collapses threads to glyphs inside the page margin: no column)
  // body revision balloons keep the column even while the Comments panel hosts the threads
  const markupExtra = useMemo(() => {
    if (revisionDisplay === 'simple') return 0
    const commentCards = commentsUi.cards && comments.some((c) => !c.parentId && !c.done)
    const revKinds = balloonsActive ? balloonRevKinds(balloonMode) : null
    const revBalloons =
      !!revKinds &&
      revisionRanges.some((r) => {
        const k = balloonKindOf(r.kind)
        return k != null && revKinds.has(k)
      })
    return (revBalloons ? tcOptions.balloonWidth : 0) + (commentCards ? COMMENT_AREA_W : 0)
  }, [
    comments,
    revisionRanges,
    commentsUi.cards,
    revisionDisplay,
    balloonMode,
    balloonsActive,
    tcOptions.balloonWidth,
  ])
  // inline comment tags ([initials N]) read the thread list through the decoration state
  useEffect(() => {
    setInlineComments(comments)
    if (editor && balloonMode === 'inline')
      editor.view.dispatch(editor.state.tr.setMeta('addToHistory', false))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [comments])

  /** Uncapped width-fit ratio (%) from the scroller's measured size */
  const rawFitWidth = useCallback(() => {
    if (!section) return null
    const scroller = document.querySelector('.editor-scroll')
    if (!scroller) return null
    return ((scroller.clientWidth - 48) / (twipsToPx(section.pageWidth) + markupExtra)) * 100
  }, [section, markupExtra])

  /** Last auto-fit: if current zoom still equals its value → "fit mode", re-fit on size changes */
  const lastFitRef = useRef<{ mode: ZoomFitMode; value: number } | null>(null)
  const zoomLiveRef = useRef(zoom)
  useEffect(() => {
    zoomLiveRef.current = zoom
  }, [zoom])

  /** Word's page-width / text-width / whole-page zoom: the zoom ratio for the current window size */
  const fitZoomFor = useCallback(
    (mode: ZoomFitMode): number | null => {
      if (!section) return null
      const scroller = document.querySelector('.editor-scroll')
      if (!scroller) return null
      const pad = 48
      const wFit =
        ((scroller.clientWidth - pad) / (twipsToPx(section.pageWidth) + markupExtra)) * 100
      const hFit = ((scroller.clientHeight - pad) / twipsToPx(section.pageHeight)) * 100
      const textW = twipsToPx(section.pageWidth - section.marginLeft - section.marginRight)
      const tFit = ((scroller.clientWidth - pad) / (textW + markupExtra)) * 100
      if (mode === 'multi') {
        // Word: a whole page tall and as many across as fit, never fewer than two
        const slotW = twipsToPx(section.pageWidth) + markupExtra
        const byHeight = clampDocsZoom(Math.floor(hFit))
        if (multipageColumns(scroller.clientWidth, slotW, byHeight) >= 2) return byHeight
        return clampDocsZoom(Math.floor(multipageZoomForColumns(scroller.clientWidth, slotW, 2)))
      }
      // whole page = the entire page visible, so it must fit both dimensions;
      // floor, not round: rounding up would push the page past the pane edge
      const next = mode === 'width' ? wFit : mode === 'text' ? tFit : Math.min(wFit, hFit)
      return clampDocsZoom(Math.floor(next))
    },
    [section, markupExtra],
  )
  // Word: Multiple Pages turns the side-by-side arrangement on and it stays
  // through later zooming (a manual zoom-out alone keeps the single column);
  // One Page / Page Width / Text Width turn it off
  const [multipageMode, setMultipageMode] = useState(false)
  /** fit without touching the arrangement: the resize follow-up / overflow clamp below */
  const refitZoom = useCallback(
    (mode: ZoomFitMode) => {
      const applied = fitZoomFor(mode)
      if (applied == null) return
      lastFitRef.current = { mode, value: applied }
      setZoom(applied)
    },
    [fitZoomFor],
  )
  const zoomFit = useCallback(
    (mode: ZoomFitMode) => {
      setMultipageMode(mode === 'multi')
      refitZoom(mode)
    },
    [refitZoom],
  )

  // On scroller size changes (window/AI-dock/nav-pane toggles): follow with a
  // re-fit while in fit mode, and clamp any manual zoom back down to width-fit
  // whenever the page no longer fits horizontally — the canvas must never
  // overflow the pane on a resize (same contract as the slides stage). A manual
  // zoom smaller than fit is left alone.
  useEffect(() => {
    const el = document.querySelector('.editor-scroll')
    if (!doc || !el) return
    let lastWidth = -1
    const ro = new ResizeObserver(() => {
      const raw = rawFitWidth()
      if (raw == null) return
      const widthChanged = el.clientWidth !== lastWidth
      lastWidth = el.clientWidth
      const lf = lastFitRef.current
      if (lf != null && Math.abs(zoomLiveRef.current - lf.value) <= 0.5) {
        refitZoom(lf.mode) // fit mode: follow the container in the user's chosen fit
        return
      }
      // The overflow test uses the uncapped ratio: a manual zoom that still
      // fits (or is smaller than fit) is the user's choice and must be kept.
      // A height-only change (a classic horizontal scrollbar appearing for the
      // zoomed-in page) must not clamp a zoom the user just chose.
      if (!widthChanged || zoomLiveRef.current <= raw + 0.5) return
      refitZoom('width')
    })
    ro.observe(el)
    return () => ro.disconnect()
  }, [doc, refitZoom, rawFitWidth])

  // markup column appearing/disappearing changes the canvas width without a
  // pane resize: apply the same follow/clamp rules as the ResizeObserver above
  useEffect(() => {
    if (!doc) return
    const raw = rawFitWidth()
    if (raw == null) return
    const lf = lastFitRef.current
    if (lf != null && Math.abs(zoomLiveRef.current - lf.value) <= 0.5) refitZoom(lf.mode)
    else if (zoomLiveRef.current > raw + 0.5) refitZoom('width')
  }, [doc, markupExtra, rawFitWidth, refitZoom])

  // Read Mode opens at 1:1 — the same page size as the Page Preview — and the
  // user's zoom / fit mode is restored on exit (wheel zoom still works inside).
  // Layout effect: it must snapshot/restore before the ResizeObserver reacts to
  // the .read-mode chrome toggling (RO callbacks fire at layout time, after
  // layout effects but before passive effects), or zoomFit would rewrite
  // lastFitRef first; zoomLiveRef is synced here for the same reason.
  const preReadZoomRef = useRef<{
    zoom: number
    fit: { mode: ZoomFitMode; value: number } | null
  } | null>(null)
  useLayoutEffect(() => {
    if (readMode) {
      preReadZoomRef.current = { zoom: zoomLiveRef.current, fit: lastFitRef.current }
      lastFitRef.current = null
      zoomLiveRef.current = 100
      setZoom(100)
    } else if (preReadZoomRef.current) {
      const prev = preReadZoomRef.current
      preReadZoomRef.current = null
      lastFitRef.current = prev.fit
      zoomLiveRef.current = prev.zoom
      setZoom(prev.zoom)
    }
  }, [readMode])

  // separator rows above the note areas: the notes parts' separator paragraphs
  // under the document's styles (module-level: the slicer, gaps and preview read one row)
  useLayoutEffect(() => {
    const rowOf = (sep?: NoteSeparator) =>
      noteSeparatorRow(
        section?.docGrid,
        doc && sep ? resolveNoteStyle(doc.parsed, sep.styleId, sep.spacing) : undefined,
      )
    setNoteSeparatorRows({
      footnote: rowOf(doc?.parsed.footnoteSeparator),
      endnote: rowOf(doc?.parsed.endnoteSeparator),
    })
  }, [doc, section])

  // note paragraph metrics per pStyle + direct spacing (notes without either use Normal/docDefaults)
  const noteStyleOf = useMemo(() => {
    const cache = new Map<string, ReturnType<typeof resolveNoteStyle>>()
    return (note?: Pick<NoteInfo, 'styleId' | 'spacing' | 'richParas'>) => {
      const runStyle = noteRunStyle(note?.richParas)
      const key = `${note?.styleId ?? ''}|${JSON.stringify(note?.spacing ?? null)}|${runStyle.fontFamily ?? ''}|${runStyle.sizeHalfPoints ?? ''}`
      let v = cache.get(key)
      if (!v) {
        v = doc
          ? { ...resolveNoteStyle(doc.parsed, note?.styleId, note?.spacing), ...runStyle }
          : {}
        cache.set(key, v)
      }
      return v
    }
  }, [doc])

  // per-note render metrics: resolved style line height/size plus the entry's
  // height at the renderers' exact styles (DOM wrap truth; the char-width
  // estimate stays as the DOM-less fallback and the parity runner's model)
  const noteRenderInfoOf = useMemo(() => {
    const cache = new Map<
      string,
      { height: number; lineHeightPx: number; fontSizePt: number; fontFamily?: string }
    >()
    return (
      fn: NoteInfo | undefined,
      no: number,
      sec: SectionSettings,
      kind: NoteKind,
      fmt?: string,
    ): { height: number; lineHeightPx: number; fontSizePt: number; fontFamily?: string } => {
      const contentW = twipsToPx(sec.pageWidth - sec.marginLeft - sec.marginRight)
      // the grid snaps the line height: sections of one width can differ in it
      const key = `${kind}|${fn?.id ?? ''}|${no}|${fmt ?? ''}|${Math.round(contentW)}|${sec.docGrid?.type ?? ''}:${sec.docGrid?.linePitch ?? ''}|${fn?.styleId ?? ''}|${fn?.text ?? ''}|${revisionDisplay}`
      let v = cache.get(key)
      if (!v) {
        const style = noteStyleOf(fn)
        const { lineHeightPx, fontSizePt, fontFamily } = noteRowMetrics(style, sec.docGrid)
        const richParas = noteRunsInView(revisionDisplay, fn?.richParas)
        const height =
          measureNoteHeightDom(
            {
              no,
              ...(fmt ? { fmt } : {}),
              text: fn?.text ?? '',
              ...(richParas ? { richParas } : {}),
              ...(fn?.noRefMark ? { noRefMark: true as const } : {}),
            },
            kind,
            contentW,
            lineHeightPx,
            fontSizePt,
            fontFamily,
          ) ??
          estimateFootnoteHeight(fn?.text ?? '', contentW, sec.docGrid, undefined, style, richParas)
        v = { height, lineHeightPx, fontSizePt, ...(fontFamily ? { fontFamily } : {}) }
        cache.set(key, v)
      }
      return v
    }
    // note-list deps only reset the cache (keys carry the note text, but
    // formatting-only edits keep it — a fresh map re-measures them)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [footnotes, endnotes, noteStyleOf, revisionDisplay])

  // the numbers the body shows right now: inserting or deleting a note renumbers
  // its in-text marks in document order, and the note areas must say the same
  const noteMarkNumbers = useMemo(() => {
    const out = new Map<string, { num: number; fmt?: string }>()
    editorRef.current?.state.doc.descendants((node) => {
      if (node.type.name !== 'docNoteRef') return true
      const key = `${node.attrs.kind}:${node.attrs.id}`
      if (!out.has(key))
        out.set(key, {
          num: Number(node.attrs.num),
          ...(typeof node.attrs.fmt === 'string' ? { fmt: node.attrs.fmt } : {}),
        })
      return false
    })
    return out
    // the lists change whenever marks are added or removed, the note settings when they are renumbered
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [footnotes, endnotes, doc, sectionNoteProps, noteRenumbered])

  // display number of a note: the mark in the body, else the engine's body-order
  // numbering (numStart / eachSect applied), else the list position
  const noteNo = useCallback(
    (kind: 'footnote' | 'endnote', id: string, index: number): number =>
      noteMarkNumbers.get(`${kind}:${id}`)?.num ??
      doc?.parsed.noteNumbers?.[`${kind}:${id}`] ??
      index + 1,
    [noteMarkNumbers, doc],
  )
  // the number format of a note: its section's, as its mark in the body carries it
  const noteFmt = useCallback(
    (kind: 'footnote' | 'endnote', id: string): string | undefined => {
      const mark = noteMarkNumbers.get(`${kind}:${id}`)
      return mark ? mark.fmt : doc?.parsed.noteNumFmts?.[`${kind}:${id}`]
    },
    [noteMarkNumbers, doc],
  )

  // page-bottom heights (px) reserved for footnote references inside a block, one per
  // reference in run order, measured with the marker the page will draw
  const footnoteBandsOf = useCallback(
    (b: Block): Array<{ heightPx: number }> => {
      const runs = blockNoteScanRuns(b)
      if (runs.length === 0 || footnotes.length === 0) return []
      const noOf = new Map(footnotes.map((f, i) => [f.id, noteNo('footnote', f.id, i)]))
      const bands: Array<{ heightPx: number }> = []
      for (const run of runs) {
        if (run.noteRef?.kind !== 'footnote') continue
        const sec =
          sections.find((s) => (b.docxIndex ?? 0) <= s.lastBlockIndex)?.settings ?? section
        if (!sec) continue
        const fn = footnotes.find((f) => f.id === run.noteRef!.id)
        bands.push({
          heightPx: noteRenderInfoOf(
            fn,
            noOf.get(run.noteRef!.id) ?? 0,
            sec,
            'footnote',
            noteFmt('footnote', run.noteRef!.id),
          ).height,
        })
      }
      // the once-per-page separator is charged by the pagination engine, not per block
      return bands
    },
    [footnotes, sections, section, noteRenderInfoOf, noteNo, noteFmt],
  )

  // typed w:docGrid line pitch (pt) when every section shares one; null = no snapping
  const gridPitchPt = useMemo(() => docGridPitchPt(sections), [sections])
  const typedGridNow = useMemo(
    () => sections.some((s) => sectionGridPitchPt(s) != null),
    [sections],
  )
  useEffect(() => {
    const parsed = doc?.parsed
    if (!parsed) return
    const kept = liveTypedGridRef.current
    const first = kept?.parsed !== parsed
    liveTypedGridRef.current = { parsed, typed: typedGridNow }
    // a document just read: its CSS was built from the same sections
    if (first || kept.typed === typedGridNow) return
    const live = liveStylesRef.current
    setDocCss(liveDocCss(live?.parsed === parsed ? { ...parsed, styles: live.styles } : parsed))
  }, [typedGridNow, doc?.parsed, liveDocCss])
  // mixed grids: .doc-page keeps the first typed pitch as the inheritance
  // fallback (floats/notes/web view); print blocks get per-section overrides
  // via sectionGridPitchSpecs on the layout channel
  const mixedGridPitchPt = useMemo(() => {
    if (gridPitchPt != null) return null
    for (const s of sections) {
      const p = sectionGridPitchPt(s)
      if (p != null) return p
    }
    return null
  }, [sections, gridPitchPt])

  // w:docGrid charSpace character grid: uniform per-character letter-spacing
  // delta (pt); mixed docs deliver it per block via sectionCharSpaceSpecs
  const charSpacePt = useMemo(() => docCharSpacePt(sections), [sections])
  // the default paragraph style's size, which a snapToChars section counts its cells from
  const gridBasePtRef = useRef(12)
  // balanceSingleByteDoubleByteWidth: single-byte characters get half the delta
  const balancedGrid = doc?.parsed.balanceDbcsSpacing === true

  // page color: Word lays out one extra blank Normal paragraph under section 1's header
  const hfPhantom = useMemo(
    () => (doc && pageColor ? hfPhantomSpec(doc.parsed) : null),
    [doc, pageColor],
  )

  // compatibilityMode < 15: header wrapTopAndBottom bands clamp to the page top (hfStripGeom)
  const legacyBands = (doc?.parsed.compatibilityMode ?? 0) < 15
  // single-section header/footer push-down: body top = max(marginTop, headerDist + header height)
  const singleHfPx = useMemo((): SectionHfHeights => {
    if (!section) return { headerPx: 0, footerPx: 0 }
    const contentW = twipsToPx(section.pageWidth - section.marginLeft - section.marginRight)
    // live titlePg: the first page draws the live first-page variant (pageHfOf),
    // so its reserved heights track the live state, not the parsed parts
    return hfVariantHeightsPx(
      (kind, variant) => hfLayoutResolved(kind, hfResolveAt(0, kind, variant), hfPhantom),
      () => contentW,
      hfStripGeom(section, legacyBands),
      titlePg,
      evenOddHf,
    )
    // hfMeasureEpoch: re-measure once the doc-scoped styles have committed
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [section, hfResolveAt, hfPhantom, titlePg, evenOddHf, legacyBands, hfMeasureEpoch])
  const effTopSingle = section ? effectiveTopPx(section, singleHfPx.headerPx) : 0
  const effBottomSingle = section ? effectiveBottomPx(section, singleHfPx.footerPx) : 0
  const canvasSection = sections[0]?.settings ?? section
  const canvasBox = canvasSection ? sectionPageBox(canvasSection) : null
  // the canvas's first page draws the variant pageHfOf picks for it (titlePg
  // first page, else the even variant when its number is even): its body
  // starts under that strip, as the preview slicer already assumes
  const firstPageHeaderPx = useMemo(
    (): number =>
      hfVariantReservedPx(
        singleHfPx,
        hfVariantOf(titlePg, true, evenOddHf, sections[0]?.pageNumberStart ?? 1),
      ).headerPx,
    [singleHfPx, titlePg, evenOddHf, sections],
  )
  const canvasTop = section ? canvasContentTopPx(sections, section, firstPageHeaderPx) : 0
  // trailing paper padding belongs to the last page, which draws the default footer
  const canvasBottom = canvasSection ? effectiveBottomPx(canvasSection, singleHfPx.footerPx) : 0

  const singleVariantGeom = useMemo(
    () => ({
      ...(section ? singleVariantGeomOf(section, singleHfPx) : {}),
      ...(sections[0]?.pageNumberStart !== undefined
        ? { pageNumberStart: sections[0].pageNumberStart }
        : {}),
    }),
    [section, singleHfPx, sections],
  )

  // multi-section: estimated heights of each section's header/footer variants (capacity per section)
  const hfHeightsOf = useCallback(
    (secs: SectionInfo[]): SectionHfHeights[] => {
      // pending section breaks make `secs` differ from the saved list: resolve against them
      const liveState: HfSectionState = { ...hfState, sections: secs }
      return sectionHfHeights(
        secs,
        (i, kind, variant) =>
          hfLayoutResolved(kind, resolveHf(liveState, i, kind, variant), hfPhantom),
        { titlePg, evenOddHf, legacyBands },
      )
    },
    // hfMeasureEpoch: re-measure once the doc-scoped styles have committed
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [hfState, hfPhantom, titlePg, evenOddHf, legacyBands, hfMeasureEpoch],
  )

  // paragraphs without w:pStyle take the default paragraph style's constraints (widowControl off in Normal)
  const defaultParaStyle = useMemo(
    () =>
      doc
        ? [...doc.parsed.styles.values()].find((s) => s.type === 'paragraph' && s.isDefault)
        : undefined,
    [doc],
  )
  // pagination-constraint injection: docxIndex → parse-layer semantics (keepNext/keepLines/widow/table-row flags).
  // DOM measurement only has geometry; these constraints decide page cut points (no orphan headings / unbreakable lines / repeated table headers).
  const blockByDocxIndex = useMemo(() => {
    const m = new Map<number, Block>()
    for (const b of doc?.parsed.blocks ?? [])
      if (b.docxIndex !== null && !m.has(b.docxIndex)) m.set(b.docxIndex, b)
    return m
  }, [doc])
  const blockMetaOf = useCallback(
    (docxIndex: number): BlockMeta | undefined => {
      const b = blockByDocxIndex.get(docxIndex)
      if (!b) return undefined
      if (b.type === 'table') {
        // notes referenced inside cells reserve their page-bottom area like a
        // paragraph's: the slicer charges them to the rows holding the marks
        const fnBands = footnoteBandsOf(b)
        const fnExtra = fnBands.reduce((s, band) => s + band.heightPx, 0)
        const styleKeepNext = (id: string) => doc?.parsed.styles.get(id)?.display?.keepNext === true
        const xml = b.originalXml ?? ''
        const flagged =
          /tblHeader|cantSplit|keepNext|<w:trHeight\b/.test(xml) ||
          [...xml.matchAll(/<w:pStyle w:val="([^"]+)"/g)].some((m) => styleKeepNext(m[1]))
        const modern = (doc?.parsed.compatibilityMode ?? 0) >= 15
        if (!flagged && fnExtra === 0 && !modern) return undefined
        const cellWidowOff =
          modern &&
          tableWidowOff(
            xml,
            defaultParaStyle?.display?.widowControl ?? doc?.parsed.docDefaults?.widowControl,
            (id) => doc?.parsed.styles.get(id)?.display?.widowControl,
          )
        return {
          ...(flagged ? { tableRowFlags: tableRowFlags(xml, styleKeepNext) } : {}),
          ...(modern ? { modernTableHeaders: true } : {}),
          ...(cellWidowOff ? { cellWidowOff: true } : {}),
          ...(fnExtra > 0 ? { footnoteExtraPx: fnExtra, footnoteBands: fnBands } : {}),
        }
      }
      const styleOf = (id: string | undefined) =>
        (id ? doc?.parsed.styles.get(id) : defaultParaStyle)?.display
      const flags = b.format
        ? paraPaginationMeta(b.format)
        : imageParaPaginationMeta(b.originalXml, styleOf, doc?.parsed.docDefaults)
      const fnBands = footnoteBandsOf(b)
      const fnExtra = fnBands.reduce((s, band) => s + band.heightPx, 0)
      const modern = (doc?.parsed.compatibilityMode ?? 0) >= 15
      if (!flags && fnExtra === 0 && !modern) return undefined
      return {
        ...flags,
        ...(modern ? { modernAnchorFloats: true } : {}),
        ...(fnExtra > 0 ? { footnoteExtraPx: fnExtra, footnoteBands: fnBands } : {}),
      }
    },
    [doc, blockByDocxIndex, defaultParaStyle, footnoteBandsOf],
  )
  // shadow BlockPlan inputs (I9): the model-side twin of the measureBlocks scrape
  const planProps = useMemo(
    () => (doc ? planPropsOf(doc.parsed, sections, revisionDisplay === 'original') : null),
    [doc, sections, revisionDisplay],
  )
  // the BlockPlan of the current editor state keyed by rendered element; built
  // per measure so a transaction between passes cannot leave it stale
  const blockPlans = useCallback(
    (): PlanOf | undefined =>
      editor && planProps
        ? planByElement(editor.state.doc, editor.view, planProps, floatShiftsOf(editor.state))
        : undefined,
    [editor, planProps],
  )

  // per-page footnote collection: the page of the referencing block → that page's footnote entries (number/text/estimated height)
  const pageFootnotesOf = useCallback(
    (blocks: BlockBox[], slices: PageSlice[]): PageNoteItem[][] => {
      const out: PageNoteItem[][] = slices.map(() => [])
      if (!doc || footnotes.length === 0 || !section) return out
      const noOf = new Map(footnotes.map((f, i) => [f.id, noteNo('footnote', f.id, i)]))
      for (const b of blocks) {
        if (b.docxIndex === undefined) continue
        const pb = blockByDocxIndex.get(b.docxIndex)
        if (!pb) continue
        const ids = blockNoteScanRuns(pb)
          .filter((r) => r.noteRef?.kind === 'footnote')
          .map((r) => r.noteRef!.id)
        if (ids.length === 0) continue
        const blockPage = pageOfY(slices, b.top)
        const sec = sections.find((s) => b.docxIndex! <= s.lastBlockIndex)?.settings ?? section
        for (let ri = 0; ri < ids.length; ri++) {
          const id = ids[ri]
          const fn = footnotes.find((f) => f.id === id)
          if (!fn) continue
          // notes go to the page holding their reference line (a split
          // paragraph spreads its notes like Word); marker offsets were
          // resolved into noteBands (run order) by applyBlockMeta
          const band = b.noteBands?.length === ids.length ? b.noteBands[ri] : undefined
          const page = band
            ? pageOfY(slices, b.top + (b.spaceBeforePx ?? 0) + band.offset)
            : blockPage
          const fmt = noteFmt('footnote', id)
          const info = noteRenderInfoOf(fn, noOf.get(id) ?? 0, sec, 'footnote', fmt)
          const richParas = noteRunsInView(revisionDisplay, fn.richParas)
          out[page]?.push({
            no: noOf.get(id) ?? 0,
            ...(fmt ? { fmt } : {}),
            id,
            text: fn.text,
            ...(richParas ? { richParas } : {}),
            ...(fn.noRefMark ? { noRefMark: true as const } : {}),
            height: info.height,
            lineHeightPx: info.lineHeightPx,
            fontSizePt: info.fontSizePt,
            ...(info.fontFamily ? { fontFamily: info.fontFamily } : {}),
          })
        }
      }
      return out
    },
    [
      doc,
      blockByDocxIndex,
      footnotes,
      sections,
      section,
      noteRenderInfoOf,
      noteNo,
      noteFmt,
      revisionDisplay,
    ],
  )

  // an endnote entry as drawn in a section: its height and line box at that section's width and grid
  const endnoteItemIn = useCallback(
    (n: NoteInfo, i: number, sec: SectionSettings): PageNoteItem => {
      const no = noteNo('endnote', n.id, i)
      const fmt = noteFmt('endnote', n.id)
      const info = noteRenderInfoOf(n, no, sec, 'endnote', fmt)
      const richParas = noteRunsInView(revisionDisplay, n.richParas)
      return {
        ...(fmt ? { fmt } : {}),
        no,
        id: n.id,
        text: n.text,
        ...(richParas ? { richParas } : {}),
        ...(n.noRefMark ? { noRefMark: true as const } : {}),
        height: info.height,
        lineHeightPx: info.lineHeightPx,
        fontSizePt: info.fontSizePt,
        ...(info.fontFamily ? { fontFamily: info.fontFamily } : {}),
      }
    },
    [noteRenderInfoOf, noteNo, noteFmt, revisionDisplay],
  )
  // endnote-area entries (shared by pagination preview and page slicing), in the final section
  const allEndnoteItems = useMemo<PageNoteItem[]>(() => {
    if (!section || endnotes.length === 0) return []
    const sec = sections[sections.length - 1]?.settings ?? section
    return endnotes.map((n, i) => endnoteItemIn(n, i, sec))
  }, [endnotes, sections, section, endnoteItemIn])

  // w:endnotePr w:pos="sectEnd": each section's endnotes print at its end (a section
  // that suppresses them hands them on); the last section's stay in the end area
  const sectEndnotesOn =
    noteSettings.endnote.pos === 'sectEnd' && sections.length > 1 && endnotes.length > 0
  const endnoteHomesRef = useRef<EndnoteHomes | null>(null)
  // read while rendering, so the first layout of a file already has the rows in place
  const endnoteHomes = useMemo(() => {
    if (!editor || !sectEndnotesOn) return null
    const next = endnoteHomesOf(editor.state.doc, sections)
    if (endnoteHomesRef.current?.sig !== next.sig) endnoteHomesRef.current = next
    return endnoteHomesRef.current
    // docVersion: the doc identity is what changes, and it is not a React value
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editor, sectEndnotesOn, sections, endnotes, docVersion])
  const endnoteLayout = useMemo(() => {
    if (!sectEndnotesOn || !endnoteHomes) return null
    const { groups, tail } = groupSectionEndnotes(
      allEndnoteItems,
      endnoteHomes,
      sections.map((s) => !!s.settings.noEndnote),
    )
    const listed = new Map(endnotes.map((n, i) => [n.id, [n, i] as const]))
    return {
      tail,
      // a row takes the width and grid of the section it is printed in
      groups: groups.map((group) => {
        const sec = sections[group.section]?.settings
        return {
          ...group,
          items: group.items.map((item) => {
            const at = listed.get(item.id)
            return sec && at ? endnoteItemIn(at[0], at[1], sec) : item
          }),
        }
      }),
    }
  }, [sectEndnotesOn, endnoteHomes, allEndnoteItems, sections, endnotes, endnoteItemIn])
  /** the entries of the end-of-document area */
  const endnoteItems = endnoteLayout?.tail ?? allEndnoteItems
  const sectEndnoteGroups = endnoteLayout?.groups ?? NO_SECT_ENDNOTES
  const tailEndnotes = useMemo(() => {
    if (!endnoteLayout) return endnotes
    const ids = new Set(endnoteLayout.tail.map((item) => item.id))
    return endnotes.filter((n) => ids.has(n.id))
  }, [endnotes, endnoteLayout])
  const editNoteRef = useRef(editNote)
  editNoteRef.current = editNote
  // before the pagination pass below: it measures the rows this puts in the flow
  useEffect(() => {
    if (editor)
      setSectionEndnotes(editor.view, sectEndnoteGroups, sections, (id) =>
        editNoteRef.current('endnote', id),
      )
  }, [editor, sectEndnoteGroups, sections])

  // canvas column mode:
  //  - 'uniform': every section shares one equal-width multi-column spec (and is LTR) —
  //    whole-page CSS multicol renders it (browser splits paragraphs across columns natively)
  //  - 'mixed': some multi-column section coexists with other specs (or RTL columns) —
  //    per-block column-layout decorations paint the engine's regions (block granularity)
  //  - 'none': no multi-column sections.
  // equalWidth="0" (unequal local layout columns) is not modeled, matching the engine's scope
  const colMode = useMemo<'none' | 'uniform' | 'mixed'>(() => {
    if (sections.length === 0) return 'none'
    if (!sections.some((s) => sectionColumns(s) > 1)) return 'none'
    const g0 = sectionColGeom(sections[0])
    const uniform =
      !sections.some(sectionBidi) &&
      // a same-count nextColumn boundary advances a column and a continuous one
      // balances the closed columns mid-page — whole-page CSS multicol can't
      // paint either, so such documents go through mixed mode
      !sections.some(
        (s, i) => i > 0 && (s.startType === 'nextColumn' || s.startType === 'continuous'),
      ) &&
      sections.every((s) => {
        const g = sectionColGeom(s)
        return (
          g.cols === g0.cols &&
          g.cols > 1 &&
          g.equalWidth &&
          Math.abs(g.colWidthPx - g0.colWidthPx) < 0.5 &&
          Math.abs(g.gapPx - g0.gapPx) < 0.5
        )
      })
    return uniform ? 'uniform' : 'mixed'
  }, [sections])

  // uniform-mode geometry for the whole-page CSS multicol path / measuring width swap
  const colFlow = useMemo(
    () => (colMode === 'uniform' ? sectionColGeom(sections[0]) : null),
    [colMode, sections],
  )

  // sectPr w:vAlign pages carry visual block translates (vAlignShiftSpecs), so
  // measurement must neutralize them exactly like mixed-column translates
  const hasVAlign = useMemo(
    () => sections.some((s) => s.settings.vAlign === 'center' || s.settings.vAlign === 'bottom'),
    [sections],
  )
  // vertical-text sections (verticalTextSpecs) ride the same translate channel
  const hasVertical = useMemo(() => sections.some((s) => sectionVertical(s.settings)), [sections])

  // single-flow measuring state for the columned canvas: uniform mode temporarily drops
  // the CSS columns and sets the width to the column width; mixed mode neutralizes the
  // per-block translates/gap compression (widths stay — line boxes must reflect column
  // wrapping). Either way DOM measurement yields 1-D coordinates matching the engine's
  // column flow (synchronous layout round-trip, no visible flicker)
  const measureSingleFlow = useCallback(
    function run<T>(pm: HTMLElement, fn: () => T): T {
      if (
        (colMode === 'none' && !hasVAlign && !hasVertical && !mirrorMargins) ||
        viewMode !== 'print'
      )
        return fn()
      pm.classList.add('measuring-columns')
      try {
        return fn()
      } finally {
        pm.classList.remove('measuring-columns')
      }
    },
    [colMode, hasVAlign, hasVertical, mirrorMargins, viewMode],
  )

  // column-flow geometry gate: when the canvas column layout is inactive, measure as full-width single flow; the geometry must drop cols to match
  const colGeomsFor = useCallback(
    (geoms: SectionGeom[]): SectionGeom[] => {
      if (colMode !== 'none' && viewMode === 'print') return geoms
      for (const g of geoms) if (g.cols) g.cols = undefined
      return geoms
    },
    [colMode, viewMode],
  )

  // real TOC page-number backfill: each heading's (docHeading) page from the
  // current pass's PageCut list. Returns page numbers matching docHeadings in
  // document order 1:1; null when not computable (page numbers left blank).
  const headingPages = useCallback((): number[] | null => {
    if (!editor || !section) return null
    const layout = layoutNow()
    if (!layout || layout.cuts.length === 0) return null
    const { pages, misses } = headingPagesOf(
      layout.blocks,
      layout.cuts,
      collectHeadings(editor.state.doc, editor.storage.listNumbering?.styles).map(
        (h) => editor.view.nodeDOM(h.pos) as HTMLElement | null,
      ),
    )
    const dbg = (window as unknown as { __pageDebug?: Record<string, unknown> }).__pageDebug
    if (dbg) dbg.headingPageMisses = ((dbg.headingPageMisses as number) ?? 0) + misses
    return pages
  }, [editor, section, layoutNow])

  // status-bar page number: real page slicing (same algorithm as the pagination preview). Edits remeasure with debounce; scrolling only relocates
  useEffect(() => {
    if (!doc || !section) {
      setPageInfo({ current: 1, total: 1 })
      setLastPageNo(null)
      setFirstPageNo(null)
      setEdgeSectionPages(null)
      pageLayoutRef.current = null
      return
    }
    const scroller = document.querySelector('.editor-scroll')
    if (!scroller) return
    // Blink lays text out at the zoomed font size and snaps line boxes to device
    // pixels, so block heights read under a 10% (or 88%) canvas zoom differ from
    // the 100% layout by up to a pixel per line — enough to move page breaks and
    // to park a paragraph in the wrong column of an unequal-width section
    // (genoffice#1970). Passes measure the canvas at zoom 1 and restore the
    // view zoom before the viewport-dependent reads (locate).
    let factor = zoom / 100
    const mTopPx = canvasTop
    // shared paper = the widest section's page; narrower pages are centered on it
    const paperWPx = paperWidthPx(sections, sections[0]?.settings ?? section)
    const pageLeftOf = (set: SectionSettings) => pageLeftPx(set, paperWPx)
    const mixedPaper = sections.some((s) => pageLeftOf(s.settings) > 0.5)
    const contentH = twipsToPx(section.pageHeight) - effTopSingle - effBottomSingle
    let slices: PageSlice[] = []
    let timer: number | null = null
    let rowFillSig = ''
    // passes a pass schedules for itself (widths / suppression applied late);
    // a document whose column widths never settle must not paginate forever
    const chain = new FollowUpChain()
    let selfScheduled = false
    // set by the cleanup: document.fonts.ready and other promise callbacks of
    // this run can still fire after the effect was re-created (zoom change)
    // and would schedule a pass of this stale closure next to the live one
    let disposed = false
    let retrigger: string[] = []
    // several reasons in one pass debounce into one follow-up; decoration-only
    // re-decisions (justify shrink, float flow) arrive after a pass and re-run
    // it from the changed block, through the same chain
    const followUpAt = (why: string, dirtyIndex: number | null) => {
      retrigger.push(why)
      if (chain.request(why, dirtyIndex) !== 'run') return
      selfScheduled = true
      onUpdate(null, dirtyIndex, true)
    }
    const followUp = (why: string) => followUpAt(why, null)
    // anchored drawings that get their size after a pass (picture decode, late
    // group layout) re-run it from their block; the preview never prints a pass
    // whose floats were measured before the boxes existed
    let passGen = 0
    const floatWatch = new FloatWatch((changed) => {
      const pm = pmEl()
      let top: Element | null = changed[0]
      while (top && top.parentElement && top.parentElement !== pm) top = top.parentElement
      const bi = lastBlocks.findIndex((b) => b.el === top)
      followUpAt('float-geometry', bi >= 0 ? bi : null)
    })
    let secWidthSig = ''
    let charSpaceSig = ''
    const colWidthPass = newWidthPassState()
    const pmEl = () => document.querySelector('.editor-scroll .ProseMirror') as HTMLElement | null
    const locate = () => {
      const pm = pmEl()
      if (!pm || slices.length === 0) return
      const pmRect = pm.getBoundingClientRect()
      const scrollRect = scroller.getBoundingClientRect()
      const origin = pmRect.top + mTopPx * factor
      // Multiple Pages: the active page is the one under the live window, probe its
      // first content line instead of the viewport middle
      const mpIdx = multipageActiveRef.current
      const mpFrame = mpIdx == null ? null : mpFramesRef.current[mpIdx]
      const wrapTop = pm.closest('.page-wrap')?.getBoundingClientRect().top ?? pmRect.top
      const midScreen = mpFrame
        ? wrapTop + (mpFrame.top + mTopPx + 1) * factor
        : Math.min(scrollRect.top + scrollRect.height / 2, pmRect.bottom)
      // slices use gapless virtual coordinates; subtract the page-gap height above the midpoint
      // (including mid-paragraph inline gaps and repeated-header clone rows, which carry
      // page-repeat-header but not page-gap)
      // page-turn gaps only, like the page frames: a split floating table's
      // in-table gaps count (they are page turns) while its carry spacer does
      // not — the two sums agree past the spacer, and inside the float's pages
      // only the former maps the midpoint onto the page shown there. Its
      // repeated-header clones are float content (the carry covers only the
      // gap bands; the continuation's page starts at the gap's bottom)
      const fromDom = () => {
        let gapAbove = 0
        // mixed-column pages pull their gap (and everything below) up over the
        // vacated stacked space: add that back to land in the gapless coordinates
        let pullUpAbove = 0
        let insideGap = false
        for (const gap of pm.querySelectorAll<HTMLElement>(
          `${PAGE_TURN_SEL}, .page-repeat-header`,
        )) {
          if (gap.classList.contains('page-repeat-header') && insideFloatTable(gap)) continue
          const r = gap.getBoundingClientRect()
          if (r.top < midScreen) {
            gapAbove += Math.min(r.height, midScreen - r.top)
            pullUpAbove += Math.max(0, -parseFloat(gap.style.marginTop) || 0)
            if (midScreen < r.bottom && isPageTurn(gap)) insideGap = true
          }
        }
        // a midpoint inside a gap band lands on the band's top, i.e. the page
        // boundary itself: nudge it onto the page below (sub-pixel ties)
        const y = (midScreen - origin - gapAbove) / factor + pullUpAbove
        return pageAt(slices, insideGap ? y + 1 : y)
      }
      // the page whose gap box starts at or above the midpoint (inside a gap band
      // the page below it is current, where the DOM sum lands too)
      const fromCuts = () => {
        let page = 1
        for (const e of pagePaperEdges(pageCuts, mTopPx)) {
          if (pmRect.top + e.gapTop * factor > midScreen) break
          page = e.page + 1
        }
        return page
      }
      // without gap widgets (web / read mode) the pages follow the flow Y alone
      const shadowPage = pageReaderShadow.enabled ? fromDom() : null
      const page = pageCuts.some((c) => c.gap) ? fromCuts() : fromDom()
      if (shadowPage !== null && shadowPage !== page)
        pageReaderShadow.report('locate', shadowPage, page)
      // visible-page numbering so the status bar and F9 NUMPAGES agree with the gap widgets
      const current = visiblePageCount(slices, page)
      const total = visiblePageCount(slices)
      setPageInfo((prev) =>
        prev.current === current && prev.total === total ? prev : { current, total },
      )
    }
    let lastPassAt = 0
    let lastPassMs = 0
    let lastBlocks: BlockBox[] = []
    let lastPre: PageSlice[] = []
    let lastOut: SliceOutputs | null = null
    let lastSecSig = ''
    /** top-level child count at the last pass: a streaming pass whose only
     *  trigger is the appended tail (dirty ≥ this) may carry the prefix's
     *  line/row samples; any user edit below the frontier (dirty < this)
     *  disables the carry — a same-height edit would otherwise keep stale
     *  line boxes that the height gate cannot see */
    let lastPassChildCount = 0
    let carriedSamples = 0
    /** first top-level index a transaction touched since the last pass; null once a trigger needs the whole document */
    let dirtyFrom: number | null = null
    let resumePasses = 0
    let resumeReject = ''
    let resumePage = -1
    let resumeDiag: Record<string, unknown> = {}
    /** top-level indices edited since the last pass; null once any other trigger disqualified the fast path */
    let fastEdits: Set<number> | null = new Set()
    let fastPasses = 0
    let fastReject = ''
    let gapBoundaryMismatch = 0
    // I11 shadow: PageCut facts vs what the gap builder re-derives from the DOM
    let pageCutMismatch = 0
    const pageCutDiag: string[] = []
    let pageCuts: PageCut[] = []
    let planDiff: PlanDiff[] = []
    const sameBox = (a: BlockBox, b: BlockBox) =>
      a.el === b.el &&
      Math.abs(a.top - b.top) < 0.01 &&
      Math.abs(a.height - b.height) < 0.01 &&
      a.section === b.section &&
      a.docxIndex === b.docxIndex &&
      !!a.floated === !!b.floated &&
      !!a.emptyPara === !!b.emptyPara &&
      (a.spaceBeforePx ?? 0) === (b.spaceBeforePx ?? 0)
    // pages above the first changed block paginate as before: slice again from
    // the page before the one holding that block (keep-with-next and widow
    // rules reach back at most one page), provided that page opens on a block top
    const resumeFor = (
      blocks: BlockBox[],
      secSig: string,
      dirty: number | null,
    ): PassResume | null => {
      const reject = (why: string) => {
        resumeReject = why
        return null
      }
      if (dirty === null) return reject('trigger')
      if (!editor || !lastOut || lastPre.length < 2 || lastBlocks.length === 0)
        return reject('state')
      if (secSig !== lastSecSig) return reject('sections')
      if (viewMode === 'print' && hasVertical) return reject('layout')
      const n = Math.min(blocks.length, lastBlocks.length)
      let d = 0
      while (d < n && sameBox(blocks[d], lastBlocks[d])) d++
      const geomDiff = d
      const { doc: pmDoc } = editor.state
      let trBlock = -1
      if (dirty < pmDoc.childCount) {
        let pos = 0
        for (let i = 0; i < dirty; i++) pos += pmDoc.child(i).nodeSize
        const el = editor.view.nodeDOM(pos)
        const bi = blocks.findIndex((b) => b.el === el)
        if (bi < 0) return reject('block')
        trBlock = bi
        d = Math.min(d, bi)
      }
      const a = blocks[geomDiff]
      const b = lastBlocks[geomDiff]
      resumeDiag = {
        dirty,
        trBlock,
        geomDiff,
        geomField:
          a && b
            ? (
                [
                  'el',
                  'top',
                  'height',
                  'section',
                  'docxIndex',
                  'floated',
                  'emptyPara',
                  'spaceBeforePx',
                ] as const
              ).find((f) =>
                f === 'top' || f === 'height'
                  ? Math.abs((a[f] ?? 0) - (b[f] ?? 0)) >= 0.01
                  : (a[f] ?? null) !== (b[f] ?? null),
              )
            : 'length',
        counts: [blocks.length, lastBlocks.length],
      }
      if (d <= 0) return reject('top')
      if (d >= blocks.length) d = blocks.length - 1
      const dirtyPage = pageOfY(lastPre, blocks[d].top)
      for (let p = dirtyPage - 1; p > 0; p--) {
        if (!isResumePage(lastPre, p)) continue
        const page = lastPre[p]
        let k = -1
        for (let i = 0; i < d; i++) {
          if (Math.abs(blocks[i].top - page.start) < 0.5) {
            k = i
            break
          }
        }
        if (k < 0 || blocks[k].floated) continue
        // line samples and row boxes of the blocks above are geometry of the
        // unchanged pages: carry them over instead of sampling the DOM again
        for (let i = 0; i < k; i++) {
          const prev = lastBlocks[i]
          if (prev.lineBoxes) blocks[i].lineBoxes = prev.lineBoxes
          if (prev.lineLeadPx !== undefined) blocks[i].lineLeadPx = prev.lineLeadPx
          if (prev.tableRows) blocks[i].tableRows = prev.tableRows
          if (prev.colWraps) blocks[i].colWraps = prev.colWraps
          if (prev.oversizeLineH !== undefined) blocks[i].oversizeLineH = prev.oversizeLineH
        }
        resumePage = p
        return { prevSlices: lastPre, page: p, block: k, prevOut: lastOut }
      }
      return reject('page')
    }
    // a keystroke that leaves its paragraph the same height moves no page
    // boundary: keep the previous slicing instead of re-measuring the document
    const tryFastPass = (edited: Set<number>): boolean => {
      const reject = (why: string) => {
        fastReject = why
        return false
      }
      if (!editor || isPhasedContentPending() || lastBlocks.length === 0) return reject('state')
      // column/vAlign placement only translates blocks (heights compare as
      // measured); vertical-text blocks are shown as writing-mode boxes
      if (viewMode === 'print' && hasVertical) return reject('layout')
      const { doc: pmDoc } = editor.state
      const items: FastPassBlock[] = []
      let pos = 0
      let i = 0
      for (const idx of [...edited].sort((a, b) => a - b)) {
        if (idx >= pmDoc.childCount) return reject('index')
        for (; i < idx; i++) pos += pmDoc.child(i).nodeSize
        const el = editor.view.nodeDOM(pos) as HTMLElement | null
        const prev = el ? lastBlocks.find((b) => b.el === el) : undefined
        if (!el || !prev) return reject('block')
        if (
          el.classList.contains('doc-para-join-head') ||
          el.classList.contains('doc-para-join-tail')
        )
          return reject('join')
        items.push({ el, prev })
      }
      return blocksKeepSlicing(items, slices, factor) || reject('geometry')
    }
    const remeasure = () => {
      if (disposed) return
      // a pass during a PDF export re-slices under the print media and rebuilds
      // the preview; the export keeps its page list, so the new sheets print in
      // every later group
      if (scroller.closest('.app.pv-exporting')) {
        timer = window.setTimeout(remeasure, EXPORT_PASS_RETRY_MS)
        return
      }
      timer = null
      lastPassAt = performance.now()
      const edited = fastEdits
      fastEdits = new Set()
      const dirty = dirtyFrom
      dirtyFrom = Infinity
      try {
        const zoomEl = scroller.querySelector<HTMLElement>('.doc-zoom')
        // the canvas zoom as applied, not this closure's: a pass reaching the
        // DOM through a late async callback must never read a zoomed canvas at
        // factor 1 (split-row fills then land in zoomed px and rows overlap)
        const liveZoom = canvasZoomOf(zoomEl, zoom / 100)
        const swapZoom = zoomEl !== null && Math.abs(liveZoom - 1) > 0.0001
        const scrollTop = scroller.scrollTop
        const scrollLeft = scroller.scrollLeft
        if (swapZoom) {
          zoomEl.style.zoom = '1'
          factor = 1
        }
        let fast = false
        try {
          fast = Boolean(edited && edited.size > 0 && tryFastPass(edited))
          if (!fast) {
            if (!edited) fastReject = 'trigger'
            else if (edited.size === 0) fastReject = 'none'
            remeasurePass(dirty)
          }
        } finally {
          if (swapZoom) {
            zoomEl.style.zoom = String(liveZoom)
            factor = liveZoom
            scroller.scrollTop = scrollTop
            scroller.scrollLeft = scrollLeft
          }
        }
        if (fast) {
          fastPasses++
          const dbg = (window as unknown as Record<string, Record<string, unknown>>).__pageDebug
          if (dbg) dbg.fastPasses = fastPasses
        }
        locate()
      } finally {
        lastPassMs = performance.now() - lastPassAt
        // a pass that threw between the add and the pre-setPageGaps remove
        // must not leave the gap widgets hidden
        pmEl()?.classList.remove('measuring-natural')
      }
    }
    const remeasurePass = (dirty: number | null) => {
      const pm = pmEl()
      if (!pm) return
      if (selfScheduled) chain.beginPass()
      else chain.reset()
      layoutHealth.beginPass(!selfScheduled)
      selfScheduled = false
      retrigger = []
      const tStart = performance.now()
      let tMeasure = 0
      let tSlice = 0
      let sliceIters = 0
      let sliceSplit = ''
      // consecutive anchor-paragraph runs collapse onto their band union before
      // measurement (layout-affecting, idempotent)
      syncAnchorBands(pm, factor, (doc?.parsed.compatibilityMode ?? 0) >= 15)
      // later blocks were read before this pass's writes landed: one more
      // pass settles the chain (FollowUpChain refuses a repeating cause)
      const planOf = blockPlans()
      const obstacleShadow = newObstacleShadow()
      pageReaderShadow.enabled =
        import.meta.env.DEV ||
        (window as unknown as { __pageReaderShadow?: boolean }).__pageReaderShadow === true
      // the previous pass's block leads: where its gap widgets stand in the DOM
      const prevLayout = pageLayoutRef.current
      const pageLeads = prevLayout
        ? pageLeadsOf(prevLayout.cuts, prevLayout.blocks)
        : new Map<Element, number>()
      // header drawings reaching into the body of each page of the previous
      // pass (its cuts name the variant and section a page shows)
      const headerObstacleCache = new Map<number, ReturnType<typeof hfBodyObstacles>>()
      const headerObstaclesOf = (pageIdx: number) => {
        const hit = headerObstacleCache.get(pageIdx)
        if (hit) return hit
        const cut = prevLayout?.cuts[pageIdx]
        const prevSecs = lastSecListRef.current
        const si = prevSecs ? Math.min(cut?.section ?? 0, prevSecs.length - 1) : 0
        const set = prevSecs ? prevSecs[si]?.settings : section
        let out: ReturnType<typeof hfBodyObstacles> = []
        if (cut && set) {
          const end = prevSecs?.[si]?.lastBlockIndex
          const hsi = end === undefined ? -1 : hfSections.findIndex((s) => s.lastBlockIndex === end)
          const images = hfResolveAt(hsi >= 0 ? hsi : si, 'header', cut.hf.variant).images
          out = hfBodyObstacles(
            images,
            hfStripGeom(set, legacyBands),
            effectiveTopPx(set, cut.hf.headerPx),
          )
        }
        headerObstacleCache.set(pageIdx, out)
        return out
      }
      const obstaclesChanged = syncWrapObstacles(
        pm,
        factor,
        editor.view,
        planOf,
        obstacleShadow,
        pageLeads,
        headerObstaclesOf,
      )
      // natural-flow measuring state: in-paragraph gap widgets force a line break
      // at the PREVIOUS pass's cut, so measuring with them in layout re-confirms
      // that cut even after decoration-only reflows (justify space-shrink) moved
      // the natural wrap points — a justified page-last line then stayed
      // stretched sparse until an at-boundary edit (r177). Hidden, sampling AND
      // cut-anchor resolution below see the natural wrap; removed before
      // setPageGaps so widget building and later display-state reads (strip
      // alignment, cut overlays, locate) run against the real layout.
      pm.classList.add('measuring-natural')
      // joined paragraphs align against the natural flow (gap widgets hidden)
      alignJoinedParagraphs(pm, factor, editor.view)
      // a columned canvas measures + slices in the single-flow measuring state (fillLineBoxes
      // also reads the DOM for line sampling, so it must share the state); display-state DOM
      // reads like gap positioning happen outside the measuring state
      const measured = measureSingleFlow(pm, () => {
        const t0 = performance.now()
        const origin = pm.getBoundingClientRect().top + mTopPx * factor
        const { blocks, totalHeight, floats, sectBreaks } = measureBlocks(
          pm,
          origin,
          factor,
          planOf,
          true,
          floatCarriesOf(editor.state),
        )
        planDiff = planOf
          ? [
              ...shadowPlanDiff(planOf, blocks, blockMetaOf),
              ...obstaclePlanDiffs(obstacleShadow.diffs, blocks, planOf),
            ]
          : []
        carriedSamples += carryStreamedSamples(blocks, lastBlocks, {
          pending: isPhasedContentPending(),
          dirty,
          lastPassChildCount,
        })
        tMeasure = performance.now() - t0
        // rowSplits enables the per-cell row split; the canvas only needs the
        // matching row heights (the preview applies the cell shifts)
        const sliceOut: SliceOutputs = {
          rowFills: [],
          rowSplits: [],
          floatVShifts: [],
          floatPullUps: [],
          floatCarries: [],
          floatFlows: [],
          floatSplits: [],
          oversizeClips: [],
        }
        const passCtx: PassContext = {
          sections,
          section,
          delSectBreaks,
          endnoteItems,
          endnoteSeparatorPx: noteSeparatorPx('endnote'),
          hfHeightsOf,
          singleHfPx,
          singleGeom: { contentHeight: contentH, forceBreak: false, ...singleVariantGeom },
          evenOddHf,
          titlePg,
          // when the canvas column layout is inactive, measure as full-width single flow
          dropCols: colMode === 'none' || viewMode !== 'print',
          blockMetaOf,
          ...(hasVertical ? { inlineExtraOf: blockInlineExtraPx } : {}),
          factor,
          resumeFor: (bs, secSig) => resumeFor(bs, secSig, dirty) ?? undefined,
          out: sliceOut,
        }
        const flow = prepareFlow({ blocks, totalHeight, floats, sectBreaks }, passCtx)
        const secList = flow.secList
        lastSecListRef.current = secList
        const hfHs = secList ? flow.hfHeights : null
        const geoms = flow.geoms
        const t1 = performance.now()
        const { slices: s, cuts, resumed } = sliceFlow(flow, passCtx)
        if (resumed) resumePasses++
        tSlice = performance.now() - t1
        lastPre = sliceOut.preParity ?? []
        lastOut = sliceOut
        lastSecSig = flow.secSig
        sliceIters = sliceOut.iterations ?? 0
        sliceSplit = `${Math.round(sliceOut.fillMs ?? 0)}+${Math.round(sliceOut.sliceRunMs ?? 0)}`
        return {
          blocks,
          secList,
          hfHs,
          s,
          cuts,
          geoms,
          flow,
          floats,
          rowSplits: sliceOut.rowSplits ?? [],
          rowFills: sliceOut.rowFills ?? [],
          floatVShifts: sliceOut.floatVShifts ?? [],
          floatPullUps: sliceOut.floatPullUps ?? [],
          floatCarries: sliceOut.floatCarries ?? [],
          floatFlows: sliceOut.floatFlows ?? [],
          floatSplits: sliceOut.floatSplits ?? [],
          oversizeClips: sliceOut.oversizeClips ?? [],
        }
      })
      const {
        blocks,
        secList,
        hfHs,
        floats,
        rowFills,
        floatVShifts,
        floatPullUps,
        floatCarries,
        floatFlows,
        floatSplits,
        oversizeClips,
        cuts,
      } = measured
      slices = measured.s
      pageCuts = cuts
      lastBlocks = blocks
      lastPassChildCount = editor?.state.doc.childCount ?? lastPassChildCount
      const pageNotes = pageFootnotesOf(blocks, slices)
      passGen++
      const floatWatchRec = floatWatch.record(pm)
      pageLayoutRef.current = {
        passGen,
        blocks,
        slices,
        sections: secList ?? [],
        cuts,
        geoms: measured.geoms,
        hfHeights: hfHs ?? [singleHfPx],
        rowSplits: measured.rowSplits,
        carries: [],
        floats,
        floatPullUps,
        flow: measured.flow,
        pageNotes,
      }
      if (mirrorMargins) locateCaretPageRef.current()
      const blockIndex = new BlockIndex(blocks)
      // document-end footer shows the last page's displayed number, not the physical count
      if (slices.length > 0) {
        const lastIdx = slices.length - 1
        const num = secList ? pageNumbers(slices, secList)[lastIdx] : slices.length
        setLastPageNo({ num, text: cuts[lastIdx]?.display ?? String(num) })
        setFirstPageNo(cuts[0]?.display ?? null)
        const counts = sectionPageCounts(slices)
        const first = counts[0]
        const last = counts[lastIdx]
        setEdgeSectionPages((prev) =>
          prev?.first === first && prev.last === last ? prev : { first, last },
        )
      }
      // Auto-refresh TOC page numbers from the fresh slicing, like Word's field
      // update. Gated on dirtyRef — our pagination approximates Word's, so a
      // pristine open keeps the file's numbers. History-exempt: a field result
      // change is not an edit to undo.
      if (editor && dirtyRef.current && slices.length > 0) {
        const byEl = new Map(blocks.filter((b) => b.el).map((b) => [b.el as HTMLElement, b.top]))
        const headings = collectHeadings(editor.state.doc, editor.storage.listNumbering?.styles)
        // as the header/footer numbers show it: the section's format, and its chapter number
        const displays = headings.map((h) => {
          const dom = editor.view.nodeDOM(h.pos) as HTMLElement | null
          const top = dom ? byEl.get(dom) : undefined
          if (top === undefined) return undefined
          const idx = Math.min(Math.max(pageOfY(slices, top) + 1, 1), slices.length)
          return cuts[idx - 1]?.display
        })
        const tr = editor.state.tr
        if (applyTocPageDisplays(editor.state.doc, tr, headings, displays)) {
          tr.setMeta('addToHistory', false)
          const before = editor.state.doc
          editor.view.dispatch(tr)
          // a field refresh is not a user edit: keep the AI freshness baseline
          carryDocSeen(editor, before)
        }
      }
      // M4 always-on pagination in the canvas: render page gaps before page-leading blocks
      // (print view only; gaps don't count as content — measureBlocks subtracts them, so
      // slice results are gap-independent and refresh is idempotent)
      let tGapsBuild: number | undefined
      let tSetGaps: number | undefined
      let tCommit: number | undefined
      const tGaps0 = performance.now()
      if (editor) {
        const gaps: PageGapSpec[] = []
        // in-table gap height per split floating table (its anchor paragraph
        // skips the same distance to start beside the last portion)
        const carryGaps = new Map<number, number>()
        const overlayCutAnchors: LineAnchor[] = []
        const nestedGaps: NestedTableGap[] = []
        const overlayPages: number[] = []
        const carries: FlowCarry[] = []
        const lineRectsOf = createLineRectsCache()
        const gapIds = new Set<string>()
        let firstPageFloats: { els: HTMLElement[]; key: string } | undefined
        if (viewMode === 'print' && !readMode) {
          const pmRect = pm.getBoundingClientRect()
          // per-page header/footer for the gaps (previous page's footer + next page's
          // header), variant-selected like the pagination preview (first page / odd-even /
          // per-section references), with real page numbers for the '#' marker
          const parsed = doc.parsed
          // floating shapes (watermarks) must not stack into the strip (mirrors
          // PaginationPreview); they render separately as per-page behind-text images
          // measured and document-derived section lists agree except mid-transition:
          // bridge them by the break paragraph so the strip matches what an edit targets
          const resolveLive = (si: number, kind: 'header' | 'footer', variant: HfView) => {
            const end = secList?.[si]?.lastBlockIndex
            const hsi =
              end === undefined ? -1 : hfSections.findIndex((s) => s.lastBlockIndex === end)
            return hfResolveAt(hsi >= 0 ? hsi : si, kind, variant)
          }
          const pageHfOf = (
            pageIdx: number,
            kind: 'header' | 'footer',
          ): {
            value: HeaderFooter | null
            images?: HfImage[]
            floats: HfImage[]
            section: number
            variant: HfView
            owner: number
          } => {
            // row-anchored logos render inside the strip (off their anchor row), not on the float layer
            const split = (imgs?: HfImage[] | null) => ({
              images: imgs?.filter((img) => !img.floating || hfImageHangsOnPara(img)),
              floats: [
                ...(kind === 'header' && pendingWatermarkImage ? [pendingWatermarkImage] : []),
                ...(imgs?.filter(
                  (img) =>
                    img.floating &&
                    !hfImageHangsOnPara(img) &&
                    !(watermarkDirty && (img.wordArt || img.watermark)),
                ) ?? []),
              ],
            })
            const si = secList ? Math.min(slices[pageIdx].section, secList.length - 1) : 0
            const variant: HfView = cuts[pageIdx].hf.variant
            const resolved = resolveLive(si, kind, variant)
            return {
              value: resolved.value,
              ...split(resolved.images),
              section: si,
              variant,
              owner: resolved.owner,
            }
          }
          /** reserved strip heights of the variant a page shows (capacity and gap band agree) */
          const hfOfPage = (idx: number): { headerPx: number; footerPx: number } => cuts[idx].hf
          /** width a strip is laid out at: the declaring section's column (hfPartWidthPx) */
          const stripColumnW = (hf: ReturnType<typeof pageHfOf>, kind: 'header' | 'footer') =>
            secList
              ? { columnW: hfPartWidthPx(secList, hf.section, kind, hf.variant, hf.owner) }
              : {}
          // Like Word there is no variant switcher: a header/footer part is edited on a
          // page that shows it. The edge areas cover page 1's header and the last page's
          // footer; a gap strip for any other (section, variant) pair is kept even while
          // the part is empty so it still has a double-click target (an empty even-page
          // header, a first-page footer on a multi-page document, ...)
          const edgeHeaderHf = slices.length > 0 ? pageHfOf(0, 'header') : null
          const edgeFooterHf = slices.length > 0 ? pageHfOf(slices.length - 1, 'footer') : null
          const gapStripNeeded = (
            gap: ReturnType<typeof pageHfOf>,
            edge: ReturnType<typeof pageHfOf> | null,
          ) =>
            hfHasVisibleContent(gap.value, gap.images) ||
            gap.variant !== edge?.variant ||
            gap.section !== edge.section
          /** page geometry a page's floating header images position against */
          const floatBoxOf = (pageIdx: number): HfFloatBox => {
            const s = secList?.[slices[pageIdx].section]?.settings ?? section
            const hfH = hfOfPage(pageIdx)
            const side = pageMargins(s, cuts[pageIdx].pageNo, mirrorMargins)
            return {
              pageW: twipsToPx(s.pageWidth),
              pageH: twipsToPx(s.pageHeight),
              marginLeft: twipsToPx(side.left),
              marginRight: twipsToPx(side.right),
              marginTop: effectiveTopPx(s, hfH.headerPx),
              marginBottom: effectiveBottomPx(s, hfH.footerPx),
              headerDist: twipsToPx(s.headerDist ?? 720),
              sectMarginTop: twipsToPx(s.marginTop),
              sectMarginBottom: twipsToPx(s.marginBottom),
              paperX: pageLeftOf(s),
              hostLeft: twipsToPx(s.marginLeft),
            }
          }
          const pageNoTextOf = (pageIdx: number) => cuts[pageIdx].display
          const hfSig = (v: HeaderFooter | null | undefined) =>
            v ? `${v.text}·${v.pageNumber ? 1 : 0}·${v.paras?.length ?? 0}` : ''
          // identity + placement of floating header images: widget keys must change
          // when the watermark image or its position changes, not just its count
          // (dataUrl hashed over the edges only — enough to tell images apart
          // without walking megabytes of base64 per page per rebuild)
          const floatSig = (imgs: HfImage[]) =>
            imgs
              .map(
                (f) =>
                  `${f.dataUrl.length}:${hashStr(f.dataUrl.slice(0, 1024) + f.dataUrl.slice(-1024))}:${f.posXPx ?? f.posH ?? ''}:${f.posYPx ?? f.posV ?? ''}:${f.posHRel ?? ''}${f.posVRel ?? ''}:${f.widthPx ?? ''}x${f.heightPx ?? ''}${f.washout ? ':w' : ''}${f.rotationDeg ? `:r${f.rotationDeg}` : ''}${f.wordArt ? `:t${JSON.stringify(f.wordArt)}` : ''}`,
              )
              .join('|')
          const visiblePages = visiblePageCount(slices)
          const sectionPages = sectionPageCounts(slices)
          const canvasPaperW = pmRect.width / factor
          const canvasSet = secList?.[0]?.settings ?? section
          const canvasContentWPx = canvasSet
            ? twipsToPx(canvasSet.pageWidth - canvasSet.marginLeft - canvasSet.marginRight)
            : 0
          // mirrored documents pin the strips too: even pages read the swapped margins
          const mixedWidths =
            mixedPaper ||
            mirrorMargins ||
            (canvasSet != null &&
              (secList ?? []).some(
                (s) =>
                  Math.abs(
                    twipsToPx(
                      s.settings.pageWidth - s.settings.marginLeft - s.settings.marginRight,
                    ) - canvasContentWPx,
                  ) > 0.5,
              ))
          // equal-width docs: strip centered, clamped to the canvas paper. Differing-width
          // docs: strip gets its section's width and is aligned to the body blocks' left
          // edge afterwards by measurement (alignGapHfStrips) — gap-box origins vary per
          // gap kind, so no static left works here
          const sizeGapHf = (el: HTMLElement, box: { contentWidth: number }) => {
            if (mixedWidths) {
              el.style.width = `${box.contentWidth}px`
              el.style.left = '0px'
              el.style.transform = 'none'
            } else {
              el.style.width = `${Math.min(box.contentWidth, canvasPaperW)}px`
            }
          }
          // page x of a gap strip's left edge, mirroring sizeGapHf (centered, or on the body's left margin)
          const gapStripLeft = (
            set: SectionSettings,
            box: { contentWidth: number },
            pageIdx: number,
          ) =>
            mixedWidths
              ? twipsToPx(pageMargins(set, cuts[pageIdx].pageNo, mirrorMargins).left)
              : (canvasPaperW - Math.min(box.contentWidth, canvasPaperW)) / 2
          // paper x of a page's body left edge: its section's page offset plus the
          // page's own (mirrored) left margin
          const bodyLeftOf = (set: SectionSettings, pageIdx: number) =>
            pageLeftOf(set) + twipsToPx(pageMargins(set, cuts[pageIdx].pageNo, mirrorMargins).left)
          // strip geometry baked in at creation (sizeGapHf, --hf-ml): the gap key's
          // mKey carries the NEXT section's side margins only, while the footer strip
          // is sized and inset by the PREVIOUS section — a left-margin edit on that
          // section alone must not reuse the old footer widget at its stale inset
          const stripGeomSig = (set: typeof canvasSet, pageIdx: number) => {
            if (!set) return ''
            const b = sectionPageBox(set)
            return [bodyLeftOf(set, pageIdx), b.contentWidth, b.headerDist, b.footerDist]
              .map((v) => v.toFixed(1))
              .join(',')
          }
          slices.slice(1).forEach((slice, k) => {
            // a same-start predecessor that is zero-height is a deliberate blank page
            // (leading/double w:br, even/odd parity): it needs its own gap band so the
            // blank sheet paints (pad below covers its full paper height); other
            // same-start duplicates draw only one band
            if (slice.start === slices[k].start && slices[k].end > slices[k].start) return
            // gap = previous page's (its section's) bottom margin + inter-page band + this page's (its section's) top margin
            const prevSec = secList?.[slices[k].section]?.settings ?? section
            const nextSec = secList?.[slice.section]?.settings ?? section
            // effective margins after header/footer push-down (an over-tall header pushes
            // the body down); a titlePg section's first page uses the first-page variant's
            // heights so the gap band matches the slice capacity (firstContentHeight)
            const nextHf = hfOfPage(k + 1)
            const prevHf = hfOfPage(k)
            const metrics = {
              marginTop: effectiveTopPx(nextSec, nextHf.headerPx),
              marginBottom: effectiveBottomPx(prevSec, prevHf.footerPx),
              marginLeft: twipsToPx(nextSec.marginLeft),
              marginRight: twipsToPx(nextSec.marginRight),
              // header/footer strip inset on the paper (alignGapHfStrips): survives
              // the inline/in-cell gaps below overriding marginLeft/Right with the host
              // block's paper offset (which folds in paragraph indents and cell positions)
              sectionMarginLeft: bodyLeftOf(nextSec, k + 1),
              sectionMarginRight: twipsToPx(nextSec.marginRight),
              sectionMarginTop: twipsToPx(nextSec.marginTop),
              ...(mixedPaper
                ? { pageLeft: pageLeftOf(nextSec), pageWidth: twipsToPx(nextSec.pageWidth) }
                : {}),
            }
            // a page ended early (explicit break / section break / keepNext) leaves unused
            // content height (PageCut.padPx); pad the gap so the canvas paints the full paper
            // height and the footer stays at the paper bottom. Uniform multi-column pages span
            // columns × height with the browser compressing the flow, skip; mixed-column pages
            // use the engine's physical height and pull the gap up over the vacated stacked space.
            const prevCut = cuts[k]
            const cut = cuts[k + 1]
            const stacked = !slices[k].regions || colMode === 'mixed'
            const items = pageNotes[k] ?? []
            const fnH =
              items.length > 0
                ? items.reduce((s, n) => s + n.height, 0) + noteSeparatorPx('footnote')
                : 0
            const remaining = stacked ? prevCut.padPx : 0
            const pullUp = stacked ? Math.max(0, prevCut.usedPx - prevCut.physUsedPx) : 0
            // used excludes the reserved footnote height (footnoteExtraPx inflates capacity
            // bookkeeping, not DOM coordinates), and the notes area already extends the gap
            // by fnH: pad covers only the rest of the shortfall
            const pad = Math.max(0, Math.round(remaining - fnH))
            // previous page's footer (bottom-margin band) + next page's header (top-margin
            // band), so the canvas shows headers/footers on every page like Word
            const gapFooter = pageHfOf(k, 'footer')
            const gapHeader = pageHfOf(k + 1, 'header')
            const hfEls: HTMLElement[] = []
            if (gapStripNeeded(gapFooter, edgeFooterHf)) {
              const box = sectionPageBox(prevSec)
              const el = makeGapHfEl({
                kind: 'footer',
                value: gapFooter.value ?? { text: '' },
                images: gapFooter.images,
                pageNo: pageNoTextOf(k),
                pageTotal: visiblePages,
                sectionPages: sectionPages[k],
                geom: {
                  ...hfStripGeom(prevSec, legacyBands),
                  stripLeft: gapStripLeft(prevSec, box, k),
                  ...stripColumnW(gapFooter, 'footer'),
                },
              })
              Object.assign(el.dataset, {
                hfKind: 'footer',
                hfSection: String(gapFooter.section),
                hfVariant: gapFooter.variant,
                hfPageIdx: String(k),
                hfPage: pageNoTextOf(k),
                hfTotal: String(visiblePages),
                hfSectionPages: String(sectionPages[k]),
              })
              el.style.top = 'auto'
              el.style.bottom = `${GAP_BAND + metrics.marginTop + box.footerDist}px`
              sizeGapHf(el, box)
              // the footer belongs to the page ABOVE the gap: its own section's
              // left margin, not the next section's (alignGapHfStrips)
              el.style.setProperty('--hf-ml', `${bodyLeftOf(prevSec, k)}px`)
              hfEls.push(el)
            }
            if (gapStripNeeded(gapHeader, edgeHeaderHf)) {
              const box = sectionPageBox(nextSec)
              const el = makeGapHfEl({
                kind: 'header',
                value: gapHeader.value ?? { text: '' },
                images: gapHeader.images,
                pageNo: pageNoTextOf(k + 1),
                pageTotal: visiblePages,
                sectionPages: sectionPages[k + 1],
                geom: {
                  ...hfStripGeom(nextSec, legacyBands),
                  stripLeft: gapStripLeft(nextSec, box, k + 1),
                  ...stripColumnW(gapHeader, 'header'),
                },
              })
              Object.assign(el.dataset, {
                hfKind: 'header',
                hfSection: String(gapHeader.section),
                hfVariant: gapHeader.variant,
                hfPageIdx: String(k + 1),
                hfPage: pageNoTextOf(k + 1),
                hfTotal: String(visiblePages),
                hfSectionPages: String(sectionPages[k + 1]),
              })
              el.style.bottom = 'auto'
              // the strip sits at w:headerDist even under a fixed top margin above it: Word lets
              // the body overprint the header there (probe 2026-10-30), as the first page draws it
              el.style.top = `calc(100% + ${box.headerDist - metrics.marginTop}px)`
              sizeGapHf(el, box)
              el.style.setProperty('--hf-ml', `${bodyLeftOf(nextSec, k + 1)}px`)
              hfEls.push(el)
            }
            // next page's floating header images (picture watermarks): behind-text,
            // positioned from that page's origin (the gap's bottom edge is marginTop
            // above it), like PaginationPreview's per-page watermark layer
            for (const img of gapHeader.floats) {
              hfEls.push(
                tagHfFloatEl(makeHfFloatImgEl(img, floatBoxOf(k + 1), 'gap'), gapHeader, k + 1),
              )
            }
            const hfProps =
              hfEls.length > 0
                ? {
                    hfEls,
                    // key must cover everything baked into the widgets (both pages'
                    // formatted numbers + total count, both sections' strip geometry,
                    // and the (section, variant) each strip edits — a Different First
                    // Page / Odd & Even toggle can swap the variant behind identical
                    // text), or stale PAGE/NUMPAGES / strip insets / edit targets
                    // survive reuse
                    hfKey: `${pageNoTextOf(k)}·${pageNoTextOf(k + 1)}·${visiblePages}·${sectionPages[k]}/${sectionPages[k + 1]}·${gapFooter.section}${gapFooter.variant[0]}·${gapHeader.section}${gapHeader.variant[0]}·${hfSig(gapFooter.value)}·${hfSig(gapHeader.value)}·f${floatSig(gapHeader.floats)}·g${mixedWidths ? 1 : 0}:${stripGeomSig(prevSec, k)}:${stripGeomSig(nextSec, k + 1)}·h${docHyphenationSig()}`,
                  }
                : {}
            // previous page's footnotes: rendered into the top of the gap (page-bottom area), with the gap enlarged by the reserved height.
            // in-table gaps carry no footnote area (absolute positioning inside a table-row is unreliable); footnotes stay in the end-of-document list
            let notes: HTMLElement | undefined
            let notesKey: string | undefined
            // the page's text area (footnotes excluded): dropped pictures fit above its bottom
            const prevContentH =
              twipsToPx(prevSec.pageHeight) -
              effectiveTopPx(prevSec, prevHf.headerPx) -
              metrics.marginBottom
            let notesMetrics: GapMetrics = { ...metrics, bodyHeight: prevContentH - fnH }
            if (items.length > 0) {
              const contentW = twipsToPx(
                prevSec.pageWidth - prevSec.marginLeft - prevSec.marginRight,
              )
              const notesLeft = bodyLeftOf(prevSec, k)
              notes = makeGapNotesEl(
                items,
                notesLeft,
                contentW,
                fnH,
                footnoteLineHeightPx(prevSec.docGrid),
                (id) => editNote('footnote', id),
              )
              notesKey = `${Math.round(fnH)}:${notesLeft.toFixed(1)}:${noteSettings.footnote.numFmt ?? ''}${notesBeneathText ? 'b' : ''}:${items.map((n) => `${n.no}${n.fmt ?? ''}${n.text}`).join('|')}:h${docHyphenationSig()}`
              notesMetrics = {
                ...notesMetrics,
                marginBottom: metrics.marginBottom + fnH,
                bodyMarginBottom: metrics.marginBottom,
              }
            }
            const markShown = () => items.forEach((n) => gapIds.add(n.id))
            // the page lead comes from the PageCut (I11); the DOM only resolves the
            // ProseMirror position / element behind it. A lead the DOM cannot honour
            // counts as a pageCutMismatch (and the dropped gap as a boundary mismatch)
            const lead = cut.lead
            const missCut = (why: string) => {
              pageCutMismatch++
              if (pageCutDiag.length < 20) pageCutDiag.push(`page ${k + 2}: ${why}`)
            }
            if (lead.kind === 'block') {
              const el = blocks[lead.block]?.el
              if (!el) {
                missCut('block lead without element')
                return
              }
              // footnotes sit at the paper bottom (Word): shift past the padding
              if (notes && pad > 0 && !notesBeneathText) notes.style.top = `${5 + pad}px`
              gaps.push({
                el,
                page: k + 1,
                boundaryY: slice.start,
                metrics:
                  pad > 0
                    ? {
                        ...notesMetrics,
                        marginBottom: notesMetrics.marginBottom + pad,
                        bodyPad: pad,
                        bodyMarginBottom: metrics.marginBottom,
                      }
                    : notesMetrics,
                ...(pullUp > 0.5 ? { pullUp } : {}),
                ...(cut.leadDropsSpaceBefore ? { suppressLeadMt: true } : {}),
                ...(cut.leadMtPx !== undefined ? { leadMtPx: cut.leadMtPx } : {}),
                ...(notes ? { notes, notesKey } : {}),
                ...hfProps,
              })
              markShown()
              return
            }
            if (lead.kind === 'tail') {
              // deliberate trailing blank page (document ends with a page break): no
              // anchor block exists — hang the gap at the document end so the blank
              // sheet paints (min-height extension below covers its paper height)
              if (k + 2 === slices.length) {
                gaps.push({
                  pos: editor.state.doc.content.size,
                  kind: 'inline',
                  page: k + 1,
                  boundaryY: slice.start,
                  metrics:
                    pad > 0
                      ? {
                          ...notesMetrics,
                          marginBottom: notesMetrics.marginBottom + pad,
                          bodyPad: pad,
                          bodyMarginBottom: metrics.marginBottom,
                        }
                      : notesMetrics,
                  ...(pullUp > 0.5 ? { pullUp } : {}),
                  ...hfProps,
                })
                markShown()
              }
              return
            }
            const b = blocks[lead.block]
            if (!b?.el) {
              missCut(`${lead.kind} lead without element`)
              return
            }
            const cutOff = lead.lineOffset ?? 0
            if (lead.kind === 'row' || lead.kind === 'cell') {
              // the previous pass's gap rows and repeated headers are not rows of the
              // measured table (domTableRows): index the same list
              const realRows = outerTableRows(b.el)
              if (lead.kind === 'row') {
                const tr = realRows[lead.row ?? -1]
                if (!tr) {
                  missCut(`row ${lead.row} beyond ${realRows.length} DOM rows`)
                  return
                }
                // the I11 shadow of the row sampler: the measured row top must still be
                // where the DOM row sits (a tr's offset includes the in-table gaps and
                // the phantom-row residue above it, like the sampler reads them)
                const gapRects = blockGapRects(b.el)
                const trTop = tr.getBoundingClientRect().top
                const gapsAbove = gapRects.reduce((s, g) => (g.top <= trTop ? s + g.height : s), 0)
                const off = (trTop - b.el.getBoundingClientRect().top - gapsAbove) / factor
                const rowTop = (b.tableRows ?? [])
                  .slice(0, lead.row)
                  .reduce((s, r) => s + r.height, 0)
                if (Math.abs(off - rowTop) >= LEAD_ROW_TOL_PX)
                  missCut(`row ${lead.row} DOM offset ${off.toFixed(1)} vs ${rowTop.toFixed(1)}`)
                // split floating table: the boundary is one of its row cuts, so the
                // page turn is an in-table gap of the float; the anchor paragraph
                // below it skips the same distance (carryGaps)
                const fsplit = floatSplits.find((f) =>
                  f.cutYs.some((y) => Math.abs(y - slice.start) < 0.5),
                )
                try {
                  const $pos = editor.view.state.doc.resolve(editor.view.posAtDOM(tr, 0))
                  // w:tblHeader repetition: the engine reserved slice.repeatHeader.height
                  // at the top of this page's column, so cloning the source header rows
                  // below the gap fills exactly that space. Clones are decorations —
                  // page-gap-inline keeps them out of the virtual coordinates.
                  let repeatHeaderEls: HTMLElement[] | undefined
                  let repeatHeaderPx = 0
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
                    if (els.length > 0) {
                      repeatHeaderEls = els
                      repeatHeaderPx = acc
                    }
                  }
                  // the spanning gap cell must cover exactly the table's column
                  // grid: a wider colSpan adds phantom columns, which collapses
                  // colgroup-less fixed-layout tables to ~1px columns (makeGapEl)
                  const gridCols = tableGridCols(tr.closest('table'))
                  for (let d = $pos.depth; d > 0; d--) {
                    if ($pos.node(d).type.name === 'docTableRow') {
                      // no notes area in table gaps: pad the full remainder
                      const tablePad = Math.round(remaining)
                      gaps.push({
                        pos: $pos.before(d),
                        kind: 'table',
                        cols: gridCols,
                        page: k + 1,
                        boundaryY: slice.start,
                        ...(fsplit ? { floatGap: true } : {}),
                        ...(repeatHeaderPx > 0 ? { repeatHeaderPx } : {}),
                        metrics:
                          tablePad > 0
                            ? {
                                ...metrics,
                                marginBottom: metrics.marginBottom + tablePad,
                                bodyMarginBottom: metrics.marginBottom,
                              }
                            : metrics,
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
                      if (fsplit) {
                        const gapH = metrics.marginBottom + tablePad + GAP_BAND + metrics.marginTop
                        carryGaps.set(fsplit.blockTop, (carryGaps.get(fsplit.blockTop) ?? 0) + gapH)
                      }
                      break
                    }
                  }
                } catch {
                  // posAtDOM failed: skip decorating this round
                  missCut(`row ${lead.row} has no ProseMirror position`)
                }
                return
              }
              // in-row cut point (page break between a cell's lines): anchor at the first line after it
              const cutRow =
                lead.row !== undefined && lead.row >= 0 ? (realRows[lead.row] ?? null) : null
              const anchor = nextLineAnchor(b.el, cutOff, factor, lineRectsOf)
              if (anchor == null) {
                missCut(`no line anchor after the cut in row ${lead.row}`)
                return
              }
              // anchors inside the read-only nested-table NodeView have no distinct PM
              // position (posAtDOM collapses them all to the node start), so no widget
              // can reach them: the gaps go into the view's DOM (syncNestedTableGaps,
              // pos is unused there), and a turn neither helper can place stays an
              // overlay marker
              if (anchorElement(anchor)?.closest('.doc-nested-table')) {
                // a turn between two of its rows: an in-table gap row, padded to the
                // paper bottom like a row cut
                const row = nestedRowAtCut(anchor, NESTED_ROW_TOL_PX * factor)
                const table = row?.closest('table')
                if (row && table) {
                  const tablePad = Math.round(remaining)
                  nestedGaps.push({
                    row,
                    gap: {
                      pos: 0,
                      kind: 'table',
                      cols: tableGridCols(table),
                      page: k + 1,
                      boundaryY: slice.start,
                      metrics:
                        tablePad > 0
                          ? {
                              ...metrics,
                              marginBottom: metrics.marginBottom + tablePad,
                              bodyMarginBottom: metrics.marginBottom,
                            }
                          : metrics,
                      ...hfProps,
                    },
                  })
                  return
                }
                // a turn between a row's lines: Word splits the row, each cell
                // carrying on at the next page's top
                // Word 2013+ cells keep widow/orphan control (as planRowSplit)
                const cellWidow = b.modernTableHeaders === true && !b.cellWidowOff
                const split = nestedRowSplit(
                  anchor,
                  factor,
                  NESTED_ROW_TOL_PX,
                  (cell, off) => nextLineAnchor(cell, off, factor, lineRectsOf),
                  cellWidow,
                )
                if (split) {
                  const last = split.cells.length - 1
                  nestedGaps.push({
                    row: split.row,
                    gap: { pos: 0, kind: 'cell', page: k + 1, boundaryY: slice.start, metrics },
                    cells: split.cells.map(({ mark, liftPx, block }, i) => {
                      const r = block.getBoundingClientRect()
                      return {
                        mark,
                        liftPx,
                        gap: {
                          pos: 0,
                          kind: 'cell' as const,
                          page: k + 1,
                          boundaryY: slice.start,
                          metrics: {
                            ...metrics,
                            marginBottom: metrics.marginBottom + liftPx,
                            bodyMarginBottom: metrics.marginBottom,
                            marginLeft: (r.left - pmRect.left) / factor - pageLeftOf(nextSec),
                            marginRight: (pmRect.right - r.right) / factor - pageLeftOf(nextSec),
                          },
                          // one set of header/footer strips, on the gap painted last
                          ...(i === last ? hfProps : {}),
                        },
                      }
                    }),
                  })
                  return
                }
                overlayCutAnchors.push(anchor)
                overlayPages.push(k + 1)
                return
              }
              const pos = posFromAnchor(editor.view, anchor)
              if (pos == null) {
                missCut(`cell lead has no ProseMirror position`)
                return
              }
              const cutCell = singleCutCell(cutRow, anchor, factor)
              if (cutCell) {
                // single-column row: insert a real inline gap band; bleed to the paper
                // edges from the anchor's block (the widget's containing block)
                const r = (
                  anchorElement(anchor)?.closest('td > *, th > *') ?? cutCell
                ).getBoundingClientRect()
                gaps.push({
                  pos,
                  kind: 'cell',
                  page: k + 1,
                  boundaryY: slice.start,
                  metrics: {
                    ...metrics,
                    // rect offsets from the paper edge already include the page margins
                    marginLeft: (r.left - pmRect.left) / factor - pageLeftOf(nextSec),
                    marginRight: (pmRect.right - r.right) / factor - pageLeftOf(nextSec),
                  },
                  ...(pullUp > 0.5 ? { pullUp } : {}),
                  ...hfProps,
                })
              } else {
                // multi-cell row: keep the zero-height dashed marker (no hfProps — it can't host header/footer strips)
                gaps.push({ pos, kind: 'cut', page: k + 1, metrics })
              }
              return
            }
            // mid-paragraph page break (line-level cut point): insert an inline gap at the
            // broken line. Fallback: font-load reflow can leave lineStartAnchor's exact-offset
            // match just outside tolerance; the first line after the cut still gets the gap
            const anchor =
              lineStartAnchor(b.el, cutOff, factor, lineRectsOf) ??
              nextLineAnchor(b.el, cutOff, factor, lineRectsOf)
            const pos = anchor ? posFromAnchor(editor.view, anchor) : undefined
            if (pos == null) {
              console.warn('[pagination] no line anchor at page boundary', slice.start)
              missCut('no line anchor at the cut')
              return
            }
            // fold the block's offset from the paper edge (page margin + indent) into negative margins so the gap spans exactly the paper width
            const elRect = b.el.getBoundingClientRect()
            gaps.push({
              pos,
              page: k + 1,
              boundaryY: slice.start,
              metrics: {
                ...notesMetrics,
                marginLeft: (elRect.left - pmRect.left) / factor - pageLeftOf(nextSec),
                marginRight: (pmRect.right - elRect.right) / factor - pageLeftOf(nextSec),
              },
              ...(pullUp > 0.5 ? { pullUp } : {}),
              ...(notes ? { notes, notesKey } : {}),
              ...hfProps,
            })
            markShown()
          })
          // first page has no gap widget; its floating header images ride a
          // dedicated zero-height widget at the document start
          if (slices.length > 0) {
            const firstHeader = pageHfOf(0, 'header')
            const floats = firstHeader.floats
            if (floats.length > 0) {
              const box = floatBoxOf(0)
              firstPageFloats = {
                els: floats.map((img) =>
                  tagHfFloatEl(makeHfFloatImgEl(img, box, 'lead'), firstHeader, 0),
                ),
                key: `${floatSig(floats)}·${Math.round(box.pageW)}x${Math.round(box.pageH)}·${Math.round(box.marginLeft)}·${Math.round(box.hostLeft ?? 0)}·${Math.round(box.marginTop)}·${Math.round(box.sectMarginTop)}·${Math.round(box.paperX ?? 0)}`,
              }
            }
          }
          // silently dropped boundaries merge two pages into one giant page — surface them
          const boundaries = slices.slice(1).filter((s, k) => s.start !== slices[k].start).length
          if (gaps.length + nestedGaps.length + overlayCutAnchors.length !== boundaries) {
            gapBoundaryMismatch++
            console.warn(
              `[pagination] ${gaps.length + nestedGaps.length + overlayCutAnchors.length} page gaps built for ${boundaries} boundaries`,
            )
          }
          // DOM-only (contenteditable=false subtrees the save path never reads);
          // a pristine open keeps the file's cached results like the refresh above
          const tocLines = pm.querySelectorAll<HTMLElement>('.doc-toc-line[data-toc-anchor]')
          if (tocLines.length > 0 && slices.length > 0) {
            const anchorIdx = new Map<string, number>()
            for (const b of parsed.blocks) {
              if (b.docxIndex == null) continue
              for (const a of b.hiddenBookmarks ?? []) anchorIdx.set(a, b.docxIndex)
            }
            const topByIdx = new Map<number, number>()
            for (const b of blocks) {
              if (b.docxIndex != null) topByIdx.set(b.docxIndex, b.top)
            }
            backfillTocPageDom(
              tocLines,
              (anchor) => {
                const idx = anchorIdx.get(anchor)
                const top = idx === undefined ? undefined : topByIdx.get(idx)
                if (top === undefined) return undefined
                // pageAt is 1-based; pageNoTextOf indexes nums/slices 0-based
                return pageNoTextOf(Math.max(0, Math.min(pageOfY(slices, top), slices.length - 1)))
              },
              dirtyRef.current,
            )
          }
        }
        tGapsBuild = performance.now() - tGaps0
        const tSet0 = performance.now()
        const syncMs: Array<[string, number]> = []
        let tStage = tSet0
        const stage = (name: string) => {
          const now = performance.now()
          syncMs.push([name, Math.round(now - tStage)])
          tStage = now
        }
        const layoutBatch = new LayoutBatch(editor.view)
        // split declared-height rows: resolve the engine's target heights to tr elements
        const rowFillEls: Array<{ el: Element; targetPx: number; extraPx?: number }> = []
        for (const f of rowFills) {
          const b = blocks[blockIndex.firstAtTop(f.blockTop, (bb) => !!bb.tableRows)]
          if (!b?.el) continue
          const trs = Array.from(b.el.querySelectorAll('tr')).filter(
            (tr) =>
              !tr.closest('.doc-nested-table') &&
              !tr.classList.contains('page-gap') &&
              !tr.classList.contains('page-repeat-header'),
          )
          const tr = trs[f.row]
          if (tr)
            rowFillEls.push({
              el: tr,
              targetPx: f.targetPx,
              ...(f.extraPx ? { extraPx: f.extraPx } : {}),
            })
        }
        setRowFills(editor.view, rowFillEls, layoutBatch, factor)
        stage('rowFills')
        // oversized single-line blocks: resolve the engine's page-bottom clips to their elements
        const oversizeEls: Array<{ el: HTMLElement; clipPx: number }> = []
        for (const c of oversizeClips) {
          const b =
            blocks[blockIndex.firstAtTop(c.blockTop, (bb) => bb.oversizeLineH !== undefined)]
          if (b?.el) oversizeEls.push({ el: b.el, clipPx: c.clipPx })
        }
        setOversizeClips(editor.view, oversizeEls)
        stage('oversize')
        // page/margin-anchored floated tables: resolve the engine's Y shifts to table elements
        const floatVEls: Array<{
          el: Element
          dyPx: number
          flow?: boolean
          carryPx?: number
          pull?: boolean
        }> = []
        // blocks by rounded top: a scan per shift was floats × blocks on picture-heavy
        // documents; the first match in document order wins, as the scan did
        const blockAtTop = (pred: (b: BlockBox) => boolean) => {
          const byTop = new Map<number, Array<{ i: number; bb: BlockBox }>>()
          blocks.forEach((bb, i) => {
            if (!pred(bb)) return
            const k = Math.round(bb.top)
            const list = byTop.get(k)
            if (list) list.push({ i, bb })
            else byTop.set(k, [{ i, bb }])
          })
          return (top: number): BlockBox | undefined => {
            const k = Math.round(top)
            let hit: { i: number; bb: BlockBox } | undefined
            for (const kk of [k - 1, k, k + 1]) {
              const c = byTop.get(kk)?.find((e) => Math.abs(e.bb.top - top) < 0.5)
              if (c && (!hit || c.i < hit.i)) hit = c
            }
            return hit?.bb
          }
        }
        const pageRelBlockAt = blockAtTop((bb) => bb.pageRelVyPx !== undefined)
        for (const f of floatVShifts) {
          const b = pageRelBlockAt(f.blockTop)
          if (b?.el) floatVEls.push({ el: b.el, dyPx: f.dyPx })
        }
        const pictureAt = blockAtTop((bb) => !!bb.floatPicture)
        for (const f of floatPullUps) {
          const b = pictureAt(f.blockTop)
          if (b?.el) floatVEls.push({ el: b.el, dyPx: f.dyPx, pull: true })
        }
        for (const f of floatCarries) {
          const b = blocks.find(
            (bb) =>
              !!bb.el &&
              !bb.floated &&
              Math.abs(bb.top - (bb.carryAppliedPx ?? 0) - f.blockTop) < 0.5,
          )
          if (b?.el && f.carryPx > 0.5) floatVEls.push({ el: b.el, dyPx: 0, carryPx: f.carryPx })
        }
        const floatTableAt = blockAtTop((bb) => Boolean(bb.floatTable))
        for (const f of floatFlows) {
          const b = floatTableAt(f.blockTop)
          if (b?.el) floatVEls.push({ el: b.el, dyPx: 0, flow: true })
        }
        // split floating table: its anchor paragraph (the next block) starts
        // beside the last portion — a flow spacer carries it past the earlier
        // portions, a gap-sized spacer past the table's in-table gaps
        for (const f of floatSplits) {
          const bi = blockIndex.firstAtTop(f.blockTop, (bb) => !!bb.el && !!bb.floatTable)
          const anchor = bi >= 0 ? blocks[bi + 1] : undefined
          if (!anchor?.el || anchor.floated) continue
          const carry = f.blockTop + f.carryPx - (anchor.top - (anchor.carryAppliedPx ?? 0))
          if (carry > 0.5) floatVEls.push({ el: anchor.el, dyPx: 0, carryPx: carry })
          const gapPx = carryGaps.get(f.blockTop) ?? 0
          if (gapPx > 0.5) {
            gaps.push({
              el: anchor.el,
              carryPx: gapPx,
              metrics: { marginTop: 0, marginBottom: 0, marginLeft: 0, marginRight: 0 },
            })
            carries.push({ y: anchor.top, px: Math.round(gapPx) })
          }
        }
        // the cuts learn the gap widgets' geometry: every page reader below
        // (float shifts, cell boxes, frames, status bar, thumbnails) reads them
        applyGapLedger(cuts, gapLedgerOf([...gaps, ...nestedGaps.map((g) => g.gap)], overlayPages))
        if (pageLayoutRef.current) pageLayoutRef.current.carries = carries
        // no gap widgets (web / read mode, one page): the flow is the page
        const ledgered = cuts.some((c) => c.gap)
        setFloatVShifts(editor.view, floatVEls, layoutBatch)
        stage('floatV')
        pm.classList.remove('measuring-natural')
        setPageGaps(editor.view, gaps, firstPageFloats, layoutBatch)
        stage('pageGaps')
        // Draft keeps what decides where lines wrap, so its page breaks are the print layout's
        const wrapsAsPrint = (viewMode === 'print' || viewMode === 'draft') && !readMode
        // mixed-column canvas: paint the engine's regions via per-block width/translate decorations
        const colSpecs =
          viewMode === 'print' && !readMode && colMode === 'mixed' && secList
            ? columnLayoutSpecs(blocks, slices, secList)
            : []
        // unequal columns: a block's width follows the column it lands in,
        // which follows its measured height — one more pass per change, and a
        // block ping-ponging between columns is pinned to its narrower width
        const colGranted = widthPassGate(colWidthPass, colSpecs)
        // sectPr w:vAlign pages ride the same visual-translate channel
        const vaSpecs =
          viewMode === 'print' && !readMode && secList && hfHs
            ? vAlignShiftSpecs(blocks, slices, secList, sectionGeoms(secList, hfHs))
            : []
        // sections whose content width differs from the canvas section: per-block wrap widths
        const secWSpecs =
          wrapsAsPrint && secList && hfHs
            ? sectionWidthSpecs(blocks, secList, sectionGeoms(secList, hfHs), paperWPx)
            : []
        // sections disagreeing on the typed docGrid pitch: per-block pitch vars
        const gridSpecs = wrapsAsPrint && secList ? sectionGridPitchSpecs(blocks, secList) : []
        // sections disagreeing on the docGrid charSpace delta: per-block letter-spacing vars
        const charSpecs =
          wrapsAsPrint && secList
            ? [
                ...sectionCharSpaceSpecs(blocks, secList, doc.parsed.balanceDbcsSpacing === true),
                ...sectionCharCellSpecs(blocks, secList, gridBasePtRef.current),
              ]
            : []
        // sections disagreeing on the top margin: per-block --doc-margin-top for page-relative anchors
        const topSpecs =
          viewMode === 'print' && !readMode && secList ? sectionTopMarginSpecs(blocks, secList) : []
        // page-opening sections clear the previous section's page-bottom floats
        const sectStartSpecs =
          viewMode === 'print' && !readMode && secList ? sectionPageStartSpecs(blocks, secList) : []
        // w:mirrorMargins: even pages translate onto the swapped text column
        const mirrorSpecs =
          viewMode === 'print' && !readMode && mirrorMargins
            ? mirrorMarginSpecs(
                blocks,
                slices,
                secList ?? [],
                section,
                secList ? pageNumbers(slices, secList) : slices.map((_, n) => n + 1),
              )
            : []
        const mirrorIndSpecs =
          viewMode === 'print' && !readMode && pm.querySelector('.doc-mirror-ind')
            ? mirrorIndentSpecs(
                blocks,
                slices,
                secList ? pageNumbers(slices, secList) : slices.map((_, n) => n + 1),
              )
            : []
        // sectPr w:textDirection sections: writing-mode boxes placed sideways (width wins)
        const vertSpecs =
          viewMode === 'print' && !readMode && secList && hfHs
            ? verticalTextSpecs(blocks, slices, secList, sectionGeoms(secList, hfHs))
            : []
        // one spec per block: mixed-column placement wins the width, translates add up
        const specLists = [
          gridSpecs,
          charSpecs,
          topSpecs,
          sectStartSpecs,
          secWSpecs,
          mirrorSpecs,
          mirrorIndSpecs,
          colSpecs,
          vaSpecs,
          vertSpecs,
        ].filter((l) => l.length > 0)
        let layoutSpecs = specLists.flat()
        if (specLists.length > 1) {
          const mergedSpecs = new Map<HTMLElement, ColumnBlockPlacement>()
          for (const s of layoutSpecs) {
            const prev = mergedSpecs.get(s.el)
            mergedSpecs.set(
              s.el,
              prev ? { ...prev, ...s, dx: prev.dx + s.dx, dy: prev.dy + s.dy } : s,
            )
          }
          layoutSpecs = [...mergedSpecs.values()]
        }
        setColumnLayout(editor.view, layoutSpecs, layoutBatch)
        stage('columnLayout')
        const tCommit0 = performance.now()
        layoutBatch.commit()
        tCommit = performance.now() - tCommit0
        stage('commit')
        // page turns between nested-table rows: gap rows inside the read-only views
        syncNestedTableGaps(editor.view, nestedGaps)
        // in-table gap bands live in the spanning cell's coordinate space:
        // re-anchor them to the paper before the strips are measured/aligned
        alignTableGapFills(pm, factor)
        stage('tableGapFills')
        if ((secWSpecs.length > 0 || mirrorMargins) && section) {
          const cSet = secList?.[0]?.settings ?? section
          alignGapHfStrips(pm, pageLeftOf(cSet) + twipsToPx(cSet.marginLeft), factor)
        }
        stage('hfStrips')
        // after setPageGaps: widget insertion is synchronous, so anchor rects are final
        if (ledgered)
          stampGapFlowTops(pm, cuts, pm.getBoundingClientRect().top + mTopPx * factor, factor)
        syncFloatShifts(
          pm,
          floats,
          pm.getBoundingClientRect().top + mTopPx * factor,
          factor,
          mTopPx - twipsToPx((sections[0]?.settings ?? section)?.marginTop ?? 0),
          ledgered ? { cuts, carries } : undefined,
        )
        stage('floatShifts')
        // Word keeps anchored objects on the page: cell boxes lifted past the
        // paper top by a negative anchor offset are pushed back down
        clampCellBoxTops(pm, pm.getBoundingClientRect().top, factor)
        clampCellImageTops(pm, factor)
        pinCellBoxesToPage(pm, factor, cuts)
        pinPageRelXPictures(pm, factor, cuts)
        pinParagraphRelVPictures(
          pm,
          planOf,
          new Set(gaps.flatMap((g) => ('el' in g && g.suppressLeadMt ? [g.el] : []))),
        )
        stage('clampCells')
        syncCutOverlays((pm.closest('.page-wrap') as HTMLElement) ?? pm, overlayCutAnchors, factor)
        stage('cutOverlays')
        const draftLayer = draftBreaksRef.current
        if (draftLayer) {
          // columns and vertical text flow as one plain column here: its page starts are not the print layout's
          const breaks =
            viewMode === 'draft' && colMode === 'none' && !hasVertical
              ? draftAutoBreaks(cuts, blocks, {
                  endsInPageBreak: (i) => blocks[i]?.el?.querySelector('br.doc-page-br') != null,
                  startsWithMark: (i) =>
                    blocks[i]?.breakBeforeBr === true ||
                    blocks[i]?.el?.classList.contains('page-break-before') === true,
                })
              : []
          // the layer sits outside the zoomed canvas: its px are screen px
          const layerTop = draftLayer.getBoundingClientRect().top
          const flowTop = pm.getBoundingClientRect().top + mTopPx * factor
          syncDraftBreaks(
            draftLayer,
            breaks.map((b) => {
              const el = b.block === undefined ? undefined : blocks[b.block]?.el
              const above = b.block ? blocks[b.block - 1]?.el : undefined
              const top = el ? el.getBoundingClientRect().top : flowTop + b.y * factor
              // between two paragraphs the rule sits in the space between them
              const from = el && above ? above.getBoundingClientRect().bottom : top
              return (Math.min(from, top) + top) / 2 - layerTop
            }),
          )
        }
        {
          // page border (w:pgBorders): per-page overlay boxes (w:display can
          // exclude pages; the border must not run through the page gaps)
          const pbSec = secList?.[0]?.settings ?? section
          const borderStyle = pbSec ? pageBorderStyleOf(pbSec) : null
          const wrapEl = (pm.closest('.page-wrap') as HTMLElement) ?? pm
          const firstFrame = pbSec
            ? { left: pageLeftOf(pbSec), width: twipsToPx(pbSec.pageWidth) }
            : undefined
          syncPageBorders(wrapEl, borderStyle, factor, firstFrame, cuts)
          stage('pageBorders')
          syncLineNumbers(wrapEl, pm, blocks, secList ?? [], factor, blockMetaOf, firstFrame, cuts)
          stage('lineNumbers')
          syncColumnRules(
            wrapEl,
            pm,
            blocks,
            viewMode === 'print' && !readMode ? (secList ?? []) : [],
            factor,
            firstFrame,
            cuts,
          )
          stage('columnRules')
          // differing-width documents paint one sheet per page (the shared paper is transparent)
          syncPageSheets(
            wrapEl,
            factor,
            mixedPaper && viewMode === 'print' && !readMode ? (firstFrame ?? null) : null,
            cuts,
          )
        }
        syncMarginAnnotations(
          (pm.closest('.page-wrap') as HTMLElement) ?? pm,
          pm,
          comments,
          factor,
          doc.parsed.blocks,
          editor.view,
          marginHandlersRef.current,
          cuts,
        )
        stage('sheets+annotations')
        tSetGaps = performance.now() - tSet0
        if (obstaclesChanged) followUp('obstacles')
        // a fill just applied to a split row changed the table height under the slices
        const fSig = rowFillsSig(rowFills)
        if (fSig !== rowFillSig) {
          rowFillSig = fSig
          followUp('rowFills')
        }
        // freshly applied wrap widths (section widths, unequal column widths,
        // vertical-text line lengths) change line breaks: one follow-up remeasure with them in the DOM
        const wSig = wrapWidthSig([...secWSpecs, ...colSpecs, ...vertSpecs, ...sectStartSpecs])
        let wrapWidthsChanged = false
        if (wSig !== secWidthSig) {
          secWidthSig = wSig
          wrapWidthsChanged = true
          followUp('width')
        }
        if (colGranted) {
          wrapWidthsChanged = true
          followUp('columns')
        }
        // freshly applied per-block letter-spacing changes line breaks the same way
        const cSig =
          charSpecs.length === 0
            ? ''
            : `${charSpecs.length}:${[...new Set(charSpecs.map((s) => s.charSpacePt ?? `c${s.charCellPt}`))].join(',')}`
        if (cSig !== charSpaceSig) {
          charSpaceSig = cSig
          wrapWidthsChanged = true
          followUp('charSpace')
        }
        // the shrink loops measured line breaks the new widths just moved;
        // their re-decision rides the follow-up pass already scheduled
        if (wrapWidthsChanged) document.dispatchEvent(new Event(WRAP_WIDTHS_COMMITTED_EVENT))
        // the last page paints as a full sheet like the ones above it:
        // extend the canvas to that page's paper bottom, measured from the last gap
        const bands = pagePaperEdges(cuts, mTopPx).filter((e) => e.gap.kind !== 'cut')
        const lastBand = bands[bands.length - 1]
        if (lastBand && slices.length > 1) {
          const last = slices[slices.length - 1]
          const lastSec = secList?.[last.section]?.settings ?? section
          const lastHf = hfHs?.[last.section] ?? singleHfPx
          const paperTop = lastBand.gapBottom - effectiveTopPx(lastSec, lastHf.headerPx)
          if (pageReaderShadow.enabled) {
            const gapEls = pm.querySelectorAll(PAGE_TURN_SEL)
            const lastGapEl = gapEls[gapEls.length - 1]
            const domTop = lastGapEl
              ? (lastGapEl.getBoundingClientRect().bottom - pm.getBoundingClientRect().top) /
                  factor -
                effectiveTopPx(lastSec, lastHf.headerPx)
              : NaN
            if (!(Math.abs(domTop - paperTop) <= 1))
              pageReaderShadow.report('lastPaper', domTop.toFixed(1), paperTop.toFixed(1))
          }
          const paperH = Math.round(paperTop + twipsToPx(lastSec.pageHeight))
          pm.style.minHeight = `${paperH}px`
          // a mixed-column last page keeps its columns stacked in the flow with no
          // following gap to pull up over the vacated space: pin the box to the
          // paper bottom (the stacked tail below holds nothing visible)
          const stackedTail = last.regions && colMode === 'mixed' && (last.physHeight ?? 0) > 0
          if (stackedTail) pm.style.height = `${paperH}px`
          else pm.style.removeProperty('height')
        } else {
          pm.style.removeProperty('min-height')
          pm.style.removeProperty('height')
        }
        // endnote area: Word puts it right after the last body line, not at the page
        // bottom — anchor it to the flow end measured in the final display state
        setEndnotesAreaTop(
          endnoteItems.length > 0
            ? endnotesAnchorY(
                pm,
                (pm.closest('.page-wrap') ?? pm).getBoundingClientRect().top,
                factor,
              )
            : null,
        )
        // for real-device verification/troubleshooting: current slices and block geometry (read-only snapshot, no functional dependency)
        ;(window as unknown as Record<string, unknown>).__pageDebug = {
          slices,
          fastPasses,
          fastReject,
          gapBoundaryMismatch,
          pageCutMismatch,
          pageCutDiag,
          get pageReaderMismatch() {
            return pageReaderShadow.count
          },
          get pageReaderDiag() {
            return pageReaderShadow.diag
          },
          pageReaderShadowed: pageReaderShadow.enabled,
          pageCuts,
          floatSplits,
          pageCutSig: pageCutSignature(pageCuts),
          passGen,
          floatBoxes: floatWatchRec.boxes,
          floatsZeroHeight: floatWatchRec.zeroHeight,
          planDiff,
          planDiffHistogram: planDiffHistogram(planDiff),
          obstacleShadow: { wrappers: obstacleShadow.wrappers, boxes: obstacleShadow.boxes },
          layoutHealth: layoutHealth.snapshot(),
          carriedSamples,
          resumePasses,
          resumeReject,
          resumePage,
          resumeDiag,
          colMode,
          retrigger,
          retriggerLog: chain.log,
          followUps: chain.passes,
          // live: a refusal after the last pass must be visible too
          get loopStops() {
            return chain.stops
          },
          get tabFreezes() {
            return tabLayoutDiag.freezes
          },
          sliceIterationsCapped: lastOut?.iterationsCapped ?? false,
          sliceRunaway: lastOut?.runaway ?? false,
          rowSplitOverflows: lastOut?.rowSplitOverflows ?? 0,
          colSpecs: colSpecs.map((s) => ({
            w: s.widthPx === undefined ? null : Math.round(s.widthPx),
            dx: Math.round(s.dx),
            dy: Math.round(s.dy),
            cls: s.el.className.slice(0, 30),
          })),
          secs: secList?.map((s, i) => ({
            startType: s.startType,
            cols: sectionColumns(s),
            first: s.firstBlockIndex,
            last: s.lastBlockIndex,
            contentH: hfHs
              ? Math.round(sectionGeoms(secList, hfHs)[i]?.contentHeight ?? -1)
              : undefined,
          })),
          blocks: blocks.map((b) => ({
            top: b.top,
            height: b.height,
            docxIndex: b.docxIndex,
            section: b.section,
            empty: b.emptyPara,
            nLines: b.lineBoxes?.length,
            oversize: b.oversizeLineH,
          })),
          tableRows: blocks
            .filter((b) => b.tableRows)
            .map((b) => ({
              top: b.top,
              rows: b.tableRows!.map((r) => ({
                h: Math.round(r.height),
                cb: r.contentBottom === undefined ? null : Math.round(r.contentBottom),
                cuts: r.cutYs?.map((c) => Math.round(c)) ?? null,
              })),
            })),
          remeasureMs: performance.now() - tStart,
          measureMs: tMeasure,
          sliceMs: tSlice,
          sliceIters,
          sliceSplit,
          commitMs: tCommit,
          syncMs,
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
      setLayoutRev((v) => v + 1)
    }
    const onUpdate = (
      fastIndex: number | null = null,
      dirtyIndex: number | null = fastIndex,
      isFollowUp = false,
    ) => {
      if (disposed) return
      if (fastIndex === null) fastEdits = null
      else fastEdits?.add(fastIndex)
      dirtyFrom = dirtyIndex === null || dirtyFrom === null ? null : Math.min(dirtyFrom, dirtyIndex)
      // while the tail streams, chunks land every few frames: a plain debounce
      // would either never fire (dense chunks) or pay a whole-document pass
      // per chunk (sparse ones); keep the pass already due and pace new ones
      if (isPhasedContentPending()) {
        if (timer) return
        // a whole-document pass grows with the streamed content; keep it to a
        // fraction of the streaming time so the tail lands sooner
        const gap = Math.min(
          STREAMING_PASS_GAP_MAX_MS,
          Math.max(STREAMING_PASS_GAP_MS, lastPassMs * STREAMING_PASS_DUTY),
        )
        const delay = Math.max(300, gap - (performance.now() - lastPassAt))
        timer = window.setTimeout(remeasure, delay)
        return
      }
      if (timer) window.clearTimeout(timer)
      // a whole-document pass after every pause blocks typing on long
      // documents: wait longer when the last pass was slow (Word paginates in
      // the background too); short documents keep the 300 ms feel. A typed
      // edit arriving after a follow-up re-arms the timer with the edit delay.
      timer = window.setTimeout(
        remeasure,
        isFollowUp
          ? FOLLOW_UP_PASS_DELAY_MS
          : Math.min(EDIT_PASS_DEBOUNCE_MAX_MS, Math.max(300, lastPassMs * EDIT_PASS_DUTY)),
      )
    }
    // while a phased open streams its tail, this effect re-runs as the section
    // state fills in: a synchronous whole-document pass per run made a long
    // open quadratic, so the paced scheduler takes it (the last chunk's update
    // still lands a full pass once the tail has settled)
    flushPaginationRef.current = (fresh = false) => {
      if (fresh) onUpdate()
      if (timer === null) return
      window.clearTimeout(timer)
      remeasure()
    }
    if (isPhasedContentPending()) onUpdate()
    else remeasure()
    // async @font-face loading triggers a full reflow (line-break points change); pagination
    // must be remeasured, and cached line samples invalidated (block heights may not change)
    // header/footer strips probed under a fallback face re-measure too
    // (debounced: loadingdone fires per face, thousands of times on CJK documents)
    let hfTimer: number | undefined
    const onFontsChanged = (reprobeHf: boolean) => {
      bumpLineSampleFontEpoch()
      // strut metrics measured under a fallback face before a bundled font loaded
      if (takeSettledMetricFaces()) {
        const parsed = fileCtxRef.current.doc?.parsed
        const live = liveStylesRef.current
        if (parsed) {
          setDocCss(
            liveDocCss(live?.parsed === parsed ? { ...parsed, styles: live.styles } : parsed),
          )
        }
      }
      if (reprobeHf) {
        bumpHfProbeFontEpoch()
        if (hfTimer) window.clearTimeout(hfTimer)
        hfTimer = window.setTimeout(() => setHfMeasureEpoch((e) => e + 1), 300)
      }
      onUpdate()
    }
    // the epoch bump re-runs this effect; an already-settled ready promise must not re-probe
    const fontsLoading = document.fonts.status === 'loading'
    document.fonts.ready.then(() => onFontsChanged(fontsLoading)).catch(() => {})
    const onLoadingDone = () => onFontsChanged(true)
    document.fonts.addEventListener('loadingdone', onLoadingDone)
    // a CJK font's localized name resolves through the installed face's other
    // spelling: the alias map lands over IPC, after the first layout, so the
    // document is re-styled and re-paginated like a settled metric face. An
    // update that touches no declared font must not reflow anything.
    const onAliases = (names: readonly string[]) => {
      seenAliasSeqRef.current = fontAliasAnswers().seq
      const parsed = fileCtxRef.current.doc?.parsed
      if (!parsed) return
      if (!docFontsCoveredBy(parsed, names).length) return
      bumpLineSampleFontEpoch()
      const live = liveStylesRef.current
      setDocCss(liveDocCss(live?.parsed === parsed ? { ...parsed, styles: live.styles } : parsed))
      onUpdate()
    }
    const offFontAliases = onFontAliasesChanged(onAliases)
    // an answer that landed between the open's setDocCss and this subscription
    const landed = fontAliasAnswers()
    if (landed.seq !== seenAliasSeqRef.current) onAliases(landed.names)
    scroller.addEventListener('scroll', locate, { passive: true })
    const onDocUpdate = ({ transaction }: { transaction: Transaction }) => {
      resetWidthPassHistory(colWidthPass)
      onUpdate(singleInlineEditIndex(transaction), firstChangedTopLevelIndex(transaction))
    }
    editor?.on('update', onDocUpdate)
    // justify-shrink re-decides via decoration-only transactions that move the
    // wrap points of justified paragraphs; tiptap 'update' fires only on doc
    // changes, so without this the inline gap widgets keep their pre-shrink cut
    // positions and force a stretched-sparse page-last line (r177). The stale
    // block re-samples via the shrink fingerprint in lineSampleSig.
    // the shrink list covers every justified paragraph; only the entries that
    // differ from the last list moved a wrap point, so the pass resumes there
    let shrinkSig = new Map<number, string>()
    const onShrinkTr = (props: { transaction: Transaction }) => {
      const decos = props.transaction.getMeta(justifyShrinkPluginKey) as
        Array<{ from: number; to: number; type?: { attrs?: { style?: string } } }> | undefined
      if (decos && editor) {
        const next = new Map<number, string>()
        // guard: DecorationSet.create nulls out consumed entries of the array
        // it is handed — a plugin seeing this meta first must not crash us
        for (const d of decos) if (d) next.set(d.from, `${d.to}:${d.type?.attrs?.style ?? ''}`)
        let minPos = Infinity
        for (const [from, sig] of next)
          if (shrinkSig.get(from) !== sig) minPos = Math.min(minPos, from)
        for (const from of shrinkSig.keys()) if (!next.has(from)) minPos = Math.min(minPos, from)
        shrinkSig = next
        if (minPos === Infinity) return
        const { doc: pmDoc } = editor.state
        followUpAt('justify-shrink', pmDoc.resolve(Math.min(minPos, pmDoc.content.size)).index(0))
        return
      }
      if (props.transaction.getMeta(floatFlowChangedMeta)) followUp('float-flow')
      if (hyphenationRelayout(props.transaction)) {
        bumpLineSampleFontEpoch()
        followUp('hyphenation')
      }
      const tabFrom = props.transaction.getMeta(tabReflowMeta) as number | undefined
      if (tabFrom !== undefined && editor) {
        const { doc: pmDoc } = editor.state
        followUpAt('tab-wrap', pmDoc.resolve(Math.min(tabFrom, pmDoc.content.size)).index(0))
      }
    }
    editor?.on('transaction', onShrinkTr)
    // an equation wrapped to its column after the pass measured it is taller now
    const onMathWrap = (e: Event) => {
      const bi = lastBlocks.findIndex((b) => b.el === (e as CustomEvent<HTMLElement>).detail)
      followUpAt('math-wrap', bi >= 0 ? bi : null)
    }
    document.addEventListener(MATH_WRAP_CHANGED_EVENT, onMathWrap)
    return () => {
      disposed = true
      document.removeEventListener(MATH_WRAP_CHANGED_EVENT, onMathWrap)
      floatWatch.dispose()
      flushPaginationRef.current = () => {}
      if (timer) window.clearTimeout(timer)
      if (hfTimer) window.clearTimeout(hfTimer)
      document.fonts.removeEventListener('loadingdone', onLoadingDone)
      offFontAliases()
      scroller.removeEventListener('scroll', locate)
      editor?.off('update', onDocUpdate)
      editor?.off('transaction', onShrinkTr)
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
    blockPlans,
    pageFootnotesOf,
    endnoteItems,
    sectEndnoteGroups,
    noteSettings,
    notesBeneathText,
    editNote,
    canvasTop,
    effTopSingle,
    effBottomSingle,
    singleVariantGeom,
    singleHfPx,
    hfHeightsOf,
    measureSingleFlow,
    colGeomsFor,
    colMode,
    hasVertical,
    header,
    footer,
    hfVariants,
    titlePg,
    evenOddHf,
    mirrorMargins,
    autoHyphenation,
    hfState,
    hfSections,
    hfResolveAt,
    legacyBands,
    comments,
    commentsUi.cards,
    commentsUi.panel,
    // display-mode toggles reflow the text (No Markup hides deletions) and gate
    // the change bars, but dispatch no doc change: remeasure must follow them
    // (revisions-in-balloons lifts deletions out of the flow the same way)
    revisionDisplay,
    balloonMode,
    delSectBreaks,
    pendingWatermarkImage,
  ])

  // section at the cursor: the target the Layout tab acts on
  useEffect(() => {
    if (!editor || sections.length === 0) {
      setActiveSection(0)
      return
    }
    const locateSection = () => {
      const { $head } = editor.state.selection
      const topIndex = $head.depth > 0 ? $head.index(0) : 0
      setActiveSection(sectionIndexAtBlock(editor.state.doc, sections, topIndex))
      if (!mirrorMargins) return
      const layout = pageLayoutRef.current
      const el = $head.depth > 0 ? editor.view.nodeDOM($head.before(1)) : null
      // a fresh layout decoration redraws the node, so the measured element may be
      // the block's previous DOM: fall back to its data-idx identity
      const idx = el instanceof HTMLElement ? el.dataset.idx : undefined
      const block =
        layout && el
          ? layout.blocks.find(
              (b) => b.el === el || (idx !== undefined && String(b.docxIndex) === idx),
            )
          : undefined
      setCaretPageNo(
        block && layout
          ? (pageNumbers(layout.slices, layout.sections as SectionInfo[])[
              pageOfY(layout.slices, block.top)
            ] ?? 1)
          : 1,
      )
    }
    locateCaretPageRef.current = locateSection
    locateSection()
    editor.on('selectionUpdate', locateSection)
    return () => {
      editor.off('selectionUpdate', locateSection)
    }
  }, [editor, sections, mirrorMargins])

  // status-bar word count: the whole text is re-walked per doc, so pace it
  // (Word refreshes its count on idle too). A trailing throttle, not a
  // debounce: the streamed tail of a phased open would keep resetting one
  const [wordCount, setWordCount] = useState(0)
  // Word's "selected/total" pair; 0 means no selection, so the bar falls back
  // to the plain total. Cheap enough (a text range, not a document walk) to
  // follow the selection directly rather than ride the throttle above.
  const [selectionWordCount, setSelectionWordCount] = useState(0)
  const [selectionHead, setSelectionHead] = useState(0)
  // changes outside the body: notes and header / footer parts follow their models,
  // textboxes and note anchors the document (revisionRanges refreshes with it)
  const storyRevisions = useMemo<StoryRevision[]>(() => {
    if (!editor || !doc) return []
    const pmDoc = editor.state.doc
    return [
      ...collectHfRevisions(hfState),
      ...collectTextboxRevisions(pmDoc),
      ...collectNoteRevisions(pmDoc, 'footnote', footnotes),
      ...collectNoteRevisions(pmDoc, 'endnote', endnotes),
    ]
    // revisionRanges: the throttled document-change trigger
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editor, doc, hfState, footnotes, endnotes, revisionRanges])
  storyRevisionsRef.current = storyRevisions
  const allRevisions = useMemo(
    () => (storyRevisions.length ? [...revisionRanges, ...storyRevisions] : revisionRanges),
    [revisionRanges, storyRevisions],
  )
  const revisionCount = allRevisions.length
  const shownRevisions = useMemo(
    () => filterRevisions(allRevisions, revisionFilter),
    [allRevisions, revisionFilter],
  )
  // the review pane's author list in ink order: the document's own user keeps
  // the plain default ink, AI edits and human collaborators get colored inks
  // (revisionAuthorsInkOrder / #2654); the state mirror tracks Settings renames
  const [ownAuthor, setOwnAuthor] = useState(authorName())
  useEffect(() => onAuthorNameChanged(() => setOwnAuthor(authorName())), [])
  const reviewAuthors = useMemo(
    () => revisionAuthorsInkOrder(allRevisions, ownAuthor),
    [allRevisions, ownAuthor],
  )
  const reviewEntries = useMemo(
    () => (editor ? buildReviewEntries(editor.state.doc, shownRevisions, reviewAuthors) : []),
    [editor, shownRevisions, reviewAuthors],
  )
  const revisionAuthorStyle = useMemo(
    () => revisionAuthorCss(reviewAuthors, revisionFilter.hiddenAuthors, movesByAuthor(tcOptions)),
    [reviewAuthors, revisionFilter.hiddenAuthors, tcOptions],
  )
  // the story view's four-ink split (own / others / AI; #2654), scoped under
  // .view-story so the page view's per-author rainbow is untouched
  const storyAuthorStyle = useMemo(
    () =>
      viewMode === 'story'
        ? storyAuthorshipCss(
            allRevisions,
            ownAuthor,
            revisionFilter.hiddenAuthors,
            movesByAuthor(tcOptions),
          )
        : '',
    [viewMode, allRevisions, ownAuthor, revisionFilter.hiddenAuthors, tcOptions],
  )
  const isActiveEntry = (e: ReviewEntry) =>
    e.rev.story
      ? e.rev.story.id === activeStoryId
      : activeStoryId === null && selectionHead >= e.rev.from && selectionHead <= e.rev.to
  const activeReviewIndex = useMemo(
    () => reviewEntries.findIndex(isActiveEntry),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [reviewEntries, selectionHead, activeStoryId],
  )
  // the AI panel list counts every pending change, untouched by the Show Markup filters
  const allReviewEntries = useMemo(
    () => (editor ? buildReviewEntries(editor.state.doc, allRevisions, reviewAuthors) : []),
    [editor, allRevisions, reviewAuthors],
  )
  const activeAllReviewIndex = useMemo(
    () => allReviewEntries.findIndex(isActiveEntry),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [allReviewEntries, selectionHead, activeStoryId],
  )
  useEffect(() => {
    if (!editor) return
    const refreshSelection = () => {
      setSelectionWordCount(selectedWordCount(editor.state.doc, editor.state.selection))
      setSelectionHead(editor.state.selection.head)
      // a caret move in the body ends the stop on a change outside it
      if (lastStoryRef.current !== null) setActiveStoryId(null)
    }
    refreshSelection()
    editor.on('selectionUpdate', refreshSelection)
    return () => {
      editor.off('selectionUpdate', refreshSelection)
    }
  }, [editor])
  useEffect(() => {
    if (!editor) return
    let timer = 0
    let slowWalk = false
    const refresh = () => {
      timer = 0
      const started = performance.now()
      setWordCount(wordCountOfDoc(editor.state.doc))
      setRevisionRanges(doc ? revisionsOfDoc(editor.state.doc) : [])
      setSelectionWordCount(selectedWordCount(editor.state.doc, editor.state.selection))
      slowWalk = performance.now() - started > WORD_COUNT_SLOW_WALK_MS
    }
    const onUpdate = () => {
      if (!timer)
        timer = window.setTimeout(
          refresh,
          slowWalk ? WORD_COUNT_SLOW_INTERVAL_MS : WORD_COUNT_THROTTLE_MS,
        )
    }
    refresh()
    editor.on('update', onUpdate)
    return () => {
      window.clearTimeout(timer)
      editor.off('update', onUpdate)
    }
  }, [editor, doc, docLoading])

  /** the counts of Word's Word Count dialog: pages/lines estimated from the current layout */
  const documentStats = useCallback((): DocStats | null => {
    if (!editor) return null
    const text = documentTextForWordCount(editor.state.doc)
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
      if (node.isTextblock && documentTextForWordCount(node).trim() !== '') paragraphs++
      return true
    })
    return { ...textStats(text), pages: pageInfo.total, paragraphs, lines }
  }, [editor, zoom, pageInfo.total])
  const openStats = useCallback(() => {
    const body = documentStats()
    if (!body) return
    const notes = textStats([...footnotes, ...endnotes].map((note) => note.text).join('\n'))
    // a note's lines are not laid out where they can be measured: one per paragraph
    setStats({ body, notes: { ...notes, pages: 0, lines: notes.paragraphs } })
  }, [documentStats, footnotes, endnotes])

  // Review > Editor, the Tools menu and Word's F7 all run the same AI proofread
  // behind the one-time whole-document-rewrite acknowledgement
  const runAiProofread = useCallback(() => {
    if (
      localStorage.getItem(AI_REWRITE_ACK_KEY) !== '1' &&
      !window.confirm(t('ribbonAiRewriteConfirm'))
    )
      return
    localStorage.setItem(AI_REWRITE_ACK_KEY, '1')
    setShowAi(true)
    setAiPreset({ text: t('ribbonEditorPrompt'), nonce: Date.now(), autoRun: true })
  }, [])

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

  // Save runs ahead of the editor keymaps (capture phase): with a non-Latin
  // layout active the keystroke otherwise reaches the editor's own bindings and
  // the report sees Ctrl+S undo the last edit instead of saving the document.
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && isCommandLetter(e, 's')) {
        e.preventDefault()
        e.stopPropagation()
        void save(e.shiftKey)
      }
    }
    window.addEventListener('keydown', handler, true)
    return () => window.removeEventListener('keydown', handler, true)
  }, [save])

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
      if ((e.metaKey || e.ctrlKey) && isCommandLetter(e, 'o')) {
        e.preventDefault()
        void openFile()
      }
      if ((e.metaKey || e.ctrlKey) && isCommandLetter(e, 'f')) {
        e.preventDefault()
        if (doc) {
          // bump even when the panel is already open: focus returns to the find box
          setShowFind(true)
          setFindFocusInput((n) => n + 1)
        }
      }
      // Word's replace: Ctrl+H everywhere (macOS Cmd+H is the system hide role,
      // which never reaches the renderer, so this branch is Ctrl+H there too)
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey && isCommandLetter(e, 'h')) {
        e.preventDefault()
        if (doc) {
          setShowFind(true)
          setFindFocusReplace((n) => n + 1)
        }
      }
      // Go To: ⌥⌘G on the Mac (Word for Mac), Ctrl+G elsewhere (⇧⌘G stays Word Count)
      if (
        (e.metaKey || e.ctrlKey) &&
        !e.shiftKey &&
        e.altKey === IS_MAC &&
        e.code === 'KeyG' &&
        doc
      ) {
        e.preventDefault()
        setShowFind(true)
        setFindFocusGoTo((n) => n + 1)
      }
      // Word dialog shortcuts: Font ⌘D / Paragraph ⌥⌘M / Hyperlink ⌘K
      if (
        (e.metaKey || e.ctrlKey) &&
        !e.shiftKey &&
        !e.altKey &&
        isCommandLetter(e, 'd') &&
        doc &&
        canEdit
      ) {
        e.preventDefault()
        setShowFontDialog(true)
      }
      if ((e.metaKey || e.ctrlKey) && e.altKey && e.code === 'KeyM' && doc && canEdit) {
        e.preventDefault()
        setShowParaDialog(true)
      }
      if (
        (e.metaKey || e.ctrlKey) &&
        !e.shiftKey &&
        !e.altKey &&
        isCommandLetter(e, 'k') &&
        doc &&
        canEdit
      ) {
        e.preventDefault()
        setShowLinkModal(true)
      }
      if (e.key === 'F9' && e.altKey && doc && editor) {
        e.preventDefault()
        toggleAllFieldCodes(editor.view)
      } else if (e.key === 'F9' && doc && editor?.isEditable) {
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
      // Word line-spacing shortcuts ⌘1 / ⌘2 / ⌘5; e.code because shifted-digit
      // layouts (AZERTY) make e.key unreliable
      const SPACING_KEYS: Record<string, number> = { Digit1: 1, Digit2: 2, Digit5: 1.5 }
      if (
        (e.metaKey || e.ctrlKey) &&
        !e.shiftKey &&
        !e.altKey &&
        e.code in SPACING_KEYS &&
        editor &&
        canEdit
      ) {
        e.preventDefault()
        runCommand('line-spacing', String(SPACING_KEYS[e.code]))
      }
      // Paragraph styles ⌥⌘0 Normal / ⌥⌘1..3 headings (Ctrl+Shift+N is taken by New Window)
      const STYLE_KEYS: Record<string, 'p' | 'h1' | 'h2' | 'h3'> = {
        Digit0: 'p',
        Digit1: 'h1',
        Digit2: 'h2',
        Digit3: 'h3',
      }
      if (
        (e.metaKey || e.ctrlKey) &&
        e.altKey &&
        !e.shiftKey &&
        e.code in STYLE_KEYS &&
        editor &&
        canEdit
      ) {
        e.preventDefault()
        // textboxes have no heading nodes — same gate as the ribbon style gallery
        if (!getActiveSubEditor())
          applyParagraphStyle(editor, STYLE_KEYS[e.code], ribbonStylesRef.current)
      }
      // Word grow/shrink font ⇧⌘. / ⇧⌘, — via the ribbon closure so bursts stay coalesced
      if (
        (e.metaKey || e.ctrlKey) &&
        e.shiftKey &&
        !e.altKey &&
        (e.code === 'Period' || e.code === 'Comma') &&
        editor &&
        canEdit
      ) {
        e.preventDefault()
        runCommand(e.code === 'Period' ? 'font-grow' : 'font-shrink')
      }
      // Word's one-point nudge ⌘] / ⌘[
      if (
        (e.metaKey || e.ctrlKey) &&
        !e.shiftKey &&
        !e.altKey &&
        (e.code === 'BracketRight' || e.code === 'BracketLeft') &&
        editor &&
        canEdit
      ) {
        e.preventDefault()
        ribbonActionsRef.current.nudgeFontSize?.(e.code === 'BracketRight' ? 1 : -1)
      }
      // Superscript / subscript. Word's own ⌘= / ⇧⌘= collide with the zoom
      // accelerators for the "=" half only, so the unshifted pair is on ⌘. / ⌘,
      // (the same keys the shifted grow/shrink pair uses).
      const vertAlign = e.shiftKey
        ? e.code === 'Equal'
          ? 'superscript'
          : null
        : e.code === 'Period'
          ? 'superscript'
          : e.code === 'Comma'
            ? 'subscript'
            : null
      if ((e.metaKey || e.ctrlKey) && !e.altKey && vertAlign && editor && canEdit) {
        e.preventDefault()
        runCommand(vertAlign)
      }
      // Word's Shift+F3 case ring
      if (e.shiftKey && e.key === 'F3' && editor && canEdit) {
        e.preventDefault()
        const target = getActiveSubEditor() ?? editor
        applyCase(target, nextCaseMode(selectionText(target)))
      }
      // Clear character formatting ⌃␣ (Word uses Ctrl+Space on both platforms)
      if (e.ctrlKey && !e.metaKey && !e.shiftKey && !e.altKey && e.code === 'Space' && canEdit) {
        e.preventDefault()
        ;(getActiveSubEditor() ?? editor)?.chain().focus().unsetAllMarks().run()
      }
      // Indent ⌃M / outdent ⌃⇧M and hanging indent ⌘T / ⇧⌘T. Ctrl-only on both
      // platforms: ⌘M minimizes the window, and indenting a paragraph on the way
      // out would be a silent edit.
      if (e.ctrlKey && !e.metaKey && !e.altKey && e.code === 'KeyM' && editor && canEdit) {
        e.preventDefault()
        if (!getActiveSubEditor()) stepParagraphIndent(editor, e.shiftKey ? -1 : 1)
      }
      if ((e.metaKey || e.ctrlKey) && !e.altKey && e.code === 'KeyT' && editor && canEdit) {
        e.preventDefault()
        if (!getActiveSubEditor()) stepHangingIndent(editor, e.shiftKey ? -1 : 1)
      }
      // Clear paragraph formatting — Word's Ctrl+Q (⌘Q quits on macOS)
      if (e.ctrlKey && !e.metaKey && !e.shiftKey && !e.altKey && e.code === 'KeyQ' && canEdit) {
        e.preventDefault()
        if (editor && !getActiveSubEditor()) clearParagraphFormatting(editor)
      }
      // Formatting marks: ⌘8 in Word for Mac, Ctrl+Shift+8 on Windows
      if ((e.metaKey || e.ctrlKey) && !e.altKey && e.code === 'Digit8' && doc) {
        e.preventDefault()
        runCommand('show-marks')
      }
      // Track changes ⇧⌘E; forced by an editing restriction it cannot be turned off
      if ((e.metaKey || e.ctrlKey) && e.shiftKey && !e.altKey && e.code === 'KeyE' && doc) {
        e.preventDefault()
        if (!isProtected && !trackChangesForced) toggleTrackChangesByUser()
      }
      // Word count ⇧⌘G
      if ((e.metaKey || e.ctrlKey) && e.shiftKey && !e.altKey && e.code === 'KeyG' && doc) {
        e.preventDefault()
        runCommand('word-count')
      }
      // New comment ⌥⌘A (Word for Mac; Windows Word's Ctrl+Alt+M stays on the
      // Paragraph dialog here)
      if ((e.metaKey || e.ctrlKey) && e.altKey && e.code === 'KeyA' && doc && canEdit) {
        e.preventDefault()
        runCommand('insert-comment')
      }
      // Footnote ⌥⌘F, endnote ⌥⌘E on the Mac / Ctrl+Alt+D on Windows. The
      // endnote key is platform-exclusive: ⌥⌘D belongs to the macOS Dock.
      if ((e.metaKey || e.ctrlKey) && e.altKey && doc && canEdit) {
        const endnoteCode = IS_MAC ? 'KeyE' : 'KeyD'
        const note = e.code === 'KeyF' ? 'footnote' : e.code === endnoteCode ? 'endnote' : null
        if (note) {
          e.preventDefault()
          runCommand(note === 'footnote' ? 'insert-footnote' : 'insert-endnote')
        }
      }
      // Date / time fields ⌥⇧D / ⌥⇧T
      if (e.altKey && e.shiftKey && !e.metaKey && !e.ctrlKey && doc && canEdit) {
        const instr = e.code === 'KeyD' ? 'DATE' : e.code === 'KeyT' ? 'TIME' : null
        if (instr) {
          e.preventDefault()
          runCommand('insert-field', instr)
        }
      }
      // Proofread F7 (Word's spelling & grammar check)
      if (e.key === 'F7' && !e.shiftKey && doc) {
        e.preventDefault()
        runCommand('ai-proofread')
      }
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [
    save,
    openFile,
    editor,
    doc,
    updateFields,
    runCommand,
    trackChangesForced,
    isProtected,
    toggleTrackChangesByUser,
  ])

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

  // one host for every entry point that runs a command; rebuilt each render so a handler reads current state
  // the selection of whichever editor has the caret: a text box, a header or footer, or the body
  const hyphenationSettings = hyphenationSettingsOf(doc?.parsed, hyphenationEdit)
  const applyHyphenation = (next: HyphenationSettings & { auto: boolean }, manual = false) => {
    const parsed = doc?.parsed
    if (!editor || !parsed) return
    const edit: Partial<HyphenationSettings> = {}
    if (next.caps !== hyphenationSettings.caps) edit.caps = next.caps
    if (next.limit !== hyphenationSettings.limit) edit.limit = next.limit
    if (next.zoneTwips !== hyphenationSettings.zoneTwips) edit.zoneTwips = next.zoneTwips
    const rulesChanged = Object.keys(edit).length > 0
    if (!rulesChanged && next.auto === autoHyphenation && manual === manualHyphenationRef.current)
      return
    if (rulesChanged) setHyphenationEdit({ ...hyphenationEdit, ...edit })
    // words break at other places now with no change to the content: the cached line
    // samples are stale, and the pagination effect runs a pass on the new value
    bumpLineSampleFontEpoch()
    if (next.auto !== autoHyphenation) {
      // set before the CSS below is built; the render that follows writes the same value
      autoHyphenationRef.current = next.auto
      setAutoHyphenation(next.auto)
      setAutoHyphenationDirty(true)
    }
    manualHyphenationRef.current = manual
    setManualHyphenation(manual)
    const live = liveStylesRef.current
    setDocCss(liveDocCss(live?.parsed === parsed ? { ...parsed, styles: live.styles } : parsed))
    // Chromium hyphenates only under an explicit lang, as on open (file-actions)
    const docLang = next.auto ? parsed.docDefaults?.lang : undefined
    if (docLang) editor.view.dom.setAttribute('lang', docLang)
    else editor.view.dom.removeAttribute('lang')
    const config = hyphenationConfigOf(parsed, next.auto || manual, next)
    const elsewhere = hyphenationConfigOf(parsed, next.auto, next)
    const hand = () => {
      if (fileCtxRef.current.doc?.parsed === parsed) setHyphenation(editor, config, elsewhere)
    }
    if (config.dict && !hyphenationReady(config.dict)) void loadHyphenation(config.dict).then(hand)
    else hand()
  }
  const mathPr = mathPrEdit ?? doc?.parsed.mathPr ?? DEFAULT_MATH_PR
  // after the commit: the wrap measures against the stylesheet this render wrote
  useEffect(() => setMathWrapRules(mathPr), [mathPr])
  const applyMathPr = (next: MathPr) => {
    if (!editor || !doc) return
    setMathPrEdit(next)
    setMathLimits(next)
    refreshLimitDefaults(editor)
  }
  const selectedProjection = (): string => {
    const target = getActiveSubEditor() ?? editor
    if (!target || target.state.selection.empty) return ''
    const { from, to } = target.state.selection
    return redactTextBetween(target.getJSON() as RedactableNode, from, to, ' ').trim()
  }
  commandHostRef.current = {
    editor,
    hasDoc: Boolean(doc),
    get gridBasePt() {
      return gridBasePtRef.current
    },
    newFile: () => void newFile(),
    openFile: () => void openFile(),
    openRecent: (path) => void openRecent(path),
    save: (saveAs) => void save(saveAs),
    newFromTemplate: () => setShowTemplateGallery(true),
    saveAsTemplate: () => void saveAsTemplateImpl(fileCtxRef.current),
    nameWithAi: () => void nameWithAi(),
    exportPdf: () => void exportPdf(),
    exportHtml: () => void exportHtml(),
    exportMarkdown: () => void exportMarkdown(),
    exportImages: () => void exportImages(),
    printDoc: () => void printDoc(),
    stepZoom: (dir) => setZoom((z) => stepDocsZoom(z, dir)),
    setZoom: (pct) => setZoom(clampDocsZoom(pct)),
    zoomFit,
    toggleAi: () => setShowAi((v) => !v),
    toggleDark: () => updateDarkPage(!darkPageRef.current),
    toggleMarks: () => setShowMarks((v) => !v),
    toggleStylesPane: () => setShowStylesPane((v) => !v),
    toggleSelectionPane: () => setShowSelectionPane((v) => !v),
    toggleCitationsPane: () => setShowCitationsPane((v) => !v),
    toggleDrawing: () => {
      if (inkTool !== 'select') return setInkTool('select')
      // the pen's controls are on the Draw tab, and the ribbon puts the pen away off it
      setRibbonTabRequest({ tab: 'draw', nonce: Date.now() })
      setInkTool(inkActive.kind)
    },
    setDesignAsDefault: () => {
      const { before, after, line, multiple } = defaultStyleSpacing()
      saveDesignDefault({
        ...(themeFonts ? { fonts: themeFonts } : {}),
        ...(themeColors ? { colors: themeColors } : {}),
        ...(after != null && line != null && multiple
          ? { spacing: { ...(before ? { beforeTwips: before } : {}), afterTwips: after, line } }
          : {}),
      })
      setStatus(t('ribbonSetAsDefaultDone'))
    },
    nothingToCrop: () => setStatus(t('ribbonAutoCropNothing')),
    selectChartData: () => {
      if (editor && !selectChartForData(editor)) setStatus(t('ribbonSelectDataNone'))
    },
    stepFontSize: (dir) => ribbonActionsRef.current.stepFontSize?.(dir),
    clearFormatting: () =>
      editor && clearAllFormatting(editor, getActiveSubEditor(), ribbonStylesRef.current),
    createListDef: (levels) => {
      setListRecentsEpoch((n) => n + 1)
      return createCustomListDef(levels)
    },
    listStyleNum: (styleId) => {
      const style = listStyles.find((s) => s.styleId === styleId)
      return style ? listStyleNumImpl(numberingCtxRef.current, style) : null
    },
    openListDialog: (kind, card) => {
      setListCard(card ?? null)
      setListDialog(kind)
    },
    addListStyle: (name, levels) => {
      const styleId = addListStyleImpl(numberingCtxRef.current, name, levels, listStyles)
      if (!styleId) return
      setListDialog(null)
      setAddedListStyle(styleId)
      setShowBulletsNumbering(true)
    },
    deleteListStyle: (styleId) => {
      const style = listStyles.find((s) => s.styleId === styleId)
      if (style) deleteListStyleImpl(numberingCtxRef.current, style)
    },
    applyStyle: (styleId) => {
      const styles = ribbonStylesRef.current
      // a document without a style table still offers the built-in gallery set
      const info = (styles ?? FALLBACK_STYLES).get(styleId)
      if (editor && info)
        applyGalleryStyle(editor, getActiveSubEditor(), info, styles, formatState.charStyleId)
    },
    aiPreset: (text) => {
      setShowAi(true)
      setAiPreset({ text, nonce: Date.now(), autoRun: true })
    },
    applyHeadingLevel: (level) => {
      if (editor) applyParagraphStyle(editor, level ? `h${level}` : 'p', ribbonStylesRef.current)
    },
    checkSpelling: () => setShowSpelling(true),
    thesaurus: () => {
      // a word or a short phrase, as on the right-click menu; the model reads the projection
      const text = selectedProjection()
      if (!text || text.length > 20) return setStatus(t('appThesaurusSelect'))
      setShowAi(true)
      setAiPreset({ text: t('appSynonymsPrompt', { text }), nonce: Date.now(), autoRun: true })
    },
    translate: (scope) => {
      const langName = t(ribbonLangKey(translationLanguage(lang)))
      if (scope === 'selection') {
        // the passage goes in the request, as on the right-click menu, so a text box's is found
        const text = selectedProjection()
        if (!text) return setStatus(t('appTranslateSelect'))
        setShowAi(true)
        setAiPreset({
          text: t('appTranslateSelectionPrompt', { lang: langName, text }),
          nonce: Date.now(),
          autoRun: true,
        })
        return
      }
      // a whole-document rewrite asks once, as the ribbon's Translate does
      if (
        localStorage.getItem(AI_REWRITE_ACK_KEY) !== '1' &&
        !window.confirm(t('ribbonAiRewriteConfirm'))
      )
        return
      localStorage.setItem(AI_REWRITE_ACK_KEY, '1')
      setShowAi(true)
      setAiPreset({
        text: t('ribbonTranslatePrompt', { lang: langName }),
        nonce: Date.now(),
        autoRun: true,
      })
    },
    pickMergeDataSource: () => void pickMergeDataSource(),
    toggleMergePreview: () => setMergePreviewOn((on) => !on),
    hasMergeSource: Boolean(mergeSource),
    openRuleDialog: setMergeRule,
    onDisk: Boolean(doc?.filePath),
    moveFile: () =>
      void window.desktop.moveDoc().then((result) => {
        if (result && !result.ok) setStatus(t(moveErrorKey(result)))
      }),
    toggleReadOnlyRecommended: () => {
      const { recommended, ...rest } = writeProtection ?? {}
      const next: WriteProtection = recommended ? rest : { ...rest, recommended: true }
      setWriteProtection(Object.keys(next).length > 0 ? next : null)
      setWriteProtectionDirty(true)
      dirtyRef.current = true
    },
    listNumId: (kind, reuse) =>
      (reuse ? firstListNumId(doc?.parsed.blocks ?? EMPTY_BLOCKS, kind) : null) ??
      allocateListNumId(kind),
    openDialog: (dialog) =>
      ({
        table: setShowTableModal,
        link: setShowLinkModal,
        equation: setShowEquationModal,
        bookmark: setShowBookmarkModal,
        crossReference: setShowCrossRefModal,
        caption: setShowCaptionModal,
        document: setShowDocumentDialog,
        changeCase: setShowChangeCase,
        dateTime: setShowDateTime,
        dropCap: setShowDropCap,
        watermark: setShowWatermarkDialog,
        columns: setShowColumnsDialog,
        symbol: setShowSymbolDialog,
        pageNumbers: setShowPageNumbers,
        bulletsNumbering: setShowBulletsNumbering,
        translationLanguage: setShowTranslationLanguage,
        indexTables: setShowIndexTables,
        hyphenation: setShowHyphenation,
        equationOptions: setShowEquationOptions,
        footnote: setShowFootnoteDialog,
        revealFormatting: setShowRevealFormatting,
        recipients: setShowRecipients,
        filterRecipients: setShowFilterRecipients,
        findRecipient: setShowFindRecipient,
        properties: setShowProperties,
        rename: setShowRename,
        bordersShading: setShowBordersShading,
        fontSubstitution: setShowFontSubstitution,
        number: setShowNumber,
        // Word refuses a table with merged cells before its dialog opens
        sort: (open: boolean) => {
          if (editor && sortScope(editor.state) === 'mergedCells')
            setStatus(t(SORT_REFUSALS.mergedCells))
          else setShowSort(open)
        },
        textDirection: setShowTextDirection,
        altText: setShowAltText,
        lineNumbers: setShowLineNumbers,
        icons: setShowIconsDialog,
        paragraphSpacing: setShowParagraphSpacing,
        reduceFileSize: (open: boolean) => {
          setReduceSelection(open && editor ? selectedPictureSources(editor.state) : [])
          setShowReduceFileSize(open)
        },
        sizePosition: setShowSizePosition,
        field: setShowInsertField,
        font: setShowFontDialog,
        paragraph: setShowParaDialog,
        shortcuts: setShowShortcuts,
        customizeKeyboard: setShowCustomizeKeyboard,
        autocorrect: (open: boolean) => {
          if (open) setAutoTextList(autoTextEntries())
          setShowAutoCorrect(open)
        },
        preferences: setShowPreferences,
      })[dialog](true),
    openChartDialog: (type) => setChartModal((open) => ({ type, pick: (open?.pick ?? 0) + 1 })),
    openTableDialog: setTableDialog,
    openPasteSpecial: setPasteSpecialTarget,
    openTabsDialog: setTabsTarget,
    openMarkCitation: () => {
      if (!editor) return
      const { from, to } = editor.state.selection
      setMarkCitationAt({ text: editor.state.doc.textBetween(from, to, ' ').trim(), from, at: to })
    },
    openMarkIndexEntry: () => {
      if (!editor) return
      const { from, to } = editor.state.selection
      setMarkEntryAt({ text: editor.state.doc.textBetween(from, to, ' ').trim(), at: to })
    },
    setIndexSection: (nth, change, wrap) => {
      if (!editor || !doc) return
      const range = indexBlockRanges(editor, doc.parsed.blocks)[nth]
      if (!range) return
      const access = aiPageSetupAccess
      const sectionAt = (block: number) =>
        access.list().find((s) => block >= s.firstBlock && block <= s.lastBlock)
      const blocks = doc.parsed.blocks
      const edges = indexSectionEdges(range, access.list(), (block) => {
        const node = editor.state.doc.maybeChild(block)
        return (
          !!node &&
          node.textContent === '' &&
          xmlOfNode(node as never, blocks).includes('<w:sectPr')
        )
      })
      if (!edges) return
      const { startsIt, endsIt } = edges
      let [first, last] = range
      if (!wrap && !(startsIt && endsIt)) return
      armPageSetupUndo()
      if (!startsIt) {
        if (access.insertBreak('continuous', first - 1)) return
        first++
        last++
      }
      if (!endsIt && access.insertBreak('continuous', last)) return
      const home = sectionAt(first)
      const current = home && access.current(home.index)
      if (home && current) access.set(home.index, { settings: change(current.settings) })
      // as after a section break: the layout follows the sections of the file
      if (doc.filePath) pendingSectionSaveRef.current = true
    },
    autoMarkIndex: () => {
      if (!editor) return
      void pickAndAutoMark(editor).then((result) => {
        if (result.status === 'password') setStatus(t('appDocPwdTitle'))
        else if (result.status === 'failed')
          setStatus(t('ribbonItAutoMarkFailed', { error: result.error }))
        else if (result.status === 'marked')
          setStatus(
            result.count > 0
              ? t('ribbonItAutoMarked', { count: String(result.count) })
              : t('ribbonItAutoMarkNone'),
          )
      })
    },
    insertFile: () => {
      if (!editor) return
      const notesOf = (kind: 'footnote' | 'endnote') =>
        kind === 'footnote' ? reviewCtxRef.current.footnotes : reviewCtxRef.current.endnotes
      void pickAndInsertFile(editor, {
        notes: {
          list: notesOf,
          add: (kind, notes) => {
            const ctx = reviewCtxRef.current
            const next = [...notesOf(kind), ...notes]
            ;(kind === 'footnote' ? ctx.setFootnotes : ctx.setEndnotes)(next)
            ctx.setNotesDirty(true)
          },
        },
        askRange: (fileName, bookmarks) =>
          new Promise((answer) => setInsertFileRange({ fileName, bookmarks, answer })),
      }).then((result) => {
        if (result.status === 'password') setStatus(t('appDocPwdTitle'))
        else if (result.status === 'empty') setStatus(t('ribbonInsertFileEmpty'))
        else if (result.status === 'failed')
          setStatus(t('ribbonInsertFileFailed', { error: result.error }))
        else if (result.status === 'inserted' && result.leftOut > 0)
          setStatus(t('ribbonInsertFileLeftOut', { count: String(result.leftOut) }))
      })
    },
    openAutoText: (mode) => {
      if (!editor) return
      if (mode === 'new' && editor.state.selection.empty) {
        setStatus(t('ribbonAutoTextSelect'))
        return
      }
      setAutoText({ mode, suggestedName: selectionSnippet(editor), entries: autoTextEntries() })
    },
    citationHeading: (category) => t(CITATION_CATEGORY_LABELS[category] ?? 'ribbonToaOther'),
    openLanguage: (target) => {
      if (target.state.selection.empty) setStatus(t('ribbonProofingSelect'))
      else setLanguageTarget(target)
    },
    openTextEffects: setTextEffectsTarget,
    openFitText: (target) => {
      if (textStretch(target.state)) setFitTextTarget(target)
      else setStatus(t('ribbonAsianSelect'))
    },
    toggleAsianLayout: (target, kind) => {
      if (!toggleEaLayout(target, kind)) setStatus(t('ribbonAsianSelect'))
    },
    openPhoneticGuide: (target) => {
      const ruby = rubyTargetOf(target.state)
      if (ruby) setPhoneticGuide({ editor: target, target: ruby })
      else setStatus(t('ribbonPgSelect'))
    },
    openEqLayout: (target, kind) => {
      const found = eqLayoutTargetOf(target.state, kind)
      if (found) setEqLayoutDialog({ editor: target, kind, target: found })
      else setStatus(t(kind === 'enclose' ? 'ribbonEcSelect' : 'ribbonCcSelect'))
    },
    openPictureDialog: (kind, dataUrl) => setPictureDialog({ kind, dataUrl }),
    pickImage: () => window.desktop.pickImage(),
    openFind: (tab) => {
      setShowFind(true)
      ;({ find: setFindFocusInput, replace: setFindFocusReplace, goto: setFindFocusGoTo })[tab](
        (n) => n + 1,
      )
    },
    findStep: (dir) => {
      if (!showFind) {
        setShowFind(true)
        setFindFocusInput((n) => n + 1)
        return
      }
      setFindStep((s) => ({ dir, nonce: (s?.nonce ?? 0) + 1 }))
    },
    showLayoutTab: () => setRibbonTabRequest({ tab: 'layout', nonce: Date.now() }),
    openMailMerge: (mode) => setMergePanel((open) => ({ mode, pick: (open?.pick ?? 0) + 1 })),
    showTableDesignTab: () => setRibbonTabRequest({ tab: 'tableDesign', nonce: Date.now() }),
    defineTableStyle: (style) => {
      upsertStyleDef(
        ribbonStylesRef.current?.has('TableNormal') ? { ...style, basedOn: 'TableNormal' } : style,
      )
    },
    showObjectFormatTab: () => {
      const tab = formatState.imageSelected
        ? 'pictureFormat'
        : formatState.textboxSelected
          ? 'shapeFormat'
          : null
      if (tab) setRibbonTabRequest({ tab, nonce: Date.now() })
      else setStatus(t('ribbonFormatObjectSelect'))
    },
    insertField,
    insertNote,
    editNote,
    noNotes: () => setStatus(t('ribbonNoNotes')),
    blocks: doc?.parsed.blocks ?? EMPTY_BLOCKS,
    sources,
    citationStyle: citationStyle ?? DEFAULT_CITATION_STYLE,
    setCitationStyle: (style) => {
      setCitationStyle(style)
      setSourcesDirty(true)
    },
    headingPages: () => ribbonActions.headingPages(),
    pagesOf: (positions) => {
      const layout = editor && section ? layoutNow() : null
      if (!editor || !layout || layout.cuts.length === 0) return null
      return headingPagesOf(
        layout.blocks,
        layout.cuts,
        positions.map((pos) => editor.view.nodeDOM(pos) as HTMLElement | null),
      ).pages
    },
    review: {
      toggleSpellcheck: () => ribbonActions.onSpellcheck(!spellcheckActive),
      resetIgnoredWords: () =>
        void window.desktop.spellResetIgnored?.().then(() => {
          setRespellNonce((n) => n + 1)
          setStatus(t('ribbonIgnoredWordsReset'))
        }),
      deleteComment: (id) => (id ? deleteComment(id) : ribbonActions.onDeleteComment()),
      resolveComment: (id) => {
        const thread = id ?? commentThreadAtCaretImpl(reviewCtxRef.current)
        if (thread) resolveComment(thread, true)
      },
      replyToComment: (id) => {
        setShowComments(true)
        setCommentFocus({ id, nonce: Date.now(), reply: true })
      },
      deleteAllComments: (resolvedOnly) => ribbonActions.onDeleteAllComments(resolvedOnly),
      gotoComment: (dir) => ribbonActions.onGotoComment(dir),
      toggleComments: () => ribbonActions.onShowComments(),
      toggleTrackChanges: () => ribbonActions.onTrackChanges(!trackChanges),
      revisionDisplay,
      setRevisionDisplay: (mode) => ribbonActions.onRevisionDisplay(mode),
      toggleMarkup: (kind) => {
        const filter = revisionFilterRef.current
        if (typeof kind === 'string') {
          setRevisionFilter({ ...filter, [kind]: !filter[kind] })
          return
        }
        const hidden = filter.hiddenAuthors
        setRevisionFilter({
          ...filter,
          hiddenAuthors: hidden.includes(kind.author)
            ? hidden.filter((a) => a !== kind.author)
            : [...hidden, kind.author],
        })
      },
      setBalloonMode,
      openTrackChangesOptions: () => setShowTcOptions(true),
      openHighlightChanges: () => setShowHighlightChanges(true),
      resolveRevisions: (action, scope) =>
        action === 'accept'
          ? ribbonActions.onAcceptRevision(scope)
          : ribbonActions.onRejectRevision(scope),
      gotoRevision: (dir) => ribbonActions.onGotoRevision(dir),
      toggleReviewPane: () => ribbonActions.onToggleReviewPane(),
      compare: () => setCompareDialog('compare'),
      combine: () => setCompareDialog('combine'),
      protect: () => ribbonActions.onProtectDoc(),
    },
    design: {
      setThemeFonts: (fonts) => ribbonActions.onThemeFonts(fonts),
      setThemeColors: (colors) => ribbonActions.onThemeColors(colors),
      setPageColor: (hex) => ribbonActions.onPageColor(hex),
      setWatermark: (text) => ribbonActions.onWatermark(text),
      setParagraphSpacing: ({ beforeTwips = 0, afterTwips, line }) => {
        const styles = ribbonStylesRef.current
        const styleId = defaultParagraphStyleId(styles) ?? 'Normal'
        upsertStyleDef({
          styleId,
          ...(styles?.has(styleId) ? {} : { builtIn: true }),
          pPr: { spaceBeforeTwips: beforeTwips, spaceAfterTwips: afterTwips, lineSpacing: line },
        })
      },
    },
    view: {
      mode: viewMode,
      setMode: switchView,
      readMode,
      setReadMode: setReadModeExclusive,
      openPagePreview: () => showPreview(true),
      openZoomDialog: () => setShowZoomDialog(true),
      toggleRuler: () => setShowRuler((v) => !v),
      toggleGrid: () => setShowGrid((v) => !v),
      toggleNavigation: () => setShowNav(!showNav),
      toggleThumbnails: () => {
        if (showThumbnails) return setShowNav(false)
        setNavTabRequest({ tab: 'pages', nonce: Date.now() })
        setShowNav(true)
      },
      toggleSplit: () => setSplitView((v) => !v),
      toggleRibbon: () => ribbonActionsRef.current.toggleRibbon?.(),
      toggleFocus: () => {
        if (!focusMode) setStatus(t('ribbonFocusHint'))
        setFocusMode(!focusMode)
      },
      newWindow: () => void window.desktop.openNewTab(doc?.filePath ?? null),
      arrangeAll: () => void window.desktop.arrangeAllWindows(),
      mergeWindows: () => void window.desktop.gatherWindows(),
    },
    headerFooter: {
      header: shownHeader,
      footer: shownFooter,
      setHeader: (next) => ribbonActions.onHeader(next),
      setFooter: (next) => ribbonActions.onFooter(next),
      setFacingPageNumber: (kind, side) => ribbonActions.onFacingPageNumber(kind, side),
      edit: (kind) => ribbonActions.onHfEdit(kind),
      openPageNumberFormat: () => ribbonActions.onPageNumFormat(),
    },
    headerFooterAction: (action) => ribbonActions.onHfAction(action),
    setInkTool,
    clearInk: () => ribbonActions.onInkClearAll(),
    section: sections[activeSection]?.settings ?? section,
    setSection: (next) => ribbonActions.onSection(next),
    setSectionsIn: (scope, change, breakType) =>
      ribbonActions.onSectionsIn(scope, change, breakType),
    setMirrorMargins: (on) => ribbonActions.onMirrorMargins(on),
    setGutterAtTop: (atTop) => ribbonActions.onGutterAtTop(atTop),
    setSectionStart: (type) => ribbonActions.onSectionStart(type),
    insertSectionBreak: (type) => ribbonActions.onInsertSectionBreak(type),
    insertImage: () => editor && void insertImageViaDialog(editor),
    insertPageBreak: () => editor && insertPageBreakAt(editor),
    insertBlankPage: () => editor && insertBlankPageAt(editor),
    newComment: () => {
      setShowComments(true)
      startNewComment()
    },
    wordCount: openStats,
    aiProofread: runAiProofread,
    toggleTableGridlines,
    tableSectionWidthPx,
    notInTable: () => showToast(t('ribbonNotInTable'), 'error'),
  }
  // native application menu → renderer commands
  useEffect(() => window.desktop.onMenuCommand(runCommand), [runCommand])
  useEffect(() => {
    setKeyLayerRunner(runCommand)
    return () => setKeyLayerRunner(null)
  }, [runCommand])

  // Resolve the bookmark anchor against the original block XML, fall back to
  // matching the heading text of a TOC entry.
  const jumpToAnchor = useCallback(
    (anchor: string, tocTitle = '') => {
      let target: Element | null = null
      if (anchor === TOP_ANCHOR) {
        document.querySelector('.ProseMirror')?.firstElementChild?.scrollIntoView({
          behavior: 'smooth',
          block: 'start',
        })
        return
      }
      // the live document first: a bookmark or heading anchor added since the file was read
      const livePos = anchor && editor ? anchorPos(editor.state.doc, anchor) : null
      if (livePos !== null && editor) {
        const dom = editor.view.nodeDOM(livePos)
        if (dom instanceof Element) target = dom
      }
      if (!target && anchor && doc) {
        const block = doc.parsed.blocks.find((b) => b.originalXml?.includes(`w:name="${anchor}"`))
        if (block && block.docxIndex !== null) {
          target = document.querySelector(`.ProseMirror [data-idx="${block.docxIndex}"]`)
        }
      }
      const title = tocTitle.replace(/\s+/g, '')
      if (!target && title) {
        target =
          [
            ...document.querySelectorAll(
              '.ProseMirror h1, .ProseMirror h2, .ProseMirror h3, .ProseMirror h4, .ProseMirror h5, .ProseMirror h6',
            ),
          ].find((h) => (h.textContent ?? '').replace(/\s+/g, '') === title) ?? null
      }
      target?.scrollIntoView({ behavior: 'smooth', block: 'start' })
    },
    [doc, editor],
  )

  /** ⌘-click / Open Hyperlink: browser or mail program (main-process openExternal), in-document jump for #bookmark */
  const openLink = useCallback(
    (href: string, tocTitle = '') => {
      const target = linkTarget(href)
      if (target?.kind === 'external') window.open(target.url)
      else if (target?.kind === 'anchor' || tocTitle) {
        jumpToAnchor(target?.kind === 'anchor' ? target.name : '', tocTitle)
      }
    },
    [jumpToAnchor],
  )

  // Word: links and TOC entries jump on ⌘/Ctrl+click only; a plain click just
  // places the caret.
  const onDocClick = useCallback(
    (e: ReactMouseEvent) => {
      const tag = (e.target as HTMLElement).closest('.doc-comment-marker') as HTMLElement | null
      if (tag?.dataset.commentId) {
        setShowComments(true)
        setCommentFocus({ id: tag.dataset.commentId, nonce: Date.now() })
        return
      }
      const commentSpan = (e.target as HTMLElement).closest('.doc-comment') as HTMLElement | null
      if (commentSpan) {
        const ids = (commentSpan.dataset.commentIds ?? '').split(' ')
        const hit = comments.find((c) => !c.done && !c.parentId && ids.includes(c.id))
        if (hit) {
          setShowComments(true)
          setCommentFocus({ id: hit.id, nonce: Date.now() })
        }
      }
      const mod = e.metaKey || e.ctrlKey
      // read-only display anchors are outside contenteditable and would navigate
      // the renderer in place
      const a = (e.target as HTMLElement).closest('a[href]') as HTMLAnchorElement | null
      if (a) {
        e.preventDefault()
        if (mod) openLink(a.getAttribute('href') ?? '')
        return
      }
      if (!mod) return
      // a cross-reference inserted as a hyperlink (\\h) goes to what it refers to
      const field = (e.target as HTMLElement).closest('[data-instr-field]') as HTMLElement | null
      const linked = fieldLinkBookmark(field?.dataset.instrField ?? '')
      if (linked) {
        // to a note, the caret goes with it, as in Word: left on the reference, it would scroll the page back
        const at = editor ? anchorPos(editor.state.doc, linked) : null
        if (editor && at !== null && editor.state.doc.nodeAt(at)?.type.name === 'docNoteRef')
          editor.chain().setNodeSelection(at).scrollIntoView().run()
        return jumpToAnchor(linked)
      }
      const line = (e.target as HTMLElement).closest(
        '.doc-toc-line:not([data-toc-nolink])',
      ) as HTMLElement | null
      if (line) jumpToAnchor(line.dataset.tocAnchor ?? '', line.dataset.tocTitle ?? '')
    },
    [comments, openLink, jumpToAnchor, editor],
  )

  /** Editor right-click menu, only inside the document body */
  const onDocContextMenu = useCallback(
    (e: ReactMouseEvent) => {
      if (readMode || (isProtected && !commentsAllowed)) return
      const bubble = (e.target as HTMLElement).closest(
        '.comment-bubble, .comment-balloon, .doc-del-marker',
      ) as HTMLElement | null
      if (!bubble && !(e.target as HTMLElement).closest('.doc-page')) return
      // no preventDefault: Blink must still request a context menu so the main
      // process receives the misspelling data it relays to this menu
      const seq = ++ctxSeq.current
      window.desktop.claimContextMenu?.(seq)
      // keyboard-invoked (Shift+F10 / menu key): no pointer, open at the caret
      const fromKeyboard = e.clientX === 0 && e.clientY === 0
      let x = e.clientX
      let y = e.clientY
      let pos: number | null = null
      if (fromKeyboard && editor) {
        pos = editor.state.selection.head
        const c = editor.view.coordsAtPos(pos)
        x = c.left
        y = c.bottom
      }
      let imageSrc: string | null = null
      if (bubble && editor) {
        // a balloon stands for its anchor: the menu acts on that change or thread
        const commentId = bubble.dataset.commentId ?? null
        const at = commentId
          ? commentAnchors(editor.state.doc).find((a) => a.id === commentId)?.from
          : bubble.classList.contains('doc-del-marker')
            ? editor.view.posAtDOM(bubble, 0)
            : Number(bubble.dataset.pos)
        if (at != null && Number.isFinite(at)) {
          const size = editor.state.doc.content.size
          pos = Math.min(at, size)
          // Word selects the whole change so New Comment anchors on it
          const end = Math.min(Number(bubble.dataset.end), size)
          const doc = editor.state.doc
          // between() snaps ends that fall outside a textblock (deleted paragraph marks, tables)
          const sel =
            Number.isFinite(end) && end > pos
              ? TextSelection.between(doc.resolve(pos), doc.resolve(end))
              : TextSelection.near(doc.resolve(pos), 1)
          editor.view.dispatch(editor.state.tr.setSelection(sel))
        }
        setCtxMenu({ x, y, seq, pos, imageSrc, link: null, commentId, balloon: true })
        return
      }
      // Word behavior: right-clicking outside the selection moves the cursor there first (menu items act on the clicked block)
      if (editor) {
        // Right-clicking directly on an image / floating object selects it as a
        // node (Word shows Wrap Text / Position on a plain right-click), so the
        // context menu can offer wrap + z-order without a prior left-click.
        let protectedEl = (e.target as HTMLElement).closest(
          ".doc-protected[data-doc-protected='image'], .doc-protected.doc-img-float",
        ) as HTMLElement | null
        // behind-text pictures live under the text layer's hit box: when the
        // press hits no glyph, fall through to the picture painted below
        // (Word selects it there too)
        if (!protectedEl && !fromKeyboard) {
          const under = findFloatImageAt(x, y)
          if (under) protectedEl = under.closest('.doc-protected') as HTMLElement | null
        }
        const img = (e.target as HTMLElement).closest('img') ?? protectedEl?.querySelector('img')
        imageSrc = img?.getAttribute('src') || null
        let selectedNode = false
        // a floating run picture is its own node, outside any .doc-protected block
        const runPic =
          !protectedEl && !fromKeyboard ? runPictureAtEvent(editor.view, e.target, x, y) : null
        if (runPic) {
          editor.view.dispatch(
            editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, runPic.pos)),
          )
          imageSrc = runPic.img.getAttribute('src') || imageSrc
          selectedNode = true
        } else if (protectedEl) {
          const dom = editor.view.nodeDOM.bind(editor.view)
          let nodePos = -1
          editor.state.doc.descendants((node, pos) => {
            if (nodePos !== -1) return false
            if (node.type.name === 'docProtected' && dom(pos) === protectedEl) nodePos = pos
            return nodePos === -1
          })
          if (nodePos !== -1) {
            const tr = editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, nodePos))
            editor.view.dispatch(tr)
            selectedNode = true
          }
        }
        if (!selectedNode && !fromKeyboard) {
          const hit = editor.view.posAtCoords({ left: x, top: y })
          if (hit) {
            pos = hit.pos
            // every range: a selection of cells is one range a cell, and stays when the press lands in any
            const { selection } = editor.state
            const inside = selection.ranges.some(
              ({ $from, $to }) => hit.pos >= $from.pos && hit.pos <= $to.pos,
            )
            if (!inside) editor.commands.setTextSelection(hit.pos)
            else if (selection instanceof CellSelection) keepCellSelection(editor, selection)
          }
        }
      }
      const anchorEl = (e.target as HTMLElement).closest('a[href]') as HTMLAnchorElement | null
      const tocLine = (e.target as HTMLElement).closest(
        '.doc-toc-line:not([data-toc-nolink])',
      ) as HTMLElement | null
      const link = anchorEl
        ? { href: anchorEl.getAttribute('href') ?? '' }
        : tocLine
          ? {
              href: `#${tocLine.dataset.tocAnchor ?? ''}`,
              toc: true,
              tocTitle: tocLine.dataset.tocTitle ?? '',
            }
          : null
      setCtxMenu({ x, y, seq, pos, imageSrc, link })
    },
    [readMode, isProtected, commentsAllowed, editor],
  )

  /** double-click on a page-gap strip: edit that page's header/footer in place */
  const startGapHfEdit = (el: HTMLElement) => {
    if (readMode || isProtected || hfEditorRef.current) return
    const kind = el.dataset.hfKind as 'header' | 'footer' | undefined
    if (!kind) return
    const si = Number(el.dataset.hfSection ?? 0)
    const variant = (el.dataset.hfVariant as HfView | undefined) ?? 'default'
    const pageIdx = Number(el.dataset.hfPageIdx ?? 0)
    const page = el.dataset.hfPage ?? '1'
    const resolved = hfResolveAt(si, kind, variant)
    const value = resolved.value ?? { text: '' }
    let host = el
    let saved = Array.from(el.childNodes)
    el.replaceChildren()
    el.classList.add('page-hf-editing')
    let committed = false
    const handle = mountHfEditor(el, {
      value,
      pageNo: page,
      pageTotal: el.dataset.hfTotal ?? '1',
      sectionPages: el.dataset.hfSectionPages,
      spellcheck: spellcheckActive,
      onCommit: (next) => {
        committed = true
        // the ribbon may have unlinked the section since the editor opened
        commitHfAtRef.current(si, kind, variant, next, resolved.owner)
      },
      onExit: () => {
        host.classList.remove('page-hf-editing')
        if (!committed) host.replaceChildren(...saved)
        if (hfEditorRef.current === handle) hfEditorRef.current = null
        setHfEditing(null)
      },
    })
    hfEditorRef.current = handle
    setHfEditing({ kind, section: si, variant, pageIdx, owner: resolved.owner })
    // a re-pagination (Link to Previous, distance, another strip's commit) rebuilds the
    // gap widgets: the open editor moves into the strip's new element instead of dying
    const selector = `.page-gap-hf.page-hf-${kind}[data-hf-page-idx="${pageIdx}"]`
    const watch = () => {
      if (hfEditorRef.current !== handle) return
      if (!host.isConnected) {
        const next = document.querySelector<HTMLElement>(selector)
        if (!next) {
          handle.exit()
          return
        }
        saved = Array.from(next.childNodes)
        next.replaceChildren(handle.editor.view.dom)
        next.classList.add('page-hf-editing')
        host = next
        // detaching dropped the focus with the old widget
        handle.editor.view.focus()
      }
      requestAnimationFrame(watch)
    }
    requestAnimationFrame(watch)
  }
  const startGapHfEditRef = useRef(startGapHfEdit)
  startGapHfEditRef.current = startGapHfEdit
  const applyStoryRevisionsRef = useRef(applyStoryRevisions)
  applyStoryRevisionsRef.current = applyStoryRevisions
  const commitHfAtRef = useRef(commitHfAt)
  commitHfAtRef.current = commitHfAt
  const firstPageNoText =
    firstPageNo ?? formatPageNumber(sections[0]?.pageNumberStart ?? 1, sections[0]?.pageNumberFmt)
  const lastPageNoText = lastPageNo?.text ?? String(pageInfo.total)
  const lastPageIdx = Math.max(0, pageInfo.total - 1)
  /** open the strip editor for `kind` on page `pageIdx` (default: the document's first header / last footer) */
  const openHfEditor = (kind: 'header' | 'footer', pageIdx?: number) => {
    hfEditorRef.current?.exit()
    // only Print Layout shows the strips
    if (viewMode !== 'print') switchView('print')
    const edge =
      pageIdx === undefined || (kind === 'header' ? pageIdx === 0 : pageIdx === lastPageIdx)
    const el = edge
      ? null
      : document.querySelector<HTMLElement>(
          `.page-gap-hf.page-hf-${kind}[data-hf-page-idx="${pageIdx}"]`,
        )
    if (el) startGapHfEdit(el)
    else setHfAreaEditRequest({ kind, nonce: Date.now() })
  }
  const openHfEditorRef = useRef(openHfEditor)
  openHfEditorRef.current = openHfEditor
  // floating header pictures: moved and resized in header editing, reached from
  // the body by a double-click (Word)
  const hfFloatEditRef = useRef({ open: openHfEditor, commit: commitHfPictureAt, editable: false })
  hfFloatEditRef.current = {
    open: openHfEditor,
    commit: commitHfPictureAt,
    editable: !readMode && !isProtected,
  }
  useEffect(() => {
    const root = scrollContainerRef.current
    if (!root) return
    return installHfFloatEdit({
      root,
      editing: () => hfEditingRef.current != null,
      editable: () => hfFloatEditRef.current.editable,
      open: (pageIdx) => hfFloatEditRef.current.open('header', pageIdx),
      commit: (t, edit) =>
        hfFloatEditRef.current.commit(t.section, 'header', t.variant, t.drawing, edit),
    })
  }, [])
  const toggleEvenOddHf = (on: boolean) => {
    hfEditorRef.current?.exit()
    armPageSetupUndo()
    setEvenOddHf(on)
    setEvenOddHfDirty(true)
    setStatus(on ? t('appEvenOddOn') : t('appEvenOddOff'))
  }
  const onHfAction = (action: HfAction) => {
    const cur = hfEditing
    const sub = hfEditorRef.current?.editor
    // the ribbon button took the focus; Word keeps the caret in the strip
    const refocus = () => {
      if (hfEditorRef.current) sub?.commands.focus()
    }
    const sectionOf = (kind: 'header' | 'footer') =>
      cur?.section ?? (kind === 'header' ? 0 : lastSectionIdx)
    const variant = cur?.variant ?? 'default'
    switch (action.type) {
      case 'close':
        hfEditorRef.current?.exit()
        break
      case 'goto':
      case 'edit':
        openHfEditor(action.kind, cur?.pageIdx)
        break
      case 'remove':
        hfEditorRef.current?.exit()
        commitHfAt(sectionOf(action.kind), action.kind, variant, { text: '' })
        break
      case 'pageNumber':
        hfEditorRef.current?.exit()
        commitHfAt(sectionOf(action.kind), action.kind, variant, hfPageNumberOnly(action.align))
        break
      case 'pageNumberHere':
        if (sub) insertHfField(sub, 'PAGE')
        break
      case 'numPages':
        if (sub) insertHfField(sub, 'NUMPAGES')
        break
      case 'pageNumberFormat':
        openPgNumModal()
        break
      case 'removePageNumbers':
        hfEditorRef.current?.exit()
        for (const kind of ['header', 'footer'] as const) {
          const v = hfValueAt(sectionOf(kind), kind, variant)
          if (hfHasPageField(v)) commitHfAt(sectionOf(kind), kind, variant, hfWithoutPageMarks(v!))
        }
        break
      case 'field':
        if (sub) insertHfInstrField(sub, action.instr, fieldValue(action.instr))
        break
      case 'picture':
        if (sub) {
          void (async () => {
            const picked = await window.desktop.pickImage()
            if (!picked || !hfEditorRef.current) return
            const dataUrl = `data:${picked.mime};base64,${picked.base64}`
            const natural = await imageSizeOf(dataUrl)
            insertHfPicture(sub, { dataUrl, widthPx: natural.width, heightPx: natural.height })
          })()
        }
        break
      case 'linkToPrevious':
        if (cur) setHfLink(cur.section, cur.kind, cur.variant, action.on)
        if (!action.on) refocus()
        break
      case 'distance':
        setHfDistance(cur?.section ?? activeSection, action.kind, action.twips)
        refocus()
        break
      case 'titlePg':
        toggleTitlePg(action.on)
        refocus()
        break
      case 'evenOddHf':
        toggleEvenOddHf(action.on)
        refocus()
        break
    }
  }
  const hfEditingInfo: HfEditingInfo | null = hfEditing
    ? {
        kind: hfEditing.kind,
        section: hfEditing.section,
        sectionCount: hfSections.length,
        variant: hfEditing.variant,
        linked: hfLinkedAt(hfEditing.section, hfEditing.kind, hfEditing.variant),
        headerDistTwips: (hfSections[hfEditing.section]?.settings ?? section)?.headerDist ?? 720,
        footerDistTwips: (hfSections[hfEditing.section]?.settings ?? section)?.footerDist ?? 720,
        titlePg,
        evenOddHf,
      }
    : null
  const hfSectionLabel = (kind: 'header' | 'footer', si: number) =>
    hfSections.length > 1
      ? t(kind === 'header' ? 'appHfHeaderSectionTag' : 'appHfFooterSectionTag', { n: si + 1 })
      : null

  // Word: double-clicking a picture only selects it (the press already did);
  // the image viewer stays on the context menu
  const onDocDoubleClick = useCallback((e: ReactMouseEvent) => {
    const target = e.target as HTMLElement
    if (!target.closest('.doc-page')) return
    const gap = target.closest('.page-gap-hf') as HTMLElement | null
    if (gap) startGapHfEditRef.current(gap)
  }, [])

  // genoffice CLI (`open --block`, `selection`): the shell evaluates this hook
  useEffect(() => {
    ;(window as unknown as Record<string, unknown>).__genofficeControl = (req: ControlRequest) =>
      handleDocsControl(req, editor, doc !== null)
  })

  useEffect(() => (editor ? installSelectionBar(editor) : undefined), [editor])

  // e2e/automation hook: lets tests drive open/edit/save without native dialogs
  useEffect(() => {
    ;(window as unknown as Record<string, unknown>).__aidocs = {
      editor,
      openPath: (path: string) => openRecent(path),
      save: () => save(false),
      getStatus: () => status,
      exportPdfTo: (path: string) => exportPdf(path),
      exportHtmlTo: (path: string) => exportHtml(path),
      exportMarkdownTo: (path: string) => exportMarkdown(path),
      subEditor: () => getActiveSubEditor(),
    }
  }, [editor, openRecent, save, status, exportPdf, exportHtml, exportMarkdown])

  // Pending style definitions (Styles pane / AI define_style) show before the
  // save writes them: styles.xml is re-resolved with the patches applied and
  // the live map feeds the gallery, the pane, the format state and the doc CSS.
  const [liveStyles, setLiveStyles] = useState<{
    parsed: ParsedDocFull
    styles: Map<string, StyleInfo>
  } | null>(null)
  useEffect(() => {
    const parsed = doc?.parsed
    const upserts = Object.values(styleUpserts)
    if (!parsed || upserts.length === 0) {
      liveStylesRef.current = null
      setLiveStyles(null)
      return
    }
    let cancelled = false
    previewFontSettings(parsed, upserts, defaultFonts)
      .then((r) => {
        // a document switch or a save re-parse while resolving: their open path already set fresh styles
        if (cancelled || fileCtxRef.current.doc?.parsed !== parsed) return
        liveStylesRef.current = { parsed, styles: r.styles }
        setLiveStyles({ parsed, styles: r.styles })
        editor.storage.listNumbering.styles = r.styles
        setDocCss(liveDocCss({ ...parsed, styles: r.styles }))
      })
      .catch(() => {})
    return () => {
      cancelled = true
    }
  }, [doc, styleUpserts, defaultFonts, editor, liveDocCss])
  // shallow-stable snapshot of every editor read the ribbon displays: caret moves
  // that change none of it keep the reference, so the memoized Ribbon skips
  const formatState = useShallowStable(
    computeFormatState(
      editor,
      liveStyles?.parsed === doc?.parsed ? liveStyles?.styles : doc?.parsed.styles,
      doc?.parsed.docDefaults,
      doc?.parsed.blocks,
    ),
  )

  const ribbonStyles = useMemo(
    () =>
      doc
        ? liveStyles?.parsed === doc.parsed
          ? liveStyles.styles
          : new Map(doc.parsed.styles)
        : undefined,
    [doc, liveStyles],
  )
  const ribbonStylesRef = useRef(ribbonStyles)
  ribbonStylesRef.current = ribbonStyles
  const gridBasePt = useMemo(
    () => gridBaseSizePt(ribbonStyles, doc?.parsed.docDefaults),
    [ribbonStyles, doc],
  )
  gridBasePtRef.current = gridBasePt
  const snapsToChars = useMemo(
    () => sections.some((s) => s.settings.docGrid?.type === 'snapToChars'),
    [sections],
  )
  const charCellPt = useMemo(() => docCharCellPt(sections, gridBasePt), [sections, gridBasePt])
  const gridFont = useMemo(() => {
    const d = ribbonStyles?.get(defaultParagraphStyleId(ribbonStyles) ?? '')?.display
    const dd = doc?.parsed.docDefaults
    return {
      latin: d?.fontAscii ?? d?.font ?? dd?.asciiFont ?? '',
      eastAsia: d?.eastAsiaFont ?? dd?.eastAsiaFont ?? '',
      sizeHalfPoints: Math.round(gridBasePt * 2),
      bold: (d?.bold ?? dd?.bold) === true,
      italic: (d?.italic ?? dd?.italic) === true,
    }
  }, [ribbonStyles, doc, gridBasePt])

  /** every function prop of the memoized Ribbon, with stable identities (dispatches into the latest render's closures) */
  // ---- selection-scoped AI edit queue ----
  const getQueueItem = useCallback(
    (qid: string) => editQueueRef.current.find((item) => item.qid === qid),
    [],
  )
  const queueAdd = (instruction: string): void => {
    const { from, to, empty } = editor.state.selection
    if (empty || editQueueRef.current.length >= EDIT_QUEUE_MAX) return
    const qid = `q${++queueSeqRef.current}`
    addQueueAnchor(editor, qid, from, to)
    const capturedText = editor.state.doc
      .textBetween(from, to, ' ', ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 80)
    setEditQueue((queue) => [...queue, { qid, instruction, capturedText }])
  }
  const queueUpdate = (qid: string, instruction: string): void =>
    setEditQueue((queue) => queue.map((i) => (i.qid === qid ? { ...i, instruction } : i)))
  const queueRemove = (qid: string): void => {
    removeQueueAnchors(editor, [qid])
    setEditQueue((queue) => queue.filter((i) => i.qid !== qid))
  }
  const queueClear = (): void => {
    clearQueueAnchors(editor)
    setEditQueue([])
  }
  /** a submission hands its items to the run and drops them from the queue */
  const queueConsume = (qids: string[]): void => {
    removeQueueAnchors(editor, qids)
    setEditQueue((queue) => queue.filter((i) => !qids.includes(i.qid)))
  }
  const queueFocus = (qid: string): void => {
    const selection = selectionForAnchor(editor, qid)
    if (!selection) return
    editor.view.dispatch(editor.state.tr.setSelection(selection).scrollIntoView())
    editor.view.focus()
  }
  const askSendNow = (text: string): void => {
    setShowAi(true)
    setAiPreset({ text, nonce: Date.now(), autoRun: true })
  }

  // AI comment tools run the same review-actions code paths as the comments pane
  const aiCommentsAccess = useMemo<AiCommentsAccess>(
    () => ({
      list: () => reviewCtxRef.current.comments,
      reply: (parentId, text) =>
        replyToCommentImpl(reviewCtxRef.current, parentId, text, aiRevisionAuthor()),
      resolve: (id) => {
        const ctx = reviewCtxRef.current
        if (!ctx.comments.some((c) => c.id === id)) return false
        resolveCommentImpl(ctx, id, true)
        return true
      },
      add: (range, text, meta) =>
        addCommentAtImpl(
          reviewCtxRef.current,
          range,
          text,
          meta.author ?? aiRevisionAuthor(),
          meta.initials,
        ),
      remove: (id) => {
        const ctx = reviewCtxRef.current
        if (!ctx.comments.some((c) => c.id === id)) return false
        deleteCommentImpl(ctx, id)
        return true
      },
    }),
    [],
  )

  // AI note tools: the lists mirror this render's writes so two inserts in one
  // agent batch mint distinct ids before React state catches up
  const aiNotesMirror = useRef<Record<'footnote' | 'endnote', NoteInfo[] | null>>({
    footnote: null,
    endnote: null,
  })
  useEffect(() => {
    aiNotesMirror.current = { footnote: null, endnote: null }
  }, [footnotes, endnotes])
  const aiNotesAccess = useMemo<AiNotesAccess>(() => {
    const list = (kind: 'footnote' | 'endnote') =>
      aiNotesMirror.current[kind] ??
      (kind === 'footnote' ? reviewCtxRef.current.footnotes : reviewCtxRef.current.endnotes)
    const commit = (kind: 'footnote' | 'endnote', next: NoteInfo[]) => {
      aiNotesMirror.current[kind] = next
      ;(kind === 'footnote' ? reviewCtxRef.current.setFootnotes : reviewCtxRef.current.setEndnotes)(
        next,
      )
      reviewCtxRef.current.setNotesDirty(true)
    }
    return {
      list,
      add: (kind, text) => {
        const id = nextNoteId(list(kind))
        commit(kind, [...list(kind), { id, text, created: true }])
        return id
      },
      remove: (kind, id) => {
        const current = list(kind)
        if (!current.some((n) => n.id === id)) return false
        commit(
          kind,
          current.filter((n) => n.id !== id),
        )
        return true
      },
      replace: (kind, id, next) => {
        const current = list(kind)
        if (!current.some((n) => n.id === id)) return false
        commit(
          kind,
          current.map((n) => (n.id === id ? next : n)),
        )
        return true
      },
      protectedMarkBlock: (kind, id) => {
        const { editor, doc } = reviewCtxRef.current
        if (!editor || !doc) return null
        return protectedNoteMarkBlock(
          editor.state.doc,
          (block) =>
            typeof block.attrs.genXml === 'string'
              ? block.attrs.genXml
              : typeof block.attrs.docxIndex === 'number'
                ? (doc.parsed.blocks[block.attrs.docxIndex]?.originalXml ?? '')
                : '',
          kind,
          id,
        )
      },
    }
  }, [])

  // AI header/footer tool: reads the live HF state and writes through the same
  // commit path as on-canvas editing (variant routing, per-section edits, dirty flags).
  // The overlay mirrors this render's AI writes: an agent batch runs several tools
  // between renders, so a read right after a set must not see the pre-write state
  // (same pitfall as the comments id-minting ref mirror). Recreated per render,
  // by which time React state has caught up.
  const aiHfCtx = {
    overlay: new Map<string, HeaderFooter>(),
    valueOf(kind: 'header' | 'footer', view: HfView): HeaderFooter | null {
      const pending = this.overlay.get(`${kind}:${view}`)
      if (pending) return pending
      return hfValueAt(kind === 'header' ? 0 : lastSectionIdx, kind, view)
    },
    commit: commitHf,
    titlePg,
    evenOddHf,
    multiHf,
    locked: isProtected || readMode,
  }
  const aiHfCtxRef = useRef(aiHfCtx)
  aiHfCtxRef.current = aiHfCtx
  // AI style / watermark tools read and write the same pending stores the save path drains.
  // The upserts ref mirrors the state synchronously: an agent turn runs several style tools
  // between renders, so a define_style followed by applyStyle must see the pending entry.
  const aiStyleUpsertsRef = useRef(styleUpserts)
  aiStyleUpsertsRef.current = styleUpserts
  /** queue a style definition for the next save; shared by the AI tools and the Styles pane */
  const upsertStyleDef = useCallback((up: StyleUpsert): string | null => {
    const ctx = fileCtxRef.current
    if (!ctx.doc) return 'no document is open'
    const prev = aiStyleUpsertsRef.current[up.styleId]
    // a removal or a copied definition stands alone; a style made again after its removal starts clean
    const whole = up.remove || up.sourceXml !== undefined
    const next = {
      ...aiStyleUpsertsRef.current,
      [up.styleId]: whole
        ? up
        : prev?.remove
          ? { ...up, replace: true as const }
          : prev
            ? {
                ...prev,
                ...up,
                pPr: up.pPr || prev.pPr ? { ...prev.pPr, ...up.pPr } : undefined,
                rPr: up.rPr || prev.rPr ? { ...prev.rPr, ...up.rPr } : undefined,
              }
            : up,
    }
    aiStyleUpsertsRef.current = next
    ctx.setStyleUpserts(next)
    return null
  }, [])
  const aiDocExtras = useMemo<AiDocExtras>(
    () => ({
      styles: {
        list: () => {
          const ctx = fileCtxRef.current
          const out = new Map<string, AiStyleInfo>()
          for (const s of ctx.doc?.parsed.styles.values() ?? []) {
            if (s.linkedCharShell) continue
            out.set(s.styleId, {
              styleId: s.styleId,
              name: s.name,
              type: s.type,
              ...(s.basedOn ? { basedOn: s.basedOn } : {}),
              ...(s.headingLevel ? { headingLevel: s.headingLevel } : {}),
            })
          }
          const upserts = aiStyleUpsertsRef.current
          const parsed = (id: string) => ctx.doc?.parsed.styles.get(id)
          for (const up of Object.values(upserts)) {
            if (up.remove) {
              out.delete(up.styleId)
              continue
            }
            const cur = out.get(up.styleId)
            const headingLevel = pendingHeadingLevel(up.styleId, (id) => upserts[id], parsed)
            out.set(up.styleId, {
              styleId: up.styleId,
              name: up.name ?? cur?.name ?? up.styleId,
              type: cur?.type ?? up.type ?? 'paragraph',
              ...(up.basedOn === undefined
                ? cur?.basedOn
                  ? { basedOn: cur.basedOn }
                  : {}
                : up.basedOn
                  ? { basedOn: up.basedOn }
                  : {}),
              ...(headingLevel ? { headingLevel } : {}),
              pending: true,
            })
          }
          return [...out.values()]
        },
        upsert: upsertStyleDef,
      },
      watermark: {
        current: () => fileCtxRef.current.watermark,
        set: (spec) => {
          if (!fileCtxRef.current.doc) return 'no document is open'
          applyWatermark(spec)
          return null
        },
      },
      revisionsLocked: isProtected || trackChangesForced,
      storyRevisions: {
        list: () => storyRevisionsRef.current,
        apply: (targets, mode, all) => applyStoryRevisionsRef.current(targets, mode, all),
      },
    }),
    [upsertStyleDef, applyWatermark, isProtected, trackChangesForced],
  )

  const aiHfAccess = useMemo<AiHeaderFooterAccess>(
    () => ({
      read: () => {
        const ctx = aiHfCtxRef.current
        const textOf = (kind: 'header' | 'footer', view: HfView) => {
          const value = ctx.valueOf(kind, view)
          return value ? hfEditText(value) : ''
        }
        return {
          header: textOf('header', 'default'),
          footer: textOf('footer', 'default'),
          headerFirst: ctx.titlePg ? textOf('header', 'first') : null,
          footerFirst: ctx.titlePg ? textOf('footer', 'first') : null,
          headerEven: ctx.evenOddHf ? textOf('header', 'even') : null,
          footerEven: ctx.evenOddHf ? textOf('footer', 'even') : null,
          titlePg: ctx.titlePg,
          evenOddHf: ctx.evenOddHf,
          multiSection: ctx.multiHf,
        }
      },
      set: (kind, view, text) => {
        const ctx = aiHfCtxRef.current
        if (ctx.locked) return 'the document is read-only; headers/footers cannot be edited'
        if (view === 'first' && !ctx.titlePg) {
          setTitlePg(true)
          setTitlePgDirty(true)
          ctx.titlePg = true
        }
        if (view === 'even' && !ctx.evenOddHf) {
          setEvenOddHf(true)
          setEvenOddHfDirty(true)
          ctx.evenOddHf = true
        }
        const next = applyHfText(ctx.valueOf(kind, view), text)
        ctx.commit(kind, next, view)
        ctx.overlay.set(`${kind}:${view}`, next)
        return null
      },
    }),
    [],
  )

  // AI page-setup tool: reads the live section list (block ranges in PM
  // indexes) and writes through the Layout tab's paths. The mirror lets a
  // batch of tool calls between renders see its own writes: `sections` as they
  // stand after each call and the sectPr each would save to.
  const aiPageCtx = {
    sections: sections.map((s) => ({ ...s })),
    effective: sections.map((_s, i) => effectiveSectPrXml(i) ?? ''),
    locked: isProtected || readMode,
    isBlank: !doc,
  }
  const aiPageCtxRef = useRef(aiPageCtx)
  aiPageCtxRef.current = aiPageCtx
  const aiPageSetupAccess = useMemo<AiPageSetupAccess>(
    () => ({
      list: () => {
        const ctx = aiPageCtxRef.current
        if (!editor || ctx.isBlank || ctx.sections.length === 0) return []
        const pmDoc = editor.state.doc
        const ranges = ctx.sections.map(() => ({ first: -1, last: -1 }))
        for (let i = 0; i < pmDoc.childCount; i++) {
          const r = ranges[sectionIndexAtBlock(pmDoc, ctx.sections, i)]!
          if (r.first < 0) r.first = i
          r.last = i
        }
        return ctx.sections.map((sec, i) =>
          describeSection(sec, i, Math.max(ranges[i]!.first, 0), Math.max(ranges[i]!.last, 0)),
        )
      },
      current: (index) => aiPageCtxRef.current.sections[index],
      set: (index, resolved) => {
        const ctx = aiPageCtxRef.current
        if (ctx.locked) return 'the document is read-only; the page setup cannot be changed'
        const sec = ctx.sections[index]
        if (!sec || !editor) return `section ${index} does not exist`
        const nextXml = applyResolvedPageSetup(ctx.effective[index] ?? sec.sectPrXml, resolved)
        const next: SectionInfo = {
          ...sec,
          settings: resolved.settings,
          titlePg: resolved.titlePg ?? sec.titlePg,
          ...(resolved.pgNum
            ? { pageNumberStart: resolved.pgNum.start, pageNumberFmt: resolved.pgNum.fmt }
            : {}),
        }
        if (sec.pendingBreak) {
          if (!patchPendingSectPr(editor, ctx.sections, index, nextXml))
            return `section ${index} was removed from the document`
          next.sectPrXml = nextXml
        }
        ctx.sections[index] = next
        ctx.effective[index] = nextXml
        setSections((prev) => prev.map((s, i) => (i === index ? next : s)))
        if (sec.pendingBreak) return null
        if (ctx.sections.length <= 1 || index === ctx.sections.length - 1) {
          setSection(resolved.settings)
          setSectionDirty(true)
          if (resolved.titlePg !== undefined) {
            setTitlePg(resolved.titlePg)
            setTitlePgDirty(true)
          }
          const pgNum = resolved.pgNum
          if (pgNum) setPgNumEdit((pending) => mergePendingPgNum(pending, pgNum))
        } else {
          setSectionsDirty((d) => (d.includes(index) ? d : [...d, index]))
          if (resolved.pgNum) setPgNumDirtySections((d) => (d.includes(index) ? d : [...d, index]))
        }
        return null
      },
      insertBreak: (type, afterBlockIndex) => {
        const ctx = aiPageCtxRef.current
        if (ctx.locked) return 'the document is read-only; a section break cannot be inserted'
        if (!editor || ctx.isBlank) return 'no document is open'
        const pmDoc = editor.state.doc
        const ownerIdx = sectionIndexAtBlock(pmDoc, ctx.sections, Math.max(afterBlockIndex, 0))
        const owner = ctx.sections[ownerIdx]
        const copyXml = ctx.effective[ownerIdx]
        insertSectionBreakRef.current(
          type,
          afterBlockIndex,
          copyXml ? { sectPr: copyXml, sections: ctx.sections } : undefined,
        )
        if (!owner || !copyXml) return null
        // insertSectionBreak committed the new section to state; the mirror follows
        const created = pendingSectionAt(pmDoc, owner, afterBlockIndex, copyXml)
        // the remainder starts with the chosen type; a pending owner's sectPr was
        // patched in place by insertSectionBreak, the others get it written on save
        const remainderXml = applySectionStartType(copyXml, type)
        const ownerNext: SectionInfo = {
          ...owner,
          firstBlockIndex: created.lastBlockIndex + 1,
          startType: type,
          ...(owner.pendingBreak ? { sectPrXml: remainderXml } : {}),
        }
        ctx.sections.splice(ownerIdx, 0, created)
        ctx.sections[ownerIdx + 1] = ownerNext
        ctx.effective.splice(ownerIdx, 0, copyXml)
        ctx.effective[ownerIdx + 1] = remainderXml
        return null
      },
    }),
    [editor],
  )

  const gutterAtTop = gutterAtTopEdit ?? doc?.parsed.gutterAtTop ?? false
  const activeStartType: SectionInfo['startType'] =
    (activeSection >= sections.length - 1 ? trailingStartType : null) ??
    sections[activeSection]?.startType ??
    'nextPage'

  pageSetupStateRef.current = {
    section,
    sections,
    trailingStartType,
    mirrorMargins,
    gutterAtTop,
    titlePg,
    evenOddHf,
    notes: documentNoteProps,
    sectionNotes: sectionNoteProps,
    autoHyphenation,
    hyphenation: hyphenationSettings,
  }
  useEffect(() => {
    const armed = pageSetupUndoRef.current
    pageSetupUndoRef.current = null
    if (!armed || !editor || !pageSetupStateRef.current) return
    const change = pageSetupChange(armed.before, pageSetupStateRef.current)
    if (change) recordPageSetupChange(editor, change, armed.edited)
  })
  const replayPageSetup = (patch: PageSetupPatch) => {
    const last = sections.length - 1
    if (patch.sections) {
      setSections((prev) => patchSections(prev, patch))
      // the last section and one behind an unsaved break are not written from the dirty list
      const rewritten = Object.keys(patch.sections)
        .map(Number)
        .filter((i) => i < last && !sections[i]?.pendingBreak)
      if (rewritten.length > 0) setSectionsDirty((d) => [...new Set([...d, ...rewritten])])
    }
    if (patch.section) {
      setSection((s) => s && { ...s, ...patch.section })
      setSectionDirty(true)
    }
    // back to the file's start type is still a value to write: a save may have come in between
    if ('trailingStartType' in patch)
      setTrailingStartType(patch.trailingStartType ?? patch.sections?.[last]?.startType ?? null)
    if (patch.mirrorMargins !== undefined) {
      setMirrorMargins(patch.mirrorMargins)
      setMirrorMarginsDirty(true)
    }
    if (patch.gutterAtTop !== undefined) setGutterAtTopEdit(patch.gutterAtTop)
    if (patch.titlePg !== undefined || patch.evenOddHf !== undefined) hfEditorRef.current?.exit()
    if (patch.titlePg !== undefined) {
      setTitlePg(patch.titlePg)
      setTitlePgDirty(true)
    }
    if (patch.evenOddHf !== undefined) {
      setEvenOddHf(patch.evenOddHf)
      setEvenOddHfDirty(true)
    }
    // the marks renumber once the sections and this settle (the numbering effect)
    const notes = patch.notes
    if (notes)
      setNotePropsEdit((edit) => ({
        footnote: { ...edit?.footnote, ...notes.footnote },
        endnote: { ...edit?.endnote, ...notes.endnote },
      }))
    if (patch.hyphenation || patch.autoHyphenation !== undefined) {
      const next = {
        ...hyphenationSettings,
        ...patch.hyphenation,
        auto: patch.autoHyphenation ?? autoHyphenation,
      }
      // it dispatches, and this runs inside the undo's own dispatch
      queueMicrotask(() => applyHyphenation(next, manualHyphenationRef.current))
    }
  }
  const replayPageSetupRef = useRef(replayPageSetup)
  replayPageSetupRef.current = replayPageSetup
  useEffect(() => {
    if (!editor) return
    const onTransaction = ({ transaction }: { transaction: Transaction }) => {
      const armed = pageSetupUndoRef.current
      if (armed && transaction.docChanged && transaction.getMeta('addToHistory') !== false)
        armed.edited = true
      for (const patch of replayedPageSetup(transaction)) replayPageSetupRef.current(patch)
      for (const { fonts, converting } of replayedFontConversions(transaction))
        setFontConversions((state) =>
          converting ? withFontsConverted(state, fonts) : withFontsReverted(state, fonts),
        )
    }
    editor.on('transaction', onTransaction)
    return () => {
      editor.off('transaction', onTransaction)
    }
  }, [editor])

  const ribbonActions = useStableCallbacks({
    onTableDialog: (kind: TableDialogKind) => {
      if (
        editor &&
        editableOrToast(editor) &&
        ensureCaretInTable(editor.state, editor.view.dispatch)
      )
        setTableDialog(kind)
    },
    onToggleTableGridlines: toggleTableGridlines,
    onOpen: () => void openFile(),
    onSave: () => void save(false),
    onSaveAs: () => void save(true),
    onNameWithAi: () => void nameWithAi(),
    onSection: (next: SectionSettings) => {
      armPageSetupUndo()
      // layout applies to the cursor's section; the final section's sectPr goes through SaveOptions.section (also drives canvas geometry)
      editSectionSettings((_s, i) => (i === activeSection ? next : null))
      if (sections.length <= 1 || activeSection === sections.length - 1) {
        setSection(next)
        setSectionDirty(true)
      } else {
        setSectionsDirty((d) => (d.includes(activeSection) ? d : [...d, activeSection]))
        setStatus(t('appSectionSettingsApplied', { n: activeSection + 1 }))
      }
    },
    onSectionsIn: (
      scope: 'whole' | 'forward',
      change: (s: SectionSettings) => SectionSettings,
      breakType: SectionInfo['startType'],
    ) => {
      armPageSetupUndo()
      if (scope === 'whole') {
        editSectionSettings((s) => change(s.settings))
        if (section) {
          setSection(change(section))
          setSectionDirty(true)
        }
        const earlier = sections.slice(0, -1).map((_, i) => i)
        if (earlier.length > 0) setSectionsDirty((d) => [...new Set([...d, ...earlier])])
        return
      }
      if (!canApplyForward) return
      // a later section behind an unsaved break saves from its break paragraph: written there too
      const patched = editor
        ? rewritePendingSectPrs(editor, sections, (s, i) =>
            i > activeSection ? applySectionSettings(s.sectPrXml, change(s.settings)) : null,
          )
        : sections
      // the break's own section takes the caret's index; everything after it changes
      const made = insertSectionBreak(breakType)
      const shift = made ? 1 : 0
      const from = activeSection + shift
      const count = sections.length + shift
      setSections((prev) =>
        prev.map((s, i) =>
          i < from
            ? s
            : {
                ...s,
                settings: change(s.settings),
                sectPrXml: (i > from && patched[i - shift]?.sectPrXml) || s.sectPrXml,
              },
        ),
      )
      const earlier = Array.from({ length: count - 1 - from }, (_, k) => from + k)
      if (earlier.length > 0) setSectionsDirty((d) => [...new Set([...d, ...earlier])])
      // the last section is written from its own state
      if (section) {
        setSection(change(section))
        setSectionDirty(true)
      }
    },
    onPageBorderAll: (props: SectionSettings['pageBorderProps'] | null) => {
      const bordered = ({ pageBorderProps: _drop, ...s }: SectionSettings): SectionSettings =>
        props ? { ...s, pageBorder: true, pageBorderProps: props } : { ...s, pageBorder: false }
      armPageSetupUndo()
      editSectionSettings((s) => bordered(s.settings))
      if (section) {
        setSection(bordered(section))
        setSectionDirty(true)
      }
      const earlier = sections.slice(0, -1).map((_, i) => i)
      if (earlier.length > 0) setSectionsDirty((d) => [...new Set([...d, ...earlier])])
    },
    onPaperSizeAll: (
      portraitW: number,
      portraitH: number,
      orientation: SectionSettings['orientation'],
    ) => {
      const sized = (s: SectionSettings): SectionSettings =>
        orientation === 'landscape'
          ? { ...s, orientation, pageWidth: portraitH, pageHeight: portraitW }
          : { ...s, orientation, pageWidth: portraitW, pageHeight: portraitH }
      armPageSetupUndo()
      editSectionSettings((s) => sized(s.settings))
      if (section) {
        setSection(sized(section))
        setSectionDirty(true)
      }
      const earlier = sections.slice(0, -1).map((_, i) => i)
      if (earlier.length > 0) setSectionsDirty((d) => [...new Set([...d, ...earlier])])
    },
    onMirrorMargins: (on: boolean) => {
      if (on === mirrorMargins) return
      armPageSetupUndo()
      setMirrorMargins(on)
      setMirrorMarginsDirty(true)
    },
    onGutterAtTop: (atTop: boolean) => {
      if (atTop === gutterAtTop) return
      armPageSetupUndo()
      setGutterAtTopEdit(atTop)
      // the caret's section is written by the command that carries this change
      editSectionSettings((s, i) => {
        const moved = gutterMoved(s.settings, atTop)
        return i === activeSection || moved === s.settings ? null : moved
      })
      if (section && sections.length > 1 && activeSection !== sections.length - 1)
        setSection(gutterMoved(section, atTop))
    },
    onSectionStart: (type: SectionInfo['startType']) => {
      const sec = sections[activeSection]
      if (sec?.pendingBreak || type === activeStartType) return
      armPageSetupUndo()
      if (sec)
        setSections((prev) =>
          prev.map((s, i) => (i === activeSection ? { ...s, startType: type } : s)),
        )
      if (sections.length <= 1 || activeSection === sections.length - 1) setTrailingStartType(type)
      else setSectionsDirty((d) => (d.includes(activeSection) ? d : [...d, activeSection]))
    },
    onInsertSectionBreak: (type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage') =>
      insertSectionBreak(type),
    onPageColor: (next: string | null) => {
      setPageColor(next)
      setPageColorDirty(true)
    },
    onWatermark: (next: Watermark | null) => {
      applyWatermark(next)
      const text = typeof next === 'string' ? next : next && 'text' in next ? next.text : null
      setStatus(
        next
          ? t('appWatermarkSet', { text: text ?? t('ribbonWatermarkPicture') })
          : t('appWatermarkRemoved'),
      )
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
    onInkPresets: setInkPresets,
    onInkActivate: setInkActiveId,
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
    onZoomDialog: () => setShowZoomDialog(true),
    onDarkPage: updateDarkPage,
    onHeader: (next: HeaderFooter) => commitHfAt(0, 'header', headerAreaView, next),
    onPageNumFormat: openPgNumModal,
    onFacingPageNumber: (kind: 'header' | 'footer', side: 'inside' | 'outside') => {
      const si = kind === 'header' ? 0 : lastSectionIdx
      // odd pages are the right-hand ones: their outer edge is the right
      const odd = side === 'outside' ? 'right' : 'left'
      const even = side === 'outside' ? 'left' : 'right'
      if (!evenOddHf) toggleEvenOddHf(true)
      commitHfAt(si, kind, 'default', hfPageNumberOnly(odd))
      commitHfAt(si, kind, 'even', hfPageNumberOnly(even))
    },
    onFooter: (next: HeaderFooter) => commitHfAt(lastSectionIdx, 'footer', footerAreaView, next),
    onHfAction: (action: HfAction) => onHfAction(action),
    onHfEdit: (kind: 'header' | 'footer') => openHfEditor(kind),
    onShowRuler: setShowRuler,
    onShowNav: setShowNav,
    onShowComments: () => toggleShowComments(),
    onToggleComments: () => toggleShowComments(),
    onNewComment: startNewComment,
    onDeleteComment: () => deleteCommentAtCaretImpl(reviewCtxRef.current),
    onDeleteAllComments: (resolvedOnly: boolean) =>
      deleteAllCommentsImpl(reviewCtxRef.current, resolvedOnly),
    onGotoComment: (dir: 1 | -1) => gotoCommentImpl(reviewCtxRef.current, dir),
    onTrackChanges: setTrackChangesByUser,
    onSpellcheck: (on: boolean) => {
      if (largeDocSpellOff) setLargeDocSpellOff(false)
      setSpellcheck(on)
    },
    onRevisionDisplay: (mode: RevisionDisplayMode) => setRevisionDisplay(mode),
    onAcceptRevision: (scope: RevisionScope) => handleRevision('accept', scope),
    onRejectRevision: (scope: RevisionScope) => handleRevision('reject', scope),
    onRevisionFilter: setRevisionFilter,
    onToggleReviewPane: () => setShowReviewPane((v) => !v),
    onGotoRevision: (dir: 1 | -1) => {
      if (!editor) return
      const body = filterRevisions(collectRevisions(editor.state.doc), revisionFilterRef.current)
      const stories = filterRevisions(storyRevisionsRef.current, revisionFilterRef.current)
      const ranges: RevisionRange[] = [...body, ...stories]
      if (ranges.length === 0) {
        setStatus(t('appReviewPaneEmpty'))
        return
      }
      // after a stop outside the body the walk continues from that entry
      const at = ranges.findIndex((r) => r.story && r.story.id === lastStoryRef.current)
      setActiveStoryId(null)
      let target: RevisionRange | undefined
      if (at >= 0) target = ranges[(at + dir + ranges.length) % ranges.length]
      else {
        const anchor = editor.state.selection.from
        target =
          dir === 1
            ? (body.find((r) => r.from > anchor) ?? stories[0] ?? body[0])
            : ([...body].reverse().find((r) => r.from < anchor) ?? ranges[ranges.length - 1])
      }
      if (target) showRevision(target)
    },
    onProtectDoc: () => setShowProtectDialog(true),
    onViewMode: switchView,
    onReadMode: setReadModeExclusive,
    onShowGrid: setShowGrid,
    onSplitView: setSplitView,
    onPagePreview: () => showPreview(true),
  })

  // the panel closes; Contextual cards stay unless the panel was the List view itself
  const closeCommentsPanel = useCallback(() => {
    setCommentPanelOpen(false)
    setCommentFocus(null)
    cancelNewComment()
    if (commentsMode === 'list') setShowComments(false)
  }, [cancelNewComment, commentsMode])
  const toggleShowComments = useCallback(() => {
    if (showComments) {
      setCommentPanelOpen(false)
      setCommentFocus(null)
      setCommentPopup(null)
      cancelNewComment()
    }
    setShowComments(!showComments)
  }, [showComments, cancelNewComment])

  // the popped-up thread closes on Escape or a click outside it and its icon
  useEffect(() => {
    if (!commentPopup) return
    const onDown = (e: PointerEvent) => {
      const el = e.target as Element | null
      if (!el?.closest('.comment-bubble.comment-popup, .comment-balloon')) setCommentPopup(null)
    }
    const onKey = (e: globalThis.KeyboardEvent) => {
      if (e.key === 'Escape') setCommentPopup(null)
    }
    document.addEventListener('pointerdown', onDown, true)
    document.addEventListener('keydown', onKey, true)
    return () => {
      document.removeEventListener('pointerdown', onDown, true)
      document.removeEventListener('keydown', onKey, true)
    }
  }, [commentPopup])
  useEffect(() => {
    setCommentPopup(null)
  }, [commentsUi.cards, commentsUi.panel])

  // selection / popup changes only restyle the margin layer: redraw it without a pagination pass
  useEffect(() => {
    if (!editor || !doc) return
    const pm = document.querySelector('.editor-scroll .ProseMirror') as HTMLElement | null
    const wrap = pm?.closest('.page-wrap') as HTMLElement | null
    if (!pm || !wrap) return
    syncMarginAnnotations(
      wrap,
      pm,
      comments,
      zoom / 100,
      doc.parsed.blocks,
      editor.view,
      marginHandlersRef.current,
    )
    // only the flags on the root matter here; the pagination pass owns every other input
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [commentFocus?.id, commentPopup])

  // Simple Markup overlay callbacks (the overlay is rebuilt per pagination
  // frame; a stable object keeps the DOM listeners on the current editor)
  const marginHandlersRef = useRef<MarginHandlers>({})
  marginHandlersRef.current = {
    onChangeBar: (pos) => {
      setRevisionDisplay('all')
      if (!editor) return
      // the bar sits on the change's own line, so it is already in view: just
      // select it; row / cell bars carry a node position, so snap to the nearest text
      const { doc, tr } = editor.state
      const $pos = doc.resolve(Math.min(pos, doc.content.size))
      editor.view.dispatch(tr.setSelection(TextSelection.near($pos, 1)).setMeta(TRACK_IGNORE, true))
    },
    onCommentBalloon: (id) => {
      if (iconClickTarget(commentsUi) === 'focus') setCommentFocus({ id, nonce: Date.now() })
      else setCommentPopup(id)
    },
    onCommentSelect: (id) => setCommentFocus({ id, nonce: Date.now() }),
    onCommentReply: (id, text) => replyToComment(id, text),
    onCommentResolve: resolveComment,
    onCommentEdit: editComment,
    onCommentDelete: deleteComment,
  }

  const hasDoc = !!doc
  // a balloon names the change starting at its anchor; a caret may sit on a boundary shared by two
  const revisionsAtCtx = (all: RevisionRange[], pos: number, balloon: boolean) => {
    const hits = revisionsInSelection(all, pos, pos)
    if (!balloon) return hits
    const own = hits.filter((r) => r.from <= pos && pos < r.to)
    return own.length ? own : hits
  }
  const ctxCommentId = useMemo(() => {
    if (!ctxMenu || !editor) return null
    const ids = ctxMenu.commentId
      ? [ctxMenu.commentId]
      : commentIdsAt(editor.state, ctxMenu.pos ?? editor.state.selection.from)
    const open = ids
      .map((id) => comments.find((c) => c.id === id))
      .filter((c): c is CommentInfo => !!c && !c.done)
    const root = open.find((c) => !c.parentId) ?? open[0]
    return root ? (root.parentId ?? root.id) : null
  }, [ctxMenu, editor, comments])
  // the context menu acts on the document as it is now, not the throttled pane snapshot
  const liveRevisions = useMemo(
    () => (ctxMenu && editor ? collectRevisions(editor.state.doc) : []),
    [ctxMenu, editor],
  )
  // Undo/redo availability: refreshed on every transaction so the QAT buttons grey out when empty
  const [histState, setHistState] = useState({ canUndo: false, canRedo: false })
  useEffect(() => {
    if (!editor) return
    const refresh = ({ transaction }: { transaction?: { docChanged: boolean } } = {}) => {
      setHistState({ canUndo: editor.can().undo(), canRedo: editor.can().redo() })
      if (transaction?.docChanged) setDocVersion((v) => v + 1)
    }
    refresh()
    editor.on('transaction', refresh)
    // comment metadata rides PM undo/redo: deleted comments come back whole,
    // undo of an add removes them again (Word parity)
    const detachCommentSync = attachCommentUndoSync(
      editor,
      () => commentsLiveRef.current,
      (entries) => {
        setComments(entries)
        setCommentsDirty(true)
      },
    )
    return () => {
      detachCommentSync()
      editor.off('transaction', refresh)
    }
  }, [editor])

  // the native menu's check marks and greyed-out items follow the document, as in Word
  const editTarget = useEditTarget(editor)
  const menuHasDoc = Boolean(doc)
  const {
    editable: menuEditable,
    inTable: menuInTable,
    canMergeCells: menuCanMerge,
    labelSheet: menuLabelSheet,
    hasSelection: menuHasSelection,
    imageSelected: menuPictureSelected,
    hasNotes: menuHasNotes,
    hasChart: menuHasChart,
    canSplitTable: menuCanSplitTable,
    canComment: menuCanComment,
    inEmptyParagraph: menuInEmptyParagraph,
    hasTableOfFigures: menuHasTableOfFigures,
    hasIndex: menuHasIndex,
    hasTableOfAuthorities: menuHasTableOfAuthorities,
  } = formatState
  const menuHasComments = comments.length > 0
  const menuHasMergeSource = Boolean(mergeSource)
  const menuInTextBox = Boolean(formatState.sub)
  const menuCanRepeat = canRepeatCommand({ editor })
  useEffect(() => {
    window.desktop.reportCommandState?.(
      menuCommandState({
        aiSidebar: showAi,
        darkPage,
        hasDoc: menuHasDoc,
        editable: menuEditable,
        inTable: menuInTable,
        inTextBox: menuInTextBox,
        spellcheck: spellcheckActive,
        readOnlyRecommended: Boolean(writeProtection?.recommended),
        onDisk: Boolean(doc?.filePath),
        view: {
          mode: viewMode,
          readMode,
          ruler: showRuler,
          grid: showGrid,
          navigation: showNav,
          thumbnails: showThumbnails,
          markup: markupShown(revisionDisplay),
          ribbon: !ribbonCollapsed,
          focus: focusMode,
          reviewPane: showReviewPane,
        },
        cellsSelected: menuCanMerge,
        labelSheet: menuLabelSheet,
        hasSelection: menuHasSelection,
        pictureSelected: menuPictureSelected,
        hasNotes: menuHasNotes,
        hasChart: menuHasChart,
        canSplitTable: menuCanSplitTable,
        canComment: menuCanComment,
        inEmptyParagraph: menuInEmptyParagraph,
        hasComments: menuHasComments,
        hasTableOfFigures: menuHasTableOfFigures,
        hasIndex: menuHasIndex,
        hasTableOfAuthorities: menuHasTableOfAuthorities,
        hasMergeSource: menuHasMergeSource,
        windowsDetached,
        canUndo: histState.canUndo,
        canRedo: histState.canRedo,
        canRepeat: menuCanRepeat,
        target: editTarget,
      }),
    )
  }, [
    showAi,
    darkPage,
    menuHasDoc,
    menuEditable,
    menuInTable,
    menuInTextBox,
    spellcheckActive,
    viewMode,
    readMode,
    showRuler,
    showGrid,
    showNav,
    showThumbnails,
    revisionDisplay,
    ribbonCollapsed,
    focusMode,
    doc?.filePath,
    writeProtection?.recommended,
    showReviewPane,
    menuCanMerge,
    menuLabelSheet,
    menuHasSelection,
    menuPictureSelected,
    menuHasNotes,
    menuHasChart,
    menuCanSplitTable,
    menuCanComment,
    menuInEmptyParagraph,
    menuHasComments,
    menuHasTableOfFigures,
    menuHasIndex,
    menuHasTableOfAuthorities,
    menuHasMergeSource,
    windowsDetached,
    histState.canUndo,
    histState.canRedo,
    menuCanRepeat,
    editTarget,
  ])
  const quickActions = useMemo(
    () => (
      <>
        <button
          className="qa-btn"
          data-tip={t('appSaveShortcutTip')}
          aria-label={t('appSaveShortcutTip')}
          disabled={!hasDoc || !hasUnsavedChanges}
          onClick={() => void save(false)}
        >
          <IconSave size={16} />
        </button>
        <button
          className="qa-btn"
          data-tip={t('appUndo')}
          aria-label={t('appUndo')}
          disabled={!hasDoc || !histState.canUndo}
          onClick={() => editor?.chain().focus().undo().run()}
        >
          <IconUndo size={16} />
        </button>
        <button
          className="qa-btn"
          data-tip={t('appRedo')}
          aria-label={t('appRedo')}
          disabled={!hasDoc || !histState.canRedo}
          onClick={() => editor?.chain().focus().redo().run()}
        >
          <IconRedo size={16} />
        </button>
        <label className={`autosave-toggle ${autoSave ? 'on' : ''}`} data-tip={t('appAutoSaveTip')}>
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
    [hasDoc, hasUnsavedChanges, autoSave, editor, save, lang, histState],
  )

  // ---- Multiple Pages: the grid is a function of zoom × pane width ----
  const [paneW, setPaneW] = useState(0)
  useEffect(() => {
    const el = scrollContainerRef.current
    if (!doc || !el) return
    const ro = new ResizeObserver(() => setPaneW(el.clientWidth))
    ro.observe(el)
    setPaneW(el.clientWidth)
    return () => ro.disconnect()
  }, [doc])
  const mpPaperW = canvasBox ? paperWidthPx(sections, canvasSection) : 0
  const mpSlotW = mpPaperW + markupExtra
  // differing paper widths keep the single column (their sheets are drawn per page)
  const mpUniform = !sections.some((s) => pageLeftPx(s.settings, mpPaperW) > 0.5)
  const mpCols =
    doc && viewMode === 'print' && !readMode && mpUniform && multipageMode
      ? multipageColumns(paneW, mpSlotW, zoom)
      : 1
  const mpWanted = mpCols >= 2
  const [mpFrames, setMpFrames] = useState<PageFrame[]>([])
  const canvasPageH = canvasBox?.height ?? 0
  useEffect(() => {
    if (!mpWanted) {
      setMpFrames((prev) => (prev.length ? [] : prev))
      return
    }
    const wrap = editor?.view.dom.closest<HTMLElement>('.page-wrap')
    if (!wrap) return
    const frames = pageFramesFromGaps(wrap, zoom / 100, undefined, pageLayoutRef.current?.cuts)
    // the canvas ends where the content ends; the grid shows the last sheet whole
    const last = frames[frames.length - 1]
    if (last && last.bottom - last.top < canvasPageH) last.bottom = last.top + canvasPageH
    setMpFrames(frames)
  }, [doc, editor, mpWanted, zoom, layoutRev, canvasPageH])
  const multipage = mpWanted && mpFrames.length >= 2
  const mpLayout = useMemo(
    () =>
      multipage
        ? // every slot is a full sheet: a mid-paragraph cut leaves its canvas frame a
          // few px short of the paper, and Word's grid is equal-height
          layoutMultipage(
            mpFrames.map((f) => Math.max(canvasPageH, f.bottom - f.top)),
            mpCols,
            mpSlotW,
          )
        : null,
    [multipage, mpFrames, mpCols, mpSlotW, canvasPageH],
  )
  const [mpActive, setMpActive] = useState(0)
  const mpActiveIdx = Math.max(0, Math.min(mpActive, mpFrames.length - 1))
  multipageActiveRef.current = multipage ? mpActiveIdx : null
  mpFramesRef.current = mpFrames

  const [mpSnap, setMpSnap] = useState<CanvasSnapshot | null>(null)
  useEffect(() => {
    if (!multipage) {
      setMpSnap(null)
      return
    }
    const wrap = editor?.view.dom.closest<HTMLElement>('.page-wrap')
    if (!wrap) return
    const raf = requestAnimationFrame(() => setMpSnap(snapshotCanvas(wrap, zoom / 100)))
    return () => cancelAnimationFrame(raf)
  }, [multipage, editor, zoom, layoutRev, mpFrames, mpActiveIdx])

  /** the page under the caret, so the live window follows selection moves and page turns */
  const mpFollowCaret = useCallback(() => {
    if (!editor || multipageActiveRef.current == null) return
    const frames = mpFramesRef.current
    const wrap = editor.view.dom.closest<HTMLElement>('.page-wrap')
    if (!wrap || frames.length === 0) return
    let y: number
    try {
      y = editor.view.coordsAtPos(editor.state.selection.head).top
    } catch {
      return
    }
    const canvasY = (y - wrap.getBoundingClientRect().top) / (zoomLiveRef.current / 100)
    let idx = 0
    for (let i = 1; i < frames.length; i++) if (frames[i].top <= canvasY + 0.5) idx = i
    setMpActive(idx)
  }, [editor])
  useEffect(() => {
    if (!editor || !multipage) return
    const onTr = ({ transaction }: { transaction: Transaction }) => {
      // a typed change settles on its page after the next pagination pass
      if (!transaction.docChanged && transaction.selectionSet) mpFollowCaret()
    }
    editor.on('transaction', onTr)
    return () => {
      editor.off('transaction', onTr)
    }
  }, [editor, multipage, mpFollowCaret])
  useEffect(() => {
    if (multipage) mpFollowCaret()
  }, [multipage, mpFrames, mpFollowCaret])

  /** a click on a painted page moves the live editor there and lands the caret under the pointer */
  const mpPickRef = useRef<{ x: number; y: number } | null>(null)
  const onMpPick = useCallback((index: number, x: number, y: number) => {
    mpPickRef.current = { x, y }
    setMpActive(index)
  }, [])
  useLayoutEffect(() => {
    const pick = mpPickRef.current
    if (!pick || !editor) return
    mpPickRef.current = null
    const view = editor.view
    const wrap = view.dom.closest<HTMLElement>('.page-wrap')
    const frames = mpFramesRef.current
    const frame = frames[mpActiveIdx]
    if (!wrap || !frame) return
    const z = zoomLiveRef.current / 100
    const wrapTop = wrap.getBoundingClientRect().top
    const pageOf = (pos: number): number => {
      const y = (view.coordsAtPos(pos).top - wrapTop) / z
      let idx = 0
      for (let i = 1; i < frames.length; i++) if (frames[i].top <= y + 0.5) idx = i
      return idx
    }
    let pos = view.posAtCoords({ left: pick.x, top: pick.y })?.pos ?? null
    if (pos != null && pageOf(pos) !== mpActiveIdx) pos = null
    if (pos == null) {
      // the pointer is on blank paper whose nearest caret slot is another
      // page's: take the picked page's block nearest to the pointer instead
      const canvasY = (pick.y - wrapTop) / z
      let best: { el: HTMLElement; d: number; r: DOMRect } | null = null
      for (const el of Array.from(view.dom.children) as HTMLElement[]) {
        if (el.classList.contains('ProseMirror-widget')) continue
        const r = el.getBoundingClientRect()
        const top = (r.top - wrapTop) / z
        const bottom = (r.bottom - wrapTop) / z
        if (bottom <= frame.top || top >= frame.bottom) continue
        const d = canvasY < top ? top - canvasY : canvasY > bottom ? canvasY - bottom : 0
        if (!best || d < best.d) best = { el, d, r }
      }
      if (best) {
        const x = Math.min(Math.max(pick.x, best.r.left + 1), best.r.right - 1)
        const y = Math.min(Math.max(pick.y, best.r.top + 1), best.r.bottom - 1)
        pos = view.posAtCoords({ left: x, top: y })?.pos ?? null
      }
    }
    if (pos != null) {
      const sel = TextSelection.near(view.state.doc.resolve(pos), 1)
      view.dispatch(view.state.tr.setSelection(sel))
    }
    view.focus()
  }, [mpActiveIdx, editor])

  /** scroller offset of the active slot's top edge (CSS px) */
  const mpActiveSlotTop = useCallback((): number | null => {
    const scroller = scrollContainerRef.current
    const grid = scroller?.querySelector<HTMLElement>('.doc-zoom')
    const slot = mpLayout?.slots[mpActiveIdx]
    if (!scroller || !grid || !slot) return null
    const gridTop =
      grid.getBoundingClientRect().top - scroller.getBoundingClientRect().top + scroller.scrollTop
    return gridTop + slot.top * (zoom / 100)
  }, [mpLayout, mpActiveIdx, zoom])
  // a caret-driven page turn may land on a slot outside the viewport (zoom
  // changes keep their own anchoring below)
  const mpSeenActiveRef = useRef<number | null>(null)
  useLayoutEffect(() => {
    const scroller = scrollContainerRef.current
    const seen = mpSeenActiveRef.current
    mpSeenActiveRef.current = multipage ? mpActiveIdx : null
    if (!multipage || !scroller || !mpLayout || seen === mpActiveIdx) return
    const top = mpActiveSlotTop()
    if (top == null) return
    const slot = mpLayout.slots[mpActiveIdx]
    const bottom = top + slot.height * (zoom / 100)
    const viewTop = scroller.scrollTop
    const viewBottom = viewTop + scroller.clientHeight
    if (bottom > viewTop && top < viewBottom) return
    scroller.scrollTop = Math.max(0, top - 26)
    scrollPosRef.current = { left: scroller.scrollLeft, top: scroller.scrollTop }
  }, [multipage, mpActiveIdx, mpLayout, mpActiveSlotTop, zoom])

  const multipagePrevRef = useRef({ on: false, cols: 1 })
  useLayoutEffect(() => {
    const prevZoom = prevZoomRef.current
    prevZoomRef.current = zoom
    const anchor = zoomAnchorRef.current
    zoomAnchorRef.current = null
    const container = scrollContainerRef.current
    const page = container?.querySelector<HTMLElement>('.doc-zoom')
    if (!container || !page) return
    // the Multiple Pages grid re-flows when its column count changes: keep
    // the active page in view instead of a point that no longer exists
    const prevMp = multipagePrevRef.current
    const colsNow = multipage ? mpCols : 1
    multipagePrevRef.current = { on: multipage, cols: colsNow }
    if (multipage !== prevMp.on || colsNow !== prevMp.cols) {
      let top = mpActiveSlotTop()
      if (top == null && editor) {
        // grid collapsed: the caret's page in the single column
        try {
          const c = editor.view.coordsAtPos(editor.state.selection.head)
          top = c.top - container.getBoundingClientRect().top + container.scrollTop - 120
        } catch {
          top = null
        }
      }
      if (top != null) {
        container.scrollTop = Math.max(0, top - 26)
        scrollPosRef.current = { left: container.scrollLeft, top: container.scrollTop }
      }
      return
    }
    if (prevZoom === zoom) return

    const ratio = zoom / prevZoom
    const vx = anchor ? anchor.vx : container.clientWidth / 2
    const vy = anchor ? anchor.vy : container.clientHeight / 2
    const cs = getComputedStyle(container)
    const padLeft = parseFloat(cs.paddingLeft) || 0
    const innerWidth = container.clientWidth - padLeft - (parseFloat(cs.paddingRight) || 0)
    const containerRect = container.getBoundingClientRect()
    const pageRect = page.getBoundingClientRect()
    const pageLeft = pageRect.left - containerRect.left + container.scrollLeft
    const pageTop = pageRect.top - containerRect.top + container.scrollTop
    const prevPageLeft = padLeft + Math.max(0, (innerWidth - pageRect.width / ratio) / 2)
    const prev = scrollPosRef.current

    const docX = Math.max(0, prev.left + vx - prevPageLeft)
    const docY = Math.max(0, prev.top + vy - pageTop)
    container.scrollLeft = pageLeft + docX * ratio - vx
    container.scrollTop = pageTop + docY * ratio - vy
    scrollPosRef.current = { left: container.scrollLeft, top: container.scrollTop }
  }, [zoom, multipage, mpCols, mpActiveSlotTop, editor])

  useEffect(() => {
    if (!editor) return
    editor.view.dispatch(
      editor.state.tr
        .setMeta(formattingMarksPluginKey, {
          // Draft shows every break as a labelled rule, marks on or off
          on: showMarks || viewMode === 'draft',
          breaks: sectionBreakMap(sections),
        })
        .setMeta('addToHistory', false),
    )
  }, [editor, showMarks, viewMode, sections, lang])

  if (!editor) return null

  // canvas geometry is anchored to the first section (stable across cursor moves);
  // sections with a different content width carry per-block width decorations
  // the shared paper must cover the widest section or its content lays out past
  // the paper edge onto the editor background (a white band down the right side);
  // narrower pages are centered on it (pageLeftPx) and paint their own sheets
  const paperW = canvasBox ? paperWidthPx(sections, canvasSection) : 0
  const pageLeftOf = (set: SectionSettings) => pageLeftPx(set, paperW)
  const mixedPaper =
    viewMode === 'print' && !readMode && sections.some((s) => pageLeftOf(s.settings) > 0.5)
  // the trailing footer strip belongs to the LAST page, which is always in the
  // last section; on differing-width docs the stylesheet centering (50% of the
  // first-section paper) puts it at the wrong x, so pin it to its own section
  const lastSection = sections[sections.length - 1]?.settings ?? section
  const lastBox = lastSection ? sectionPageBox(lastSection) : null
  // pin whenever the shared paper is wider than the footer's own section (a wider
  // middle section widens the paper too) or the strip width differs from the
  // first-section default
  // mirrored documents pin the first/last page's strips to that page's own (parity)
  // margin: inside/outside differ, so the stylesheet's centering would miss the body
  const firstPageMargins = canvasSection
    ? pageMargins(canvasSection, sections[0]?.pageNumberStart ?? 1, mirrorMargins)
    : null
  const lastPageMargins = lastSection
    ? pageMargins(lastSection, lastPageNo?.num ?? 1, mirrorMargins)
    : null
  const edgeFooterStyle =
    lastBox &&
    canvasBox &&
    lastPageMargins &&
    (paperW - lastBox.width > 0.5 ||
      Math.abs(lastBox.contentWidth - canvasBox.contentWidth) > 0.5 ||
      mirrorMargins)
      ? {
          width: `${lastBox.contentWidth}px`,
          left: `${pageLeftOf(lastSection!) + twipsToPx(lastPageMargins.left)}px`,
          transform: 'none',
          bottom: `${lastBox.footerDist}px`,
        }
      : undefined
  // symmetric pin for the leading header (it belongs to the first section): the
  // widened shared paper would otherwise re-center it at the wrong x
  const edgeHeaderStyle =
    canvasBox &&
    canvasSection &&
    firstPageMargins &&
    (paperW - canvasBox.width > 0.5 || mirrorMargins)
      ? {
          width: `${canvasBox.contentWidth}px`,
          left: `${pageLeftOf(canvasSection) + twipsToPx(firstPageMargins.left)}px`,
          transform: 'none',
        }
      : undefined
  const rulerBase = sections[activeSection]?.settings ?? section
  const rulerSwap = mirrorMargins && caretPageNo % 2 === 0
  const rulerSection =
    rulerBase && rulerSwap
      ? { ...rulerBase, marginLeft: rulerBase.marginRight, marginRight: rulerBase.marginLeft }
      : rulerBase
  /** View ▸ Ruler: fixed bars around the document view (none in Multiple Pages) */
  const rulersShown = Boolean(doc && showRuler && rulerSection && !multipage)
  const docZoomStyle = {
    zoom: zoom / 100,
    '--page-w': canvasBox ? `${paperW}px` : undefined,
    '--page-h': canvasBox ? `${canvasBox.height}px` : undefined,
    '--section-content-w': canvasBox ? `${canvasBox.contentWidth}px` : undefined,
    '--page-pad': canvasSection
      ? `${canvasTop}px ${twipsToPx(canvasSection.marginRight)}px ${canvasBottom}px ${twipsToPx(canvasSection.marginLeft)}px`
      : undefined,
    '--page-mt': canvasSection ? `${canvasTop}px` : undefined,
    '--page-mb': canvasSection ? `${canvasBottom}px` : undefined,
    '--page-ml': canvasSection ? `${twipsToPx(canvasSection.marginLeft)}px` : undefined,
    '--page-mr': canvasSection ? `${twipsToPx(canvasSection.marginRight)}px` : undefined,
    '--header-dist': canvasBox ? `${canvasBox.headerDist}px` : undefined,
    '--footer-dist': canvasBox ? `${canvasBox.footerDist}px` : undefined,
    '--page-bg': pageColor ? `#${pageColor}` : undefined,
    '--page-cols': colFlow && viewMode === 'print' ? colFlow.cols : undefined,
    // end-of-document footnote area sits on the last page's sheet
    '--last-page-w': mixedPaper && lastBox ? `${lastBox.width}px` : undefined,
    '--last-page-x': mixedPaper && lastSection ? `${pageLeftOf(lastSection)}px` : undefined,
  } as CSSProperties

  // page-dark: dark paper + remapped colors (styles.css @media screen block); a
  // document with its own page color keeps it, Word-style. workspace-dark: the
  // gray canvas band around the dark paper when the chrome itself is light.
  const workspaceClass = [
    'workspace',
    darkPage && !pageColor ? 'page-dark' : '',
    darkPage && !themeDark ? 'workspace-dark' : '',
  ]
    .filter(Boolean)
    .join(' ')

  const docZoomClass = [
    'doc-zoom',
    multipage ? 'multipage' : '',
    showMarks ? 'show-marks' : '',
    `view-${viewMode}`,
    viewMode === 'story' && storyFocus ? 'story-focus-on' : '',
    showGrid ? 'show-grid' : '',
    tableGridlines ? 'show-table-gridlines' : '',
  ]
    .filter(Boolean)
    .join(' ')

  return (
    <div
      className={`app ${readMode ? 'read-mode' : ''}${viewMode === 'story' ? ' story-chrome' : ''}${focusMode ? ' focus-mode' : ''}${commentsUi.panel ? ' comments-pane-open' : ''}${revisionDisplay !== 'all' ? ` rev-display-${revisionDisplay}` : ''}${balloonsActive && balloonMode === 'revisions' ? ' rev-balloon' : ''}${revisionFilter.insDel ? '' : ' rev-hide-insdel'}${revisionFilter.format ? '' : ' rev-hide-format'}${tcOptions.balloonLines ? '' : ' markup-no-lines'}`}
      data-balloons={balloonsActive ? balloonMode : undefined}
      data-comment-cards={commentsUi.cards ? '' : undefined}
      data-comment-focus={commentFocus?.id}
      data-comment-popup={commentPopup ?? undefined}
    >
      <ToastHost />
      {docCss && <style data-doc-css="">{docCss}</style>}
      {doc && <style data-doc-css="">{mathDefaultsCss(mathPr)}</style>}
      {revisionAuthorStyle && <style data-rev-authors="">{revisionAuthorStyle}</style>}
      {storyAuthorStyle && <style data-story-authors="">{storyAuthorStyle}</style>}
      {tcOptionsStyle && <style data-rev-options="">{tcOptionsStyle}</style>}
      {liveLineFactor != null && (
        <style data-doc-css="">{`.doc-page { --doc-line-factor:${liveLineFactor} }`}</style>
      )}
      {doc && (gridPitchPt ?? mixedGridPitchPt) != null && (
        // typed w:docGrid: line-height round(up) expressions snap to this pitch
        <style>{`.doc-page { --doc-grid-pitch:${gridPitchPt ?? mixedGridPitchPt}pt }`}</style>
      )}
      {doc && charSpacePt != null && (
        // w:docGrid charSpace: every character advances natural width + this delta
        <style>{`.doc-page { --doc-char-space:${Math.round(charSpacePt * 10000) / 10000}pt${
          balancedGrid
            ? `; --doc-char-space-sb:${Math.round(singleByteCharSpacePt(charSpacePt, true) * 10000) / 10000}pt`
            : ''
        } }`}</style>
      )}
      {doc && charCellPt != null && (
        <style>{`.doc-page { --doc-char-snap:${charCellSpacingCss(charCellPt)} }`}</style>
      )}
      {doc && snapsToChars && (
        // snapToChars: a full-width character takes whole cells and sits centred in them
        <style>{`.doc-page .doc-east-asian-font:not(.doc-east-asian-halfwidth, [style*='letter-spacing'] .doc-east-asian-font, .page-gap-hf *) { letter-spacing: var(--doc-char-snap, var(--doc-char-space, normal)); position: relative; inset-inline-start: calc(var(--doc-char-snap, 0px) / 2) }`}</style>
      )}
      {doc && balancedGrid && !snapsToChars && (
        // full-width spans take the whole delta back (halfwidth kana stay single-byte);
        // an inline run letter-spacing still wins
        <style>{`.doc-page .doc-east-asian-font:not(.doc-east-asian-halfwidth, [style*='letter-spacing'] .doc-east-asian-font, .page-gap-hf *) { letter-spacing: var(--doc-char-space, normal) }`}</style>
      )}
      {doc && section && (
        // over-wide tables may spill into the margins (Word/LO), capped at the paper edge
        <style>{`.doc-page { --doc-margin-left:${twipsToPx(section.marginLeft)}px; --doc-margin-right:${twipsToPx(section.marginRight)}px; --doc-margin-top:${twipsToPx((canvasSection ?? section).marginTop)}px }`}</style>
      )}
      {/* Theme CSS comes from live state, so a Design ▸ Themes/Fonts/Colors pick shows
          on the page immediately instead of only in the saved file */}
      {doc && (
        <style data-doc-css="">
          {docThemeCss(themeFonts, themeColors, !!docBodyFont(doc.parsed))}
        </style>
      )}
      {colFlow && viewMode === 'print' && (
        // columns (sectPr w:cols): column gap follows the document's w:space; measuring-columns
        // is the single-flow measuring state (columns removed, content-box width = column width,
        // toggled instantaneously for measurement, invisible).
        // .doc-page is border-box, so the measured width must add back the left/right margin padding
        <style>{`.editor-scroll .doc-page { column-count: ${colFlow.cols}; column-gap: ${colFlow.gapPx}px; column-fill: balance; }
.editor-scroll .doc-page.measuring-columns { column-count: auto; width: ${colFlow.colWidthPx + twipsToPx(canvasSection?.marginLeft ?? section?.marginLeft ?? 0) + twipsToPx(canvasSection?.marginRight ?? section?.marginRight ?? 0)}px; }`}</style>
      )}
      <Ribbon
        actionsRef={ribbonActionsRef}
        onRibbonCollapsed={setRibbonCollapsed}
        readOnlyRecommended={Boolean(writeProtection?.recommended)}
        mergePreview={mergePreview}
        windowsDetached={windowsDetached}
        quickActions={quickActions}
        mergeSource={mergeRecipients}
        mergeRecordIndex={mergeRecordIndex}
        onMergeRecordIndex={setMergeRecordIndex}
        mergeRange={mergeRange}
        onMergeRange={setMergeRange}
        editor={editor}
        formatState={formatState}
        hasDoc={!!doc}
        onCommand={runCommand}
        editTarget={editTarget}
        styles={ribbonStyles}
        docDefaults={doc?.parsed.docDefaults}
        documentListPresets={documentLists}
        recentListPresets={recentLists}
        showAi={showAi}
        section={sections[activeSection]?.settings ?? section}
        tableGridlines={tableGridlines}
        activeSection={sections.length > 1 ? activeSection : null}
        mirrorMargins={mirrorMargins}
        pageColor={pageColor}
        watermark={watermark}
        watermarkPicture={watermarkPicture}
        watermarkDocImage={
          watermarkDirty
            ? null
            : (doc?.parsed.headerImages?.find((img) => img.wordArt || img.watermark) ?? null)
        }
        themeFonts={themeFonts}
        themeColors={themeColors}
        inkTool={inkTool}
        inkPresets={inkPresets}
        inkActiveId={inkActiveId}
        inkCount={inkAnnotations.length}
        sources={sources}
        citationStyle={citationStyle}
        zoteroNoteFields={zoteroNoteFields}
        showZotero={zoteroInstalled || zoteroDocumentData !== '' || zoteroNoteFields}
        zoom={Math.round(zoom)}
        darkPage={darkPage}
        tabRequest={ribbonTabRequest}
        header={shownHeader}
        footer={shownFooter}
        hfEditing={hfEditingInfo}
        showMarks={showMarks}
        showRuler={showRuler}
        showNav={showNav}
        showStylesPane={showStylesPane}
        commentCount={comments.length}
        showComments={showComments}
        commentsMode={commentsMode}
        onCommentsMode={setCommentsMode}
        openCommentCount={comments.filter((c) => !c.parentId && c.done !== true).length}
        resolvedCommentCount={comments.filter((c) => !c.parentId && c.done === true).length}
        commentAtCaret={commentThreadAtCaretImpl(reviewCtxRef.current) !== null}
        trackChanges={trackChanges}
        spellcheck={spellcheckActive}
        revisionDisplay={revisionDisplay}
        balloonMode={balloonMode}
        revisionCount={revisionCount}
        revisionFilter={revisionFilter}
        revisionAuthors={reviewAuthors}
        showReviewPane={showReviewPane}
        isProtected={isProtected}
        commentsAllowed={commentsAllowed}
        trackChangesForced={trackChangesForced}
        protectActive={
          isProtected || trackChangesForced || (doc?.encrypted ?? false) || !!writeProtection?.hash
        }
        viewMode={viewMode}
        readMode={readMode}
        showGrid={showGrid}
        splitView={splitView}
        {...ribbonActions}
      />

      <div className="app-main">
        {doc && (
          <div
            className={`ai-dock${showAi ? '' : ' collapsed'}${
              !aiSidebarEnabled && !showAi ? ' ai-dock-off' : ''
            }`}
          >
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
              hideRail={!aiSidebarEnabled}
              onExpand={() => setShowAi(true)}
              onCollapse={() => setShowAi(false)}
              filePath={doc?.filePath ?? null}
              editQueue={editQueue}
              onQueueEditInstruction={queueUpdate}
              onQueueRemove={queueRemove}
              onQueueClear={queueClear}
              onQueueFocus={queueFocus}
              onQueueConsume={queueConsume}
              commentsAccess={aiCommentsAccess}
              hfAccess={aiHfAccess}
              pageSetupAccess={aiPageSetupAccess}
              docExtras={aiDocExtras}
              notesAccess={aiNotesAccess}
              trackChanges={trackChanges}
              trackChangesLocked={isProtected || trackChangesForced}
              onTrackChangesToggle={() => setTrackChanges((v) => !v)}
              revisionCount={revisionCount}
              onRevisions={(action) => {
                if (!isProtected && !trackChangesForced) handleRevision(action, 'all')
              }}
              reviewEntries={allReviewEntries}
              activeReviewIndex={activeAllReviewIndex}
              onRevisionSelect={showRevision}
              onRevisionRange={(action, rev) => {
                if (!isProtected && !trackChangesForced) handleRevisionRange(action, rev)
              }}
            />
          </div>
        )}
        <div className="app-content">
          <div className={workspaceClass}>
            {doc && showFind && (
              <FindPanel
                editor={editor}
                onClose={() => {
                  setShowFind(false)
                  // else the next plain ⌘F remount would still see a truthy nonce
                  // and land focus on the replace / go-to field
                  setFindFocusReplace(0)
                  setFindFocusGoTo(0)
                  setFindStep(null)
                }}
                focusFindNonce={findFocusInput}
                focusReplaceNonce={findFocusReplace}
                focusGoToNonce={findFocusGoTo}
                stepRequest={findStep}
                layout={() => pageLayoutRef.current}
              />
            )}
            {doc && showNav && (
              <NavPane
                editor={editor}
                doc={editor.state.doc}
                zoom={zoom / 100}
                pageInfo={pageInfo}
                onClose={closeNav}
                cuts={cutsNow}
                tabRequest={navTabRequest}
                onTabRequestDone={clearNavTabRequest}
                onTabChange={setNavTab}
              />
            )}
            {doc && (
              <AiAskPopover
                editor={editor}
                queueFull={editQueue.length >= EDIT_QUEUE_MAX}
                getItem={getQueueItem}
                onSendNow={askSendNow}
                onQueueAdd={queueAdd}
                onQueueUpdate={queueUpdate}
                onQueueRemove={queueRemove}
              />
            )}
            {doc && <PasteOptionsChip editor={editor} />}
            <div className={rulersShown ? 'editor-area with-rulers' : 'editor-area'}>
              {rulersShown && rulerSection && (
                <RulerBar
                  editor={editor}
                  zoom={zoom / 100}
                  pageWidthPx={(rulerDims(rulerSection).pageWidth / 1440) * 96}
                >
                  <Ruler
                    section={rulerSection}
                    editor={editor}
                    onTabStopsChange={(stops) => {
                      if (!editor) return
                      setParaAttrs(editor, { tabStops: stops ? JSON.stringify(stops) : null })
                    }}
                    onIndentsChange={(attrs) => {
                      if (editor) setParaAttrs(editor, { ...attrs })
                    }}
                    onMarginsChange={({ left, right }) => {
                      const sec = sections[activeSection]?.settings ?? section
                      ribbonActions.onSection({
                        ...sec,
                        marginLeft: rulerSwap ? right : left,
                        marginRight: rulerSwap ? left : right,
                      })
                    }}
                  />
                </RulerBar>
              )}
              {rulersShown && rulerSection && (
                <VerticalRuler
                  section={rulerSection}
                  editor={editor}
                  zoom={zoom / 100}
                  top={rulerBarHeight(zoom / 100)}
                  onMarginsChange={({ top, bottom }) => {
                    const sec = sections[activeSection]?.settings ?? section
                    ribbonActions.onSection({ ...sec, marginTop: top, marginBottom: bottom })
                  }}
                />
              )}
              <main
                className={`editor-scroll${imageDragOver ? ' image-drop-target' : ''}${doc ? '' : ' start-screen-host'}`}
                ref={scrollContainerRef}
                onScroll={(e) => {
                  scrollPosRef.current = {
                    left: e.currentTarget.scrollLeft,
                    top: e.currentTarget.scrollTop,
                  }
                }}
                onDragOver={(e) => {
                  if (!editor?.isEditable || !isImageFileDrag(e.dataTransfer)) return
                  e.preventDefault()
                  if (!imageDragOver) setImageDragOver(true)
                }}
                onDragLeave={(e) => {
                  const next = e.relatedTarget as Node | null
                  if (!next || !e.currentTarget.contains(next)) setImageDragOver(false)
                }}
                onDrop={(e) => {
                  setImageDragOver(false)
                  const files = imageFilesFromDataTransfer(e.dataTransfer)
                  if (files.length === 0) return
                  // a drop the editor's own handler did not take (read-only
                  // islands, the page margin) still must not navigate the window
                  const handledByEditor = e.isDefaultPrevented()
                  e.preventDefault()
                  if (handledByEditor || !editor) return
                  if (editor.view.dom.contains(e.target as Node)) {
                    insertImageFilesAtCoords(editor, files, { left: e.clientX, top: e.clientY })
                  }
                }}
              >
                {doc && viewMode === 'draft' && (
                  <div className="draft-breaks" ref={draftBreaksRef} />
                )}
                {doc ? (
                  <div
                    className={docZoomClass}
                    onClick={onDocClick}
                    onContextMenu={onDocContextMenu}
                    onDoubleClick={onDocDoubleClick}
                    style={
                      multipage && mpLayout
                        ? { ...docZoomStyle, width: mpLayout.width, height: mpLayout.height }
                        : docZoomStyle
                    }
                  >
                    <div
                      className="doc-live"
                      style={
                        multipage && mpLayout
                          ? {
                              left: mpLayout.slots[mpActiveIdx].left,
                              top: mpLayout.slots[mpActiveIdx].top,
                              width: mpLayout.slots[mpActiveIdx].width,
                              height: mpLayout.slots[mpActiveIdx].height,
                            }
                          : undefined
                      }
                    >
                      <div
                        className={`${mixedPaper ? 'page-wrap paper-mixed' : 'page-wrap'}${hfEditing ? ' hf-editing' : ''}`}
                        style={
                          multipage
                            ? {
                                marginTop: -mpFrames[mpActiveIdx].top,
                                clipPath: `inset(${mpFrames[mpActiveIdx].top}px 0 calc(100% - ${mpFrames[mpActiveIdx].bottom}px) 0)`,
                              }
                            : undefined
                        }
                      >
                        {/* Boolean(): a trailing 0 (empty non-floating image list) must not render as a literal "0" text node */}
                        {Boolean(
                          multiHf ||
                          headerAreaView !== 'default' ||
                          hfAreaEditRequest?.kind === 'header' ||
                          hfHasVisibleContent(shownHeader, hfImagesOf('header')),
                        ) && (
                          <HeaderFooterArea
                            kind="header"
                            value={shownHeader ?? { text: '' }}
                            images={hfImagesOf('header')}
                            readOnly={isProtected || readMode}
                            onCommit={(next) => {
                              const target = edgeHfTarget('header', headerAreaView)
                              commitHfAt(0, 'header', target.variant, next, target.owner)
                            }}
                            pageNo={firstPageNoText}
                            pageTotal={pageInfo.total}
                            sectionPages={edgeSectionPages?.first}
                            sectionLabel={hfSectionLabel('header', 0)}
                            editRequest={
                              hfAreaEditRequest?.kind === 'header' ? hfAreaEditRequest.nonce : null
                            }
                            onEditingChange={(editing, handle) => {
                              hfEditorRef.current = editing ? handle : null
                              if (!editing) setHfAreaEditRequest(null)
                              setHfEditing(
                                editing
                                  ? {
                                      kind: 'header',
                                      section: 0,
                                      variant: headerAreaView,
                                      pageIdx: 0,
                                      owner: hfResolveAt(0, 'header', headerAreaView).owner,
                                    }
                                  : null,
                              )
                            }}
                            style={edgeHeaderStyle}
                            boxGeom={
                              canvasSection && canvasBox
                                ? {
                                    ...hfStripGeom(canvasSection, legacyBands),
                                    stripLeft: edgeHeaderStyle
                                      ? twipsToPx(firstPageMargins!.left)
                                      : (paperW - canvasBox.contentWidth) / 2,
                                  }
                                : undefined
                            }
                          />
                        )}
                        <EditorContent editor={editor} />
                        {/* footnotes already shown per page in page gaps aren't repeated at the end (last page's footnotes still live here) */}
                        <PageFootnotes
                          notes={footnotes}
                          skipIds={gapNoteIds}
                          numberOf={(n, i) => noteNo('footnote', n.id, i)}
                          fmtOf={(n) => noteFmt('footnote', n.id)}
                          onEdit={(id) => editNote('footnote', id)}
                          onDelete={(id) => deleteNote('footnote', id)}
                        />
                        <PageEndnotes
                          notes={tailEndnotes}
                          top={endnotesAreaTop}
                          numberOf={(n) => noteNo('endnote', n.id, endnotes.indexOf(n))}
                          fmtOf={(n) => noteFmt('endnote', n.id)}
                          onEdit={(id) => editNote('endnote', id)}
                          onDelete={(id) => deleteNote('endnote', id)}
                        />
                        {Boolean(
                          multiHf ||
                          footerAreaView !== 'default' ||
                          hfAreaEditRequest?.kind === 'footer' ||
                          hfHasVisibleContent(shownFooter, hfImagesOf('footer')),
                        ) && (
                          <HeaderFooterArea
                            kind="footer"
                            value={shownFooter ?? { text: '' }}
                            images={hfImagesOf('footer')}
                            readOnly={isProtected || readMode}
                            onCommit={(next) => {
                              const target = edgeHfTarget('footer', footerAreaView)
                              commitHfAt(
                                lastSectionIdx,
                                'footer',
                                target.variant,
                                next,
                                target.owner,
                              )
                            }}
                            pageNo={lastPageNoText}
                            pageTotal={pageInfo.total}
                            sectionPages={edgeSectionPages?.last}
                            linked={hfLinkedAt(lastSectionIdx, 'footer', footerAreaView)}
                            sectionLabel={hfSectionLabel('footer', lastSectionIdx)}
                            editRequest={
                              hfAreaEditRequest?.kind === 'footer' ? hfAreaEditRequest.nonce : null
                            }
                            onEditingChange={(editing, handle) => {
                              hfEditorRef.current = editing ? handle : null
                              if (!editing) setHfAreaEditRequest(null)
                              setHfEditing(
                                editing
                                  ? {
                                      kind: 'footer',
                                      section: lastSectionIdx,
                                      variant: footerAreaView,
                                      pageIdx: lastPageIdx,
                                      owner: hfResolveAt(lastSectionIdx, 'footer', footerAreaView)
                                        .owner,
                                    }
                                  : null,
                              )
                            }}
                            style={edgeFooterStyle}
                            boxGeom={
                              lastSection && lastBox
                                ? {
                                    ...hfStripGeom(lastSection, legacyBands),
                                    stripLeft: edgeFooterStyle
                                      ? twipsToPx(lastPageMargins!.left)
                                      : (paperW - lastBox.contentWidth) / 2,
                                  }
                                : undefined
                            }
                          />
                        )}
                        {(inkAnnotations.length > 0 || inkTool !== 'select') && !readMode && (
                          <InkOverlay
                            tool={isProtected ? 'select' : inkTool}
                            color={inkActive.color}
                            width={inkActive.width}
                            zoom={zoom}
                            annotations={inkAnnotations}
                            onAdd={addInk}
                            onRemove={removeInks}
                          />
                        )}
                      </div>
                    </div>
                    {/* after the live editor in DOM order: first-match lookups of
                        .ProseMirror / .doc-page / .page-wrap must find the editor */}
                    {multipage && mpLayout && (
                      <MultipageGrid
                        frames={mpFrames}
                        layout={mpLayout}
                        active={mpActiveIdx}
                        snap={mpSnap}
                        scroller={scrollContainerRef.current}
                        onPick={onMpPick}
                      />
                    )}
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
                      data-tip={t('appRemoveSplit')}
                      aria-label={t('appRemoveSplit')}
                      onClick={() => setSplitView(false)}
                    >
                      ×
                    </button>
                  </div>
                  <div className="split-pane-scroll">
                    <div className={docZoomClass.replace(' multipage', '')} style={docZoomStyle}>
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
            {doc && editor && viewMode === 'story' && !readMode && (
              <div className="story-bar" role="toolbar" aria-label={t('ribbonStoryView')}>
                <button
                  type="button"
                  className="story-bar-btn"
                  onClick={() => editor.chain().focus().toggleMark('bold').run()}
                >
                  {t('ribbonMarkerBold')}
                </button>
                <button
                  type="button"
                  className="story-bar-btn"
                  onClick={() => editor.chain().focus().toggleMark('italic').run()}
                >
                  {t('ribbonMarkerItalic')}
                </button>
                <button
                  type="button"
                  className="story-bar-btn"
                  onClick={() => setShowLinkModal(true)}
                >
                  {t('ribbonLink')}
                </button>
                <button
                  type="button"
                  className="story-bar-btn"
                  onClick={() => setListDialog('bullet')}
                >
                  {t('storyBarList')}
                </button>
                <span className="story-bar-gap" />
                <button
                  type="button"
                  className={`story-bar-btn${trackChanges ? ' on' : ''}`}
                  aria-pressed={trackChanges}
                  disabled={isProtected || trackChangesForced}
                  onClick={toggleTrackChangesByUser}
                >
                  {t('ribbonTrackChanges')}
                </button>
                <button
                  type="button"
                  className={`story-bar-btn${storyFocus ? ' on' : ''}`}
                  aria-pressed={storyFocus}
                  onClick={() => setStoryFocus((v) => !v)}
                >
                  {t('storyBarFocus')}
                </button>
                <span className="story-bar-count">{t('appWordCountN', { n: wordCount })}</span>
                <button
                  type="button"
                  className="story-bar-btn"
                  onClick={() => setViewMode(storyReturnView)}
                >
                  {t('storyBarLeave')}
                </button>
              </div>
            )}
            {doc && commentsUi.panel && (
              <CommentsPanel
                comments={comments}
                docNode={editor.state.doc}
                composing={commentComposing}
                onSubmitNew={submitNewComment}
                onReply={replyToComment}
                onEdit={editComment}
                onResolve={resolveComment}
                onCancelNew={cancelNewComment}
                onDelete={deleteComment}
                onClose={closeCommentsPanel}
                focus={commentFocus}
              />
            )}
            {doc && showStylesPane && (
              <StylesPane
                editor={editor}
                styles={ribbonStyles}
                docDefaults={doc.parsed.docDefaults}
                docName={doc.fileName}
                docBytes={doc.parsed.internal.originalBytes}
                fs={formatState}
                canEdit={formatState.editable}
                pending={styleUpserts}
                onUpsert={upsertStyleDef}
                onNotice={setStatus}
                onClose={() => setShowStylesPane(false)}
              />
            )}
            {doc && showCitationsPane && (
              <CitationsPane
                sources={sources}
                canInsert={formatState.editable}
                onInsert={(tag) => runCommand('insert-citation', tag)}
                onClose={() => setShowCitationsPane(false)}
              />
            )}
            {doc && editor && showSelectionPane && (
              <SelectionPane editor={editor} onClose={() => setShowSelectionPane(false)} />
            )}
            {doc && showReviewPane && (
              <ReviewingPane
                entries={reviewEntries}
                hiddenCount={allRevisions.length - shownRevisions.length}
                authors={reviewAuthors.map((name) => ({
                  name,
                  inkIndex: authorInkIndex(reviewAuthors, name),
                }))}
                openComments={comments.filter((c) => !c.parentId && !c.done).length}
                activeIndex={activeReviewIndex}
                canEdit={!isProtected && !trackChangesForced}
                onSelect={showRevision}
                onAccept={(rev) => handleRevisionRange('accept', rev)}
                onReject={(rev) => handleRevisionRange('reject', rev)}
                onClose={() => setShowReviewPane(false)}
              />
            )}
          </div>

          <footer className="status-bar">
            <div className="status-left">
              {doc && (
                <>
                  <button
                    className="status-item status-btn status-page"
                    data-tip={t('appGoToTitle')}
                    onClick={() => {
                      setShowFind(true)
                      setFindFocusGoTo((n) => n + 1)
                    }}
                  >
                    {t('appPageOf', { current: pageInfo.current, total: pageInfo.total })}
                  </button>
                  <button
                    className="status-item status-btn status-wordcount"
                    data-tip={t('appWordCountTitle')}
                    onClick={openStats}
                  >
                    {selectionWordCount > 0
                      ? t('appWordCountSelection', {
                          selected: selectionWordCount,
                          total: wordCount,
                        })
                      : t('appWordCountN', { n: wordCount })}
                  </button>
                  <button
                    className={`status-item status-btn status-track${trackChanges ? ' on' : ''}`}
                    data-tip={t('ribbonTrackChangesTip')}
                    aria-pressed={trackChanges}
                    disabled={isProtected || trackChangesForced}
                    onClick={toggleTrackChangesByUser}
                  >
                    {t(trackChanges ? 'appTrackChangesOn' : 'appTrackChangesOff')}
                  </button>
                </>
              )}
              {!doc && t('appReady')}
              {status && <span className="status-msg"> — {status}</span>}
            </div>
            <div className="status-right">
              {doc && (
                <div className="status-views" role="group">
                  <button
                    className={`status-view-btn${viewMode === 'print' && !readMode ? ' on' : ''}`}
                    data-tip={t('ribbonPrintLayout')}
                    aria-label={t('ribbonPrintLayout')}
                    onClick={() => switchView('print')}
                  >
                    <IconPrintLayout size={14} />
                  </button>
                  <button
                    className={`status-view-btn${viewMode === 'web' ? ' on' : ''}`}
                    data-tip={t('ribbonWebLayout')}
                    aria-label={t('ribbonWebLayout')}
                    onClick={() => switchView('web')}
                  >
                    <IconWebLayout size={14} />
                  </button>
                  <button
                    className={`status-view-btn${viewMode === 'outline' ? ' on' : ''}`}
                    data-tip={t('ribbonOutlineView')}
                    aria-label={t('ribbonOutlineView')}
                    onClick={() => switchView('outline')}
                  >
                    <IconOutlineView size={14} />
                  </button>
                  <button
                    className={`status-view-btn${viewMode === 'draft' ? ' on' : ''}`}
                    data-tip={t('ribbonDraft')}
                    aria-label={t('ribbonDraft')}
                    onClick={() => switchView('draft')}
                  >
                    <IconDraftView size={14} />
                  </button>
                  <button
                    className={`status-view-btn${viewMode === 'story' ? ' on' : ''}`}
                    data-tip={t('ribbonStoryView')}
                    aria-label={t('ribbonStoryView')}
                    onClick={() => switchView('story')}
                  >
                    <IconStoryView size={14} />
                  </button>
                  <button
                    className={`status-view-btn${readMode ? ' on' : ''}`}
                    data-tip={t('ribbonReadMode')}
                    aria-label={t('ribbonReadMode')}
                    onClick={() => setReadModeExclusive(!readMode)}
                  >
                    <IconReadMode size={14} />
                  </button>
                </div>
              )}
              <button className="zoom-btn" onClick={() => setZoom((z) => stepDocsZoom(z, -1))}>
                −
              </button>
              <input
                className="zoom-slider"
                type="range"
                min={DOCS_ZOOM_MIN}
                max={DOCS_ZOOM_MAX}
                step={10}
                value={Math.round(zoom)}
                onChange={(e) => setZoom(Number(e.target.value))}
              />
              <button className="zoom-btn" onClick={() => setZoom((z) => stepDocsZoom(z, 1))}>
                +
              </button>
              <button
                className="zoom-value"
                data-tip={t('ribbonZoomDialog')}
                disabled={!doc}
                onClick={() => setShowZoomDialog(true)}
              >
                {Math.round(zoom)}%
              </button>
            </div>
          </footer>
        </div>
      </div>

      {showShortcuts && <ShortcutsDialog onClose={() => setShowShortcuts(false)} />}
      {showCustomizeKeyboard && (
        <CustomizeKeyboardDialog
          styles={ribbonStyles ?? FALLBACK_STYLES}
          onClose={() => setShowCustomizeKeyboard(false)}
        />
      )}
      {showLinkModal && <LinkInsertModal editor={editor} onClose={() => setShowLinkModal(false)} />}
      {showBookmarkModal && (
        <BookmarkModal editor={editor} onClose={() => setShowBookmarkModal(false)} />
      )}
      {pictureDialog &&
        editor &&
        createElement(pictureDialog.kind === 'crop' ? CropDialog : CutoutDialog, {
          dataUrl: pictureDialog.dataUrl,
          onApply: (png: string) => {
            setPictureDialog(null)
            void replaceSelectedPictureBytes(editor, png)
          },
          onCancel: () => setPictureDialog(null),
        })}
      {showCaptionModal && editor && (
        <CaptionModal
          editor={editor}
          blocks={doc?.parsed.blocks ?? EMPTY_BLOCKS}
          onClose={() => setShowCaptionModal(false)}
        />
      )}
      {showCrossRefModal && (
        <CrossRefModal
          editor={editor}
          blocks={doc?.parsed.blocks ?? EMPTY_BLOCKS}
          notes={{ footnote: footnotes, endnote: endnotes }}
          pagesOf={(positions) => commandHostRef.current?.pagesOf(positions) ?? null}
          onClose={() => setShowCrossRefModal(false)}
        />
      )}
      {chartModal && (
        <ChartInsertModal
          editor={editor}
          initialType={chartModal.type}
          pick={chartModal.pick}
          onClose={() => setChartModal(null)}
        />
      )}
      {showChangeCase && (
        <ChangeCaseDialog
          onApply={(mode) => runCommand('change-case', mode)}
          onClose={() => setShowChangeCase(false)}
        />
      )}
      {showColumnsDialog && (
        <ColumnsDialog
          key={activeSection}
          section={sections[activeSection]?.settings ?? section ?? DEFAULT_SECTION}
          multiSection={sections.length > 1}
          canApplyForward={canApplyForward}
          onApply={(value) => runCommand('columns-settings', JSON.stringify(value))}
          onClose={() => setShowColumnsDialog(false)}
        />
      )}
      {editor && doc && mergePanel && (
        <MailMergePanel
          editor={editor}
          source={mergeRecipients}
          onPickSource={() => void pickMergeDataSource()}
          onInsertField={insertField}
          onMerge={mergeMailings}
          recordIndex={mergeRecordIndex}
          onRecordIndex={setMergeRecordIndex}
          range={mergeRange}
          initialMode={mergePanel.mode}
          pick={mergePanel.pick}
          onLabelSheet={(layout) => {
            if (!createLabelSheet(editor, layout)) return
            ribbonActions.onSection(layoutSection(layout, section))
            setMergePanel(null)
          }}
          onClose={() => setMergePanel(null)}
        />
      )}
      {mergeQuestion && (
        <PromptModal
          key={mergeQuestion.prompt.key}
          title={
            mergeQuestion.prompt.record === null
              ? mergeQuestion.prompt.prompt
              : `${mergeQuestion.prompt.prompt} (${t('ribbonMailingsRecord', {
                  n: mergeQuestion.prompt.record,
                  total: mergeQuestion.total,
                })})`
          }
          initial={mergeQuestion.prompt.fallback}
          allowEmpty
          onSubmit={mergeQuestion.answer}
          onClose={() => {
            // closing without an answer takes the field's default
            mergeQuestion.answer(mergeQuestion.prompt.fallback)
            setMergeQuestion((open) => (open === mergeQuestion ? null : open))
          }}
        />
      )}
      {autoText && editor && (
        <AutoTextDialog
          {...autoText}
          onSave={(name) => saveAutoText(editor, name)}
          onInsert={(entry) => insertAutoText(editor, entry)}
          onDelete={(name) => {
            deleteAutoText(name)
            setAutoText({ ...autoText, entries: autoTextEntries() })
          }}
          onClose={() => setAutoText(null)}
        />
      )}
      {markCitationAt && editor && (
        <MarkCitationDialog
          selected={markCitationAt.text}
          onMark={(citation, all) =>
            markCitation(
              editor,
              markCitationAt.at,
              {
                ...citation,
                runs: citationRunsOf(
                  editor.state.doc,
                  markCitationAt.from,
                  markCitationAt.at,
                  citation.long,
                ),
              },
              all
                ? occurrenceEnds(editor, citation.short ?? '', true).filter(
                    (pos) => pos <= markCitationAt.from || pos > markCitationAt.at,
                  )
                : [],
            )
          }
          onClose={() => setMarkCitationAt(null)}
        />
      )}
      {markEntryAt && editor && (
        <MarkIndexEntryDialog
          selected={markEntryAt.text}
          bookmarks={collectBookmarks(editor.state.doc).map((b) => b.name)}
          onMark={(entry, all) =>
            markIndexEntries(
              editor,
              all ? occurrenceEnds(editor, markEntryAt.text) : [markEntryAt.at],
              indexEntryTerm(entry),
              indexEntrySwitches(entry),
            )
          }
          onClose={() => setMarkEntryAt(null)}
        />
      )}
      {languageTarget && (
        <LanguageDialog
          current={proofingLanguageOf(languageTarget)}
          noProof={noProofOf(languageTarget)}
          onApply={(tag, skip) => setProofingLanguage(languageTarget, tag, skip)}
          onClose={() => setLanguageTarget(null)}
        />
      )}
      {textEffectsTarget && (
        <TextEffectsDialog
          fx={textFxOf(textEffectsTarget)}
          color={textFillOf(textEffectsTarget)}
          onApply={(fx, color) => applyTextFx(textEffectsTarget, fx, color)}
          onClose={() => setTextEffectsTarget(null)}
        />
      )}
      {fitTextTarget && (
        <FitTextDialog
          width={fitTextOf(fitTextTarget)}
          onApply={(twips) => setFitText(fitTextTarget, twips)}
          onClose={() => setFitTextTarget(null)}
        />
      )}
      {eqLayoutDialog?.kind === 'enclose' && (
        <EncloseCharactersDialog
          text={eqLayoutDialog.target.existing?.b ?? eqLayoutDialog.target.text}
          enclosure={eqLayoutDialog.target.existing?.a ?? null}
          style={eqLayoutDialog.target.existing?.style ?? null}
          onApply={(layout) =>
            applyEqLayout(
              eqLayoutDialog.editor,
              eqLayoutDialog.target,
              layout && { kind: 'enclose', ...layout },
            )
          }
          onClose={() => setEqLayoutDialog(null)}
        />
      )}
      {eqLayoutDialog?.kind === 'combine' && (
        <CombineCharactersDialog
          text={eqLayoutDialog.target.text}
          combined={!!eqLayoutDialog.target.existing}
          {...eqLayoutFontOf(eqLayoutDialog.editor.state, eqLayoutDialog.target)}
          onApply={(text, picked) =>
            applyEqLayout(
              eqLayoutDialog.editor,
              eqLayoutDialog.target,
              text === null ? null : { kind: 'combine', text },
              picked,
            )
          }
          onClose={() => setEqLayoutDialog(null)}
        />
      )}
      {phoneticGuide && (
        <PhoneticGuideDialog
          base={phoneticGuide.target.base}
          rt={phoneticGuide.target.rt}
          align={phoneticGuide.target.align}
          reading={phoneticGuide.target.reading}
          onApply={(edit) =>
            applyRuby(
              phoneticGuide.editor,
              phoneticGuide.target,
              edit,
              doc?.parsed.docDefaults?.eastAsiaLang ?? RUBY_LANGS[lang] ?? 'zh-CN',
            )
          }
          onClose={() => setPhoneticGuide(null)}
        />
      )}
      {tabsTarget && (
        <TabsDialog
          stops={parseTabStops(activeParaAttrs(tabsTarget).tabStops).stops}
          defaultStop={editor?.storage.tabStops.defaultTabStopTwips ?? 720}
          onApply={(stops, defaultStop) => {
            setParaAttrs(tabsTarget, {
              tabStops: tabStopsAttr(activeParaAttrs(tabsTarget).tabStops, stops),
            })
            if (editor && defaultStop !== (editor.storage.tabStops.defaultTabStopTwips ?? 720)) {
              editor.storage.tabStops.defaultTabStopTwips = defaultStop
              editor.storage.listNumbering.defaultTabTwips = defaultStop
              setDefaultTabStopEdit(defaultStop)
              document.dispatchEvent(new Event(DEFAULT_TAB_STOP_CHANGED_EVENT))
            }
          }}
          onClose={() => setTabsTarget(null)}
        />
      )}
      {mergeSource && showRecipients && (
        <RecipientsDialog
          source={mergeSource}
          excluded={mergeExcluded}
          onApply={setMergeExcluded}
          onClose={() => setShowRecipients(false)}
        />
      )}
      {mergeSource && showFilterRecipients && (
        <FilterRecipientsDialog
          fields={mergeSource.fields}
          query={mergeQuery}
          onApply={setMergeQuery}
          onClose={() => setShowFilterRecipients(false)}
        />
      )}
      {mergeRecipients && showFindRecipient && (
        <FindRecipientDialog
          fields={mergeRecipients.fields}
          onFind={(text, field) => {
            const found = findRecipient(mergeRecipients.records, mergeRecordIndex, text, field)
            if (found >= 0) setMergeRecordIndex(found)
            return found >= 0
          }}
          onClose={() => setShowFindRecipient(false)}
        />
      )}
      {mergeRule && (mergeSource || !ruleNeedsList(mergeRule)) && (
        <RuleDialog
          rule={mergeRule}
          fields={mergeSource?.fields ?? []}
          onInsert={insertField}
          onClose={() => setMergeRule(null)}
        />
      )}
      {doc && showFontSubstitution && (
        <FontSubstitutionDialog
          fonts={checkMissingFonts(
            [...new Set([...collectDocFonts(doc.parsed), ...fontsBeingRestored(fontConversions)])],
            fontSubstituteOf,
          ).filter((font) => !(font.name in (fontConversions?.pending ?? {})))}
          onSubstitute={setFontSubstitute}
          onConvert={(fonts) => {
            convertFontsInEditor(editor, fonts)
            setFontConversions((state) => withFontsConverted(state, fonts))
          }}
          onClose={() => setShowFontSubstitution(false)}
        />
      )}
      {showBordersShading && (
        <BordersShadingDialog
          {...bordersShadingTarget(editor)}
          section={section}
          onCommand={runCommand}
          onPageBorder={(props, scope) => {
            if (scope === 'whole') return ribbonActions.onPageBorderAll(props)
            if (section) ribbonActions.onSection(sectionWithPageBorder(section, props))
          }}
          onHorizontalLine={
            editor.isEditable && !getActiveSubEditor()
              ? () => runCommand('insert-horizontal-line')
              : undefined
          }
          onClose={() => setShowBordersShading(false)}
        />
      )}
      {doc && showRename && (
        <RenameDialog
          fileName={doc.fileName}
          onRename={(name) => window.desktop.renameDoc(name)}
          onClose={() => setShowRename(false)}
        />
      )}
      {editor && showNumber && (
        <NumberDialog
          onInsert={(text) =>
            (getActiveSubEditor() ?? editor).chain().focus().insertContent(text).run()
          }
          onClose={() => setShowNumber(false)}
        />
      )}
      {editor && showSort && (
        <SortDialog
          loadScope={() => sortDialogScope(editor.state)}
          onSort={(options) => {
            const refused = sortSelection(editor, options)
            if (refused) setStatus(t(SORT_REFUSALS[refused]))
          }}
          onClose={() => setShowSort(false)}
        />
      )}
      {showParagraphSpacing && (
        <ParagraphSpacingDialog
          current={defaultParagraphSpacing()}
          onApply={(next) => runCommand('design-paragraph-spacing', JSON.stringify(next))}
          onClose={() => setShowParagraphSpacing(false)}
        />
      )}
      {showTemplateGallery && (
        <TemplateGalleryDialog
          // the gallery stays up when the picker is cancelled or the replace is declined
          onCreate={(name) =>
            newFromTemplate(name).then((made) => made && setShowTemplateGallery(false))
          }
          onBrowse={() => newFromTemplate().then((made) => made && setShowTemplateGallery(false))}
          onClose={() => setShowTemplateGallery(false)}
        />
      )}
      {insertFileRange && (
        <InsertFileRangeDialog
          fileName={insertFileRange.fileName}
          bookmarks={insertFileRange.bookmarks}
          onAnswer={(range) => {
            insertFileRange.answer(range)
            setInsertFileRange(null)
          }}
        />
      )}
      {doc && showReduceFileSize && (
        <ReduceFileSizeDialog
          selectedPictures={reduceSelection.length}
          onReduce={(ppi, deleteCropped, selectedOnly) =>
            reduceFileSize(fileCtxRef.current, {
              ppi,
              deleteCropped,
              ...(selectedOnly ? { only: reduceSelection } : {}),
            })
          }
          onClose={() => setShowReduceFileSize(false)}
        />
      )}
      {showLineNumbers && (
        <LineNumbersDialog
          current={(sections[activeSection]?.settings ?? section)?.lineNumbers}
          onApply={(next) =>
            runCommand('layout-line-numbers', next ? JSON.stringify(next) : 'none')
          }
          onClose={() => setShowLineNumbers(false)}
        />
      )}
      {editor && showAltText && (
        <AltTextDialog
          current={selectedPictureAlt(editor) ?? ''}
          onApply={(alt) => setSelectedPictureAlt(editor, alt)}
          onClose={() => setShowAltText(false)}
        />
      )}
      {editor && showSizePosition && selectedPictureGeometry(editor) && (
        <SizePositionDialog
          current={selectedPictureGeometry(editor)!}
          onApply={(next) => applyPictureGeometry(editor, next)}
          onClose={() => setShowSizePosition(false)}
        />
      )}
      {editor && showTextDirection && (
        <TextDirectionDialog
          context={textDirectionContext(
            editor,
            formatState.cellTextDirection,
            sections[activeSection]?.settings ?? section,
          )}
          onApply={(direction) => runCommand('format-text-direction', direction)}
          onClose={() => setShowTextDirection(false)}
        />
      )}
      {doc && showProperties && (
        <PropertiesDialog
          fileName={doc.fileName}
          properties={doc.parsed.properties}
          pending={propertyChanges}
          loadStats={documentStats}
          loadHeadings={() =>
            editor ? collectHeadings(editor.state.doc, editor.storage.listNumbering?.styles) : []
          }
          editable={formatState.editable}
          loadFileInfo={async () => (doc.filePath ? window.desktop.fileInfo() : null)}
          onApply={(changes) => {
            setPropertyChanges((pending) => ({
              core: { ...pending?.core, ...changes.core },
              app: { ...pending?.app, ...changes.app },
              custom: { ...pending?.custom, ...changes.custom },
              customKinds: { ...pending?.customKinds, ...changes.customKinds },
            }))
            dirtyRef.current = true
          }}
          onClose={() => setShowProperties(false)}
        />
      )}
      {showRevealFormatting && (
        <RevealFormattingDialog
          format={formatState}
          styleName={revealedStyleName(
            formatState.paraStyleId ?? defaultParagraphStyleId(ribbonStyles) ?? 'Normal',
            ribbonStyles,
          )}
          onClose={() => setShowRevealFormatting(false)}
        />
      )}
      {showHyphenation && (
        <HyphenationDialog
          auto={autoHyphenation}
          settings={hyphenationSettings}
          hasPatterns={hyphenationDictFor(doc?.parsed.docDefaults?.lang) !== null}
          zoneApplies={hyphenationZoneApplies(doc?.parsed)}
          canEdit={editor?.isEditable === true}
          onApply={(next) => {
            armPageSetupUndo()
            applyHyphenation(next)
          }}
          onManual={(next) => {
            armPageSetupUndo()
            applyHyphenation({ ...next, auto: false }, true)
          }}
          onClose={() => setShowHyphenation(false)}
        />
      )}
      {manualHyphenation && editor && (
        <ManualHyphenationDialog
          editor={editor}
          onClose={() => applyHyphenation({ ...hyphenationSettings, auto: autoHyphenation })}
        />
      )}
      {showEquationOptions && (
        <EquationOptionsDialog
          current={mathPr}
          onApply={applyMathPr}
          onClose={() => setShowEquationOptions(false)}
        />
      )}
      {showIndexTables && (
        <IndexTablesDialog
          captionLabels={editor && doc ? captionLabels(editor, doc.parsed.blocks) : []}
          existing={indexTablesAtSelection(editor, doc?.parsed.blocks ?? [])}
          styles={ribbonStyles}
          pendingStyles={styleUpserts}
          onUpsertStyle={upsertStyleDef}
          onApply={(kind, spec, formatStyles) => {
            formatStyles.forEach(upsertStyleDef)
            runCommand(INDEX_TABLES_COMMAND[kind], tableSpecPayload(spec))
          }}
          onAutoMark={() => runCommand('index-automark')}
          onMarkCitation={() => runCommand('toa-mark-citation')}
          onMarkEntry={() => runCommand('mark-index-entry')}
          onClose={() => setShowIndexTables(false)}
        />
      )}
      {showTranslationLanguage && (
        <TranslationLanguageDialog
          current={translationLanguage(lang)}
          onApply={setTranslationLanguage}
          onClose={() => setShowTranslationLanguage(false)}
        />
      )}
      {showBulletsNumbering && (
        <BulletsNumberingDialog
          onApply={({ kind, levels }) =>
            levels
              ? runCommand('list-preset', JSON.stringify({ levels, recent: kind }))
              : runCommand('list-none')
          }
          onCustomize={(kind, card) =>
            runCommand('list-dialog', card === undefined ? kind : `${kind}#${card}`)
          }
          listStyles={listStyles}
          listStylesInUse={listStylesInUse(editor, doc?.parsed.numbering)}
          initialStyleId={addedListStyle}
          onAddListStyle={() => runCommand('list-dialog', 'style')}
          onDeleteListStyle={(styleId) => runCommand('list-style-delete', styleId)}
          onApplyListStyle={(styleId) =>
            styleId === NO_LIST_STYLE_ID
              ? runCommand('list-none')
              : runCommand('list-style', styleId)
          }
          numbering={listNumberingState(editor)}
          onNumbering={(mode) => (mode === 'restart' ? restartNumbering() : continueNumbering())}
          onClose={() => {
            setShowBulletsNumbering(false)
            setAddedListStyle(undefined)
          }}
        />
      )}
      {showPageNumbers && (
        <PageNumbersDialog
          showOnFirst={!activeTitlePg}
          firstLocked={activeSection < sections.length - 1}
          onApply={(place, align, showOnFirst) => {
            runCommand('page-number', `${place},${align}`)
            // Word hides the number on the first page by giving that page a header and footer of its own
            if (showOnFirst === activeTitlePg) toggleTitlePg(!showOnFirst)
          }}
          onFormat={() => runCommand('page-number-format')}
          onClose={() => setShowPageNumbers(false)}
        />
      )}
      {showIconsDialog && (
        <IconsDialog
          onInsert={async (names) => {
            await prepareLibraryIcons(names)
            setShowIconsDialog(false)
            runCommand('insert-icon', names.join(','))
          }}
          onClose={() => setShowIconsDialog(false)}
        />
      )}
      {showSymbolDialog && (
        <SymbolDialog
          onInsert={(symbol, font) =>
            font
              ? runCommand('insert-font-symbol', fontSymbolPayload(font, symbol))
              : runCommand('insert-symbol', symbol)
          }
          onClose={() => setShowSymbolDialog(false)}
        />
      )}
      {showWatermarkDialog && (
        <WatermarkDialog
          current={{
            text: watermark,
            style: watermarkStyle,
            picture: watermarkPicture,
            docImage: watermarkDirty
              ? null
              : (doc?.parsed.headerImages?.find((img) => img.wordArt || img.watermark) ?? null),
          }}
          presets={watermarkPresetsOf(t, lang)}
          pageBox={watermarkPageBoxOf(sections[activeSection]?.settings ?? section)}
          onApply={ribbonActions.onWatermark}
          onClose={() => setShowWatermarkDialog(false)}
        />
      )}
      {showDropCap && editor && (
        <DropCapDialog
          current={dropCapValueOf(editor.getAttributes('docParagraph').dropCap)}
          font={(() => {
            const letter = dropCapFontOf(editor.state)
            const shown = letter.eastAsian ? formatState.fontEastAsia : formatState.fontLatin
            return letter.font || shown || ''
          })()}
          onApply={({ position, lines, gap }, font) =>
            runCommand(
              'drop-cap',
              position === 'none'
                ? NONE
                : `${position},${lines},${gap}${font === undefined ? '' : `,${font}`}`,
            )
          }
          onClose={() => setShowDropCap(false)}
        />
      )}
      {showDateTime && (
        <DateTimeDialog
          language={editor ? proofingLanguageOf(getActiveSubEditor() ?? editor) : ''}
          onInsert={({ picture, asField, lang, fullWidth }) => {
            if (!asField) {
              const text = formatDatePicture(new Date(), picture, lang)
              runCommand('insert-symbol', fullWidth ? toFullWidth(text) : text)
              return
            }
            const instr = `DATE \\@ "${picture}"${fullWidth ? ' \\* DBCHAR' : ''}`
            const marked = editor ? proofingLanguageOf(getActiveSubEditor() ?? editor) : ''
            if (lang === 'en-US' && !marked) runCommand('insert-field', instr)
            else runCommand('insert-date-field', `${lang},${instr}`)
          }}
          onClose={() => setShowDateTime(false)}
        />
      )}
      {showInsertField && (
        <FieldDialog
          instr=""
          pageIsLive={!!getActiveSubEditor()?.schema.nodes.hfField}
          onClose={() => setShowInsertField(false)}
          onSubmit={(instr) => runCommand('insert-field', instr)}
        />
      )}
      {pasteSpecialTarget && (
        <PasteSpecialDialog
          editor={pasteSpecialTarget}
          onClose={() => setPasteSpecialTarget(null)}
        />
      )}
      {fieldEdit && (
        <FieldDialog
          instr={fieldEdit.instr}
          onClose={() => setFieldEdit(null)}
          onSubmit={(instr) => {
            if (!editor || !setFieldInstr(editor, fieldEdit, instr, fieldValue(instr))) return
            // a formula or a caption number is computed here, so its new code needs no update by Word
            const tr = editor.state.tr
            const blocks = doc?.parsed.blocks ?? EMPTY_BLOCKS
            refreshTableFormulas(tr)
            renumberSeqFields(tr, blocks, {
              skip: unreliableSeqLabelsOf(blocks, editor.schema),
            })
            if (tr.docChanged) editor.view.dispatch(tr.setMeta('appendedTransaction', tr))
          }}
        />
      )}
      {showTableModal && (
        <TableInsertModal
          editor={editor}
          onClose={() => setShowTableModal(false)}
          onInserted={() => setRibbonTabRequest({ tab: 'tableDesign', nonce: Date.now() })}
        />
      )}
      {showZoomDialog && (
        <ZoomDialog
          zoom={zoom}
          fitValue={fitZoomFor}
          onApply={(target) => (typeof target === 'number' ? setZoom(target) : zoomFit(target))}
          onClose={() => setShowZoomDialog(false)}
        />
      )}
      {showDocumentDialog && (sections[activeSection]?.settings ?? section) && (
        <DocumentDialog
          // the caret's section can settle after the dialog opens: it starts over on that section
          key={activeSection}
          sectionIndex={activeSection}
          section={(sections[activeSection]?.settings ?? section)!}
          startType={activeStartType}
          startLocked={!!sections[activeSection]?.pendingBreak}
          gutterAtTop={gutterAtTop}
          multiSection={sections.length > 1}
          mirror={mirrorMargins}
          titlePg={activeTitlePg}
          titlePgLocked={activeSection < sections.length - 1}
          evenOdd={evenOddHf}
          gridFont={gridFont}
          canApplyForward={canApplyForward}
          canSuppressEndnotes={
            noteSettings.endnote.pos === 'sectEnd' && activeSection < sections.length - 1
          }
          onApply={(changes) => {
            const {
              titlePg: _titlePg,
              evenOdd: _evenOdd,
              gridFont: fontChanges,
              ...sectionChanges
            } = changes
            // the font first: the character pitch that comes with it is counted from its size
            if (fontChanges) {
              const styleId = defaultParagraphStyleId(ribbonStyles) ?? 'Normal'
              const { latin, eastAsia, ...rest } = fontChanges
              upsertStyleDef({
                styleId,
                ...(ribbonStyles?.has(styleId) ? {} : { builtIn: true }),
                rPr: {
                  ...rest,
                  ...(latin !== undefined && { font: latin }),
                  ...(eastAsia !== undefined && { eastAsiaFont: eastAsia }),
                },
              })
              if (fontChanges.sizeHalfPoints) gridBasePtRef.current = fontChanges.sizeHalfPoints / 2
            }
            if (Object.keys(sectionChanges).length > 0)
              runCommand('document-settings', JSON.stringify(sectionChanges))
            if (changes.titlePg !== undefined) toggleTitlePg(changes.titlePg)
            if (changes.evenOdd !== undefined) toggleEvenOddHf(changes.evenOdd)
          }}
          onDefault={(page) => {
            savePageDefault(page)
            setStatus(t('ribbonDocDefaultDone'))
          }}
          onPageSetup={() => runCommand('page-setup')}
          onBorders={() => runCommand('format-borders-shading')}
          onLineNumbers={() => runCommand('layout-line-numbering-options')}
          onClose={() => setShowDocumentDialog(false)}
        />
      )}
      {showHighlightChanges && (
        <HighlightChangesDialog
          trackChanges={trackChanges}
          trackLocked={isProtected || trackChangesForced}
          display={revisionDisplay}
          print={printHighlight}
          trackMoves={trackMoves}
          onApply={({ track, moves, display, print }) => {
            if (track !== trackChanges) setTrackChangesByUser(track)
            if (moves !== trackMoves) setTrackMovesEdit(moves)
            if (display) setRevisionDisplay(display)
            setPrintHighlight(print)
          }}
          onOptions={() => setShowTcOptions(true)}
          onClose={() => setShowHighlightChanges(false)}
        />
      )}
      {compareDialog && doc && (
        <CompareDocumentsDialog
          mode={compareDialog}
          currentName={doc.fileName}
          userName={authorName()}
          canWrite={!!editor?.isEditable}
          onPick={pickCompareFile}
          onRun={(request) => void runCompare(reviewCtxRef.current, request)}
          onClose={() => setCompareDialog(null)}
        />
      )}
      {showTcOptions && (
        <TrackChangesOptionsDialog
          options={tcOptions}
          onApply={setTcOptions}
          onClose={() => setShowTcOptions(false)}
        />
      )}
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
          onCommand={runCommand}
          onRedact={startRedaction}
          onUnredact={clearRedaction}
          onViewImage={setViewImage}
          onSaveImageAs={saveImageAs}
          onRestartNumbering={restartNumbering}
          onContinueNumbering={continueNumbering}
          onChangeListLevel={changeListLevel}
          onUpdateFields={updateFields}
          onEditField={setFieldEdit}
          onOpenLink={(href) => openLink(href, ctxMenu.link?.tocTitle)}
          onRespell={() => requestAnimationFrame(() => setRespellNonce((n) => n + 1))}
          revisions={
            isProtected || trackChangesForced
              ? null
              : {
                  at: revisionsAtCtx(
                    liveRevisions,
                    ctxMenu.pos ?? editor.state.selection.from,
                    !!ctxMenu.balloon,
                  ),
                  any: liveRevisions.length > 0,
                  onRange: handleRevisionRange,
                  onAll: (action) => handleRevision(action, 'all'),
                }
          }
          comment={{ id: ctxCommentId }}
          trackChanges={isProtected || trackChangesForced ? null : { on: trackChanges }}
        />
      )}
      {viewImage && (
        <ImageViewer
          src={viewImage}
          lang={lang}
          labels={{
            zoomIn: t('ribbonZoomIn'),
            zoomOut: t('ribbonZoomOut'),
            actualSize: t('appImgActualSize'),
            fitToWindow: t('appImgFitWindow'),
            save: t('appSaveImageAs'),
            close: t('appCloseEsc'),
          }}
          onClose={() => setViewImage(null)}
          onSave={() => saveImageAs(viewImage)}
        />
      )}
      {doc && showFontDialog && (
        <FontDialog
          editor={editor}
          onDefault={(font) => {
            saveFontDefault(font)
            setStatus(t('appFontDefaultDone'))
          }}
          onClose={() => setShowFontDialog(false)}
        />
      )}
      {doc && showParaDialog && (
        <ParagraphDialog
          editor={editor}
          styles={ribbonStyles}
          pageWidth={section?.pageWidth}
          typography={{
            compressPunctuation: compressPunctEdit ?? doc?.parsed.compressPunctuation === true,
            strictKinsoku: strictKinsokuEdit ?? doc?.parsed.strictFirstAndLastChars === true,
            onChange: (compress) => {
              setCompressPunctEdit(compress)
              // the same two switches the file's value sets on open; every paragraph re-measures
              editor.storage.cjkPunctShrink.enabled = compress
              editor.storage.cjkPunctShrink.hangPunct = !compress
              document.dispatchEvent(new Event(DOC_CSS_COMMITTED_EVENT))
            },
            onStrictKinsoku: (strict) => {
              setStrictKinsokuEdit(strict)
              setStrictKinsoku(editor, strict)
            },
          }}
          onClose={() => setShowParaDialog(false)}
          onTabs={() => runCommand('format-tabs')}
          onSetDefault={(pPr) => {
            if (Object.keys(pPr).length === 0) return
            const styleId = defaultParagraphStyleId(ribbonStyles) ?? 'Normal'
            upsertStyleDef({
              styleId,
              ...(ribbonStyles?.has(styleId) ? {} : { builtIn: true }),
              pPr,
            })
            setStatus(t('appParaDefaultDone'))
          }}
        />
      )}
      {doc && redactTarget && (
        <RedactDialog
          seed={redactTarget.seed}
          onCancel={() => setRedactTarget(null)}
          onSubmit={(label) => {
            // re-assert the captured range: the command marks the selection,
            // and the reader may have clicked somewhere while deciding
            if (editor)
              editor
                .chain()
                .setTextSelection({ from: redactTarget.from, to: redactTarget.to })
                .setRedaction(label)
                .run()
            setRedactTarget(null)
          }}
        />
      )}
      {doc && tableDialog === 'properties' && (
        <TablePropertiesDialog
          editor={editor}
          sectionContentWidthPx={tableSectionWidthPx}
          onClose={() => setTableDialog(null)}
        />
      )}
      {doc && tableDialog === 'splitCells' && (
        <SplitCellsDialog editor={editor} onClose={() => setTableDialog(null)} />
      )}
      {doc && tableDialog === 'insertCells' && (
        <InsertCellsDialog
          editor={editor}
          sectionContentWidthPx={tableSectionWidthPx}
          onClose={() => setTableDialog(null)}
        />
      )}
      {doc && tableDialog === 'deleteCells' && (
        <DeleteCellsDialog editor={editor} onClose={() => setTableDialog(null)} />
      )}
      {doc && tableDialog === 'textToTable' && (
        <ConvertTextToTableDialog
          editor={editor}
          sectionContentWidthPx={tableSectionWidthPx}
          onClose={() => setTableDialog(null)}
        />
      )}
      {doc && tableDialog === 'tableToText' && (
        <ConvertTableToTextDialog editor={editor} onClose={() => setTableDialog(null)} />
      )}
      {doc && tableDialog === 'formula' && (
        <FormulaDialog editor={editor} onClose={() => setTableDialog(null)} />
      )}
      {doc && tableDialog === 'cellMargins' && (
        <CellMarginsDialog
          editor={editor}
          sectionContentWidthPx={tableSectionWidthPx}
          onClose={() => setTableDialog(null)}
        />
      )}
      {doc && listDialog === 'bullet' && (
        <DefineBulletDialog
          // a card is defined from its first level, whatever level the caret is on
          ilvl={listCard === null ? (listTarget?.ilvl ?? 0) : 0}
          initial={listCard === null ? undefined : listLibrary('bullets')[listCard]?.[0]}
          onApply={(levels, picture) => {
            if (picture) {
              const id = addPictureBullet(numberingCtxRef.current, picture)
              const at = listCard === null ? (listTarget?.ilvl ?? 0) : 0
              levels = levels.map((l, i) => (i === at ? { ...l, picBulletId: id } : l))
            } else if (listCard !== null) customizeLibraryCard('bullets', listCard, levels)
            applyDefinedList(levels, 'bullets')
          }}
          onClose={() => setListDialog(null)}
        />
      )}
      {doc && listDialog === 'number' && (
        <DefineNumberFormatDialog
          ilvl={listCard === null ? (listTarget?.ilvl ?? 0) : 0}
          initial={listCard === null ? undefined : listLibrary('numbers')[listCard]?.[0]}
          onApply={(levels) => {
            if (listCard !== null) customizeLibraryCard('numbers', listCard, levels)
            applyDefinedList(levels, 'numbers')
          }}
          onClose={() => setListDialog(null)}
        />
      )}
      {doc && listDialog === 'multi' && (
        <DefineMultilevelDialog
          styles={[...doc.parsed.styles.values()]}
          initial={listCard === null ? undefined : listLibrary('multi')[listCard]}
          onApply={(levels) => {
            if (listCard !== null) customizeLibraryCard('multi', listCard, levels)
            applyDefinedList(levels, 'multi')
          }}
          onClose={() => setListDialog(null)}
        />
      )}
      {doc && listDialog === 'style' && (
        <DefineMultilevelDialog
          styles={[]}
          takenStyleNames={[
            ...listStyles.map((s) => s.name.toLowerCase()),
            ...[...doc.parsed.styles.values()].map((s) => s.name.toLowerCase()),
          ]}
          onApply={(levels, name) => runCommand('list-style-add', JSON.stringify({ name, levels }))}
          onClose={() => setListDialog(null)}
        />
      )}
      {doc && listDialog === 'value' && listTarget && (
        <SetNumberingValueDialog
          numFmt={listTarget.level?.numFmt ?? 'decimal'}
          lvlText={listTarget.level?.lvlText ?? '%1.'}
          currentValue={currentListValue(editor, listTarget.numId, listTarget.ilvl)}
          canContinue={hasPreviousList(editor, listTarget.numId)}
          onApply={(mode, value) => {
            setNumberingValueImpl(numberingCtxRef.current, { mode, value })
            setListDialog(null)
            editor?.commands.focus()
          }}
          onClose={() => setListDialog(null)}
        />
      )}
      {doc && listDialog === 'indents' && listTarget && (
        <AdjustListIndentsDialog
          level={
            listTarget.level
              ? customLevelFromNumberingLevel(listTarget.level)
              : { numFmt: 'decimal', lvlText: '%1.', indentLeft: 720, hanging: 360 }
          }
          onApply={(patch) => {
            editListLevelImpl(numberingCtxRef.current, patch)
            setListDialog(null)
            editor?.commands.focus()
          }}
          onClose={() => setListDialog(null)}
        />
      )}

      {doc && showFootnoteDialog && (
        <FootnoteDialog
          settings={{
            footnote: activeNoteProps('footnote'),
            endnote: activeNoteProps('endnote'),
          }}
          multiSection={sections.length > 1}
          counts={{ footnote: footnotes.length, endnote: endnotes.length }}
          onApply={(next, scope) => {
            setShowFootnoteDialog(false)
            armPageSetupUndo()
            applyNoteSettings(next, scope)
          }}
          onInsert={(kind, customMark) => {
            setShowFootnoteDialog(false)
            setNotePrompt(customMark ? { kind, customMark } : { kind })
          }}
          onConvert={(mode) => convertNotesImpl(reviewCtxRef.current, mode)}
          onClose={() => {
            setShowFootnoteDialog(false)
            editor?.commands.focus()
          }}
        />
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
          canvasTop={canvasTop}
          layout={layoutNow}
          layoutRev={layoutRev}
          sections={sections}
          mirrorMargins={mirrorMargins}
          delSectBreaks={delSectBreaks}
          hfParts={doc.parsed.hfParts ?? {}}
          legacyBands={legacyBands}
          colFlow={viewMode === 'print' ? colFlow : null}
          colMode={viewMode === 'print' ? colMode : 'none'}
          hf={{
            header: hfPhantom ? hfWithPhantom(header, hfPhantom, doc.parsed.headerImages) : header,
            footer: hfPhantom ? hfFooterLine(footer, hfPhantom) : footer,
            ...hfVariants,
            ...(hfPhantom
              ? {
                  footerFirst: hfFooterLine(hfVariants.footerFirst, hfPhantom),
                  footerEven: hfFooterLine(hfVariants.footerEven, hfPhantom),
                  headerFirst: hfWithPhantom(
                    hfVariants.headerFirst,
                    hfPhantom,
                    doc.parsed.headerFirst?.images,
                  ),
                  headerEven: hfWithPhantom(
                    hfVariants.headerEven,
                    hfPhantom,
                    doc.parsed.headerEven?.images,
                  ),
                }
              : {}),
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
          watermarkDirty={watermarkDirty}
          pendingWatermark={pendingWatermarkImage}
          blockMetaOf={blockMetaOf}
          footnotesBeneathText={noteSettings.footnote.pos === 'beneathText'}
          endnoteItems={endnoteItems}
          resolveHf={(secs, si, kind, variant) =>
            hfLayoutResolved(
              kind,
              resolveHf({ ...hfState, sections: secs }, si, kind, variant),
              hfPhantom,
            )
          }
          comments={commentsUi.cards ? comments : []}
          anchorBlocks={doc.parsed.blocks}
          revisionView={balloonsActive && balloonMode !== 'inline' ? editor?.view : null}
          balloonKinds={balloonRevKinds(balloonMode)}
          clearPageGaps={() => {
            // column-layout decorations stay: the preview measures with the block widths
            // (line boxes must keep column wrapping); transforms are neutralized by its
            // measuring-columns state. Row-fill heights stay too: a split declared-height
            // row keeps its stretched height as real layout for the preview measure.
            if (editor) {
              setPageGaps(editor.view, [])
              syncNestedTableGaps(editor.view, [])
            }
            const wrap = document.querySelector('.editor-scroll .page-wrap')
            if (wrap) {
              syncCutOverlays(wrap as HTMLElement, [], 1)
              syncPageBorders(wrap as HTMLElement, null, 1)
              syncLineNumbers(wrap as HTMLElement, wrap as HTMLElement, [], [], 1)
              syncColumnRules(wrap as HTMLElement, wrap as HTMLElement, [], [], 1)
              syncPageSheets(wrap as HTMLElement, 1, null)
              clearMarginAnnotations(wrap as HTMLElement)
              clearFloatShifts(wrap as HTMLElement)
            }
          }}
          onExportPdf={() => void exportPdf()}
          onClose={() => {
            if (!exportLocksPreview) setShowPagePreview(false)
          }}
          suppressEscape={showPrintDialog || exportLocksPreview}
          behindPrintDialog={showPrintDialog}
        />
      )}

      {doc && showPrintDialog && (
        <PrintDialog editor={editor} onClose={closePrintDialog} setStatus={setStatus} />
      )}

      {showSpelling && doc && editor && (
        <SpellingDialog editor={editor} onClose={() => setShowSpelling(false)} />
      )}

      {stats && (
        <WordCountDialog stats={stats.body} notes={stats.notes} onClose={() => setStats(null)} />
      )}

      {showAutoCorrect && editor && (
        <AutoCorrectDialog
          onClose={() => setShowAutoCorrect(false)}
          selection={selectionClip(editor)}
          autoText={{
            entries: autoTextList,
            suggestedName: selectionSnippet(editor),
            onAdd: (name) => {
              saveAutoText(editor, name)
              setAutoTextList(autoTextEntries())
            },
            onInsert: (entry) => {
              setShowAutoCorrect(false)
              insertAutoText(editor, entry)
            },
            onDelete: (name) => {
              deleteAutoText(name)
              setAutoTextList(autoTextEntries())
            },
          }}
        />
      )}
      {showPreferences && <PreferencesDialog onClose={() => setShowPreferences(false)} />}

      {docPwdPrompt && (
        <PasswordDialog
          title={t('appDocPwdTitle')}
          body={t('appDocPwdBody', { name: docPwdPrompt.name })}
          label={t('appDocPwdLabel')}
          placeholder={t('appDocPwdPlaceholder')}
          value={docPwdPrompt.value}
          error={docPwdPrompt.errorKey ? t(docPwdPrompt.errorKey) : ''}
          busy={docPwdPrompt.busy}
          submitLabel={docPwdPrompt.busy ? t('appStartOpening') : t('appOk')}
          cancelLabel={t('appCancel')}
          onChange={(value) => setDocPwdPrompt({ ...docPwdPrompt, value, errorKey: '' })}
          onSubmit={() => void submitDocPwd()}
          onCancel={cancelDocPwd}
        />
      )}

      {showProtectDialog && (
        <ProtectDialog
          encrypted={doc?.encrypted ?? false}
          writeProtection={writeProtection}
          protection={protection}
          removePersonalInfo={removePersonalInfo}
          onCancel={() => setShowProtectDialog(false)}
          onApply={(result) => void applyProtectDialog(result)}
        />
      )}

      {modifyPwdPrompt && doc && (
        <PasswordDialog
          title={t('appModifyPwdTitle')}
          body={t('appModifyPwdBody', { name: doc.fileName })}
          label={t('appDocPwdLabel')}
          placeholder={t('appModifyPwdPlaceholder')}
          value={modifyPwdPrompt.value}
          error={modifyPwdPrompt.errorKey ? t(modifyPwdPrompt.errorKey) : ''}
          submitLabel={t('appOk')}
          cancelLabel={t('appOpenReadOnly')}
          onChange={(value) => setModifyPwdPrompt({ value, errorKey: '' })}
          onSubmit={() => void submitModifyPwd()}
          onCancel={() => setModifyPwdPrompt(null)}
        />
      )}
      {pgNumModal && (
        <div
          className="modal-backdrop"
          onMouseDown={(e) => e.target === e.currentTarget && setPgNumModal(null)}
        >
          <div className="modal pgnum-dialog" role="dialog" aria-label={t('appPageNumFormatTitle')}>
            <h2>{t('appPageNumFormatTitle')}</h2>
            <label className="pgnum-row">
              {t('appNumberFormat')}
              <Dropdown
                value={pgNumModal.fmt}
                ariaLabel={t('appNumberFormat')}
                options={[
                  { value: 'decimal', label: '1, 2, 3, …' },
                  { value: 'numberInDash', label: '- 1 -, - 2 -, - 3 -, …' },
                  { value: 'lowerLetter', label: 'a, b, c, …' },
                  { value: 'upperLetter', label: 'A, B, C, …' },
                  { value: 'lowerRoman', label: 'i, ii, iii, …' },
                  { value: 'upperRoman', label: 'I, II, III, …' },
                  { value: 'chineseCounting', label: '一, 二, 三, …' },
                ]}
                onPick={(v) => setPgNumModal({ ...pgNumModal, fmt: v })}
              />
            </label>
            <label className="pgnum-check">
              <input
                type="checkbox"
                name="pgnum-chapter"
                checked={pgNumModal.chapStyle !== ''}
                onChange={(e) =>
                  setPgNumModal({ ...pgNumModal, chapStyle: e.target.checked ? '1' : '' })
                }
              />
              {t('appPgNumIncludeChapter')}
            </label>
            <label className={`pgnum-row pgnum-sub${pgNumModal.chapStyle ? '' : ' disabled'}`}>
              {t('appPgNumChapterStyle')}
              <select
                name="pgnum-chapter-style"
                disabled={!pgNumModal.chapStyle}
                value={pgNumModal.chapStyle || '1'}
                onChange={(e) => setPgNumModal({ ...pgNumModal, chapStyle: e.target.value })}
              >
                {[1, 2, 3, 4, 5, 6, 7, 8, 9].map((level) => (
                  <option key={level} value={level}>
                    {t('ribbonStyleHeadingN', { n: level })}
                  </option>
                ))}
              </select>
            </label>
            <label className={`pgnum-row pgnum-sub${pgNumModal.chapStyle ? '' : ' disabled'}`}>
              {t('appPgNumSeparator')}
              <select
                name="pgnum-chapter-sep"
                disabled={!pgNumModal.chapStyle}
                value={pgNumModal.chapSep}
                onChange={(e) => setPgNumModal({ ...pgNumModal, chapSep: e.target.value })}
              >
                <option value="hyphen">-</option>
                <option value="period">.</option>
                <option value="colon">:</option>
                <option value="emDash">—</option>
                <option value="enDash">–</option>
              </select>
            </label>
            <h3 className="pgnum-title">{t('appPgNumbering')}</h3>
            <label className="pgnum-check">
              <input
                type="radio"
                name="pgnum-restart"
                checked={!pgNumModal.restart}
                onChange={() => setPgNumModal({ ...pgNumModal, restart: false })}
              />
              {t('appContinueFromPrev')}
            </label>
            <div className="pgnum-row">
              <label className="pgnum-check">
                <input
                  type="radio"
                  name="pgnum-restart"
                  checked={pgNumModal.restart}
                  onChange={() => setPgNumModal({ ...pgNumModal, restart: true })}
                />
                {t('appStartAt')}
              </label>
              <input
                type="number"
                min={0}
                aria-label={t('appStartAt')}
                value={pgNumModal.start}
                onChange={(e) =>
                  setPgNumModal({ ...pgNumModal, start: e.target.value, restart: true })
                }
              />
            </div>
            {sections.length > 1 && (
              <p className="pgnum-hint">
                {t('appPgNumSectionN', { n: Math.min(activeSection, sections.length - 1) + 1 })}
              </p>
            )}
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
