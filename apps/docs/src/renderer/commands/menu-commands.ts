import { updateLabels } from '../mail-merge/label-sheet'
import { insertHorizontalLine } from '../editor/horizontal-line'
import type { Editor } from '@tiptap/core'
import { clampDropCapLines, setDropCapFont } from '../editor/drop-cap'
import { replaceSelectedPictureBytes } from '../editor/picture-bytes'
import { setSelectedGraphicFill } from '../editor/graphics-fill'
import { insertLibraryIcons, insertPreparedIcons } from '../editor/insert-icons'
import { applyTextFxPreset } from '../editor/text-fx'
import { setSelectedBoxDirection, textboxDirectionOf } from '../editor/textbox-direction'
import {
  autoCropSelectedPicture,
  selectedPictureAlt,
  selectedPictureGeometry,
} from '../editor/picture-props'
import type { Command } from '@tiptap/pm/state'
import { deleteRow, deleteTable, mergeCells } from '@tiptap/pm/tables'
import { NONE, type ActionCommandId } from '../../shared/commands'
import { PARAGRAPH_SPACINGS } from './paragraph-spacing'
import type { MergeOutputMode } from '../../shared/ipc'
import type {
  Block,
  CitationStyleId,
  CustomNumberingLevel,
  DocGrid,
  HeaderFooter,
  LineNumbering,
  SectionInfo,
  SectionSettings,
  SourceInfo,
  StyleUpsert,
  ThemeColors,
  ThemeFonts,
} from '@genoffice/docx-engine'
import { CITATION_STYLES, DEFAULT_LINE_PITCH_TWIPS } from '@genoffice/docx-engine'
import { WORDART_PRESETS } from '@genoffice/ui'
import { runUiOps } from '../ai/ops'
import {
  TEXTBOX_HEIGHT_EMU,
  TEXTBOX_WIDTH_EMU,
  focusTextboxEditorAt,
  insertShapeAt,
  insertTableAt,
  insertTextboxAt,
  insertWordArtAt,
} from '../components/ribbon-tabs'
import { paraAttrsOf } from '../components/ribbon-format-state'
import { CHART_TYPES, type ChartType } from '../components/chart-types'
import { MERGE_RULES, PLAIN_RULES, ruleNeedsList, type MergeRule } from '../mail-merge/rules'
import type { ListDialogKind } from '../components/ListDialogs'
import { marginsFitPage, type PageMargins } from '../components/MarginDialog'
import { GRID_TYPES, MIN_LINE_PITCH_TWIPS, gutterMoved } from '../components/DocumentDialog'
import { CHAR_GRIDS, charSpaceInRange } from '../char-grid'
import { sectionVertical } from '../pagination-sections'
import { SECTION_VERTICAL, textDirectionTarget } from './text-direction'
import type { HfAction } from '../components/ribbon-hf-tab'
import type { ViewMode } from '../components/ribbon-tabs'
import type { TableDialogKind } from '../components/TableDialogs'
import { getActiveSubEditor } from '../editor/active-editor'
import { repeatable, runRecorded } from './repeat'
import type { EqLayoutKind } from '../editor/eq-layout'
import { fontSymbolContent, parseFontSymbolPayload } from '../editor/font-symbol'
import { ALIGN_FRAMES, ALIGN_MODES, alignSelectedObjects } from '../editor/object-align'
import { groupSelectedObjects, ungroupSelectedObjects } from '../editor/object-group'
import { nextNoteRefPos, noteRefNear } from '../editor/note-nav'
import { applyCase, type CaseMode } from '../editor/case-transform'
import { COVER_PRESETS, insertCoverPage, removeCoverPage } from '../editor/cover-pages'
import { setParagraphDirection, setSelectionAlign } from '../editor/direction'
import {
  bringForward,
  bringToFront,
  sendBackward,
  sendToBack,
  setFloatingWrap,
} from '../editor/floating-z-order'
import { hfHasPageField, hfPageNumberOnly, hfWithoutPageMarks } from '../editor/hf-dom'
import type { InkTool } from '../editor/ink'
import { indexOptionsOf } from '../editor/index-tables'
import { updateTocField } from '../editor/toc-update'
import { insertCitationField, restyleCitations } from '../editor/source-citations'
import { stepParagraphIndent } from '../editor/indent'
import { insertColumnBreak, insertTextWrappingBreak } from '../editor/page-break'
import { pasteFromClipboard } from '../editor/paste-actions'
import type { RevisionScope } from '../editor/revision-filters'
import type { BalloonMode, RevisionDisplayMode } from '../editor/revision-view'
import type { PasteMode } from '../editor/paste-options'
import { startShapeDrawMode } from '../editor/shape-draw'
import { fontFamilyMarkAttrs } from '../font-list'
import { rememberListPreset, type ListPresetKind } from '../list-presets'
import {
  hasIndexEntries,
  indexAtSelection,
  indexColumnsOf,
  indexesBefore,
  insertBibliography,
  insertIndex,
  insertTableOfAuthorities,
  tableAtSelection,
  tableSpecOfPayload,
  insertTableOfFigures,
  insertToc,
  updateIndex,
  updateTableOfAuthorities,
  updateTableOfFigures,
} from './reference-commands'
import {
  deleteTableColumn,
  distributeRowsEvenly,
  editableOrToast,
  enterSelectedTable,
  insertRowsOrColumns,
  selectTablePart,
  splitTableAtSelection,
  setCellAlignment,
  setCellTextDirection,
} from '../editor/table-ops'
import { formulaPlace } from '../editor/table-formula'
import {
  setTableAutoFit,
  toggleRepeatHeaderRows,
  applyTablePreset,
  tablePresetStyleXml,
  setTableLookOption,
  type TablePreset,
} from '../editor/table-properties'
import { textToTableScope } from '../editor/table-convert'
import { distributeSelectedColumns } from '../editor/table-sizing'
import { parseRunBorderAttr } from '../editor/run-border'
import {
  borderAttrsOf,
  borderSpecOfPayload,
  shadingAttrsOf,
  shadingSpecOfPayload,
} from './borders-shading'

/**
 * What a command needs from the app that hosts the editor. App builds one from
 * its own state and callbacks; a command never reaches into App directly, so
 * the menu, the ribbon and the context menu can come to run the same handlers.
 */
export interface CommandHost {
  editor: Editor | null
  hasDoc: boolean
  /** size (pt) of the default paragraph style: what a character grid's pitch is counted from */
  gridBasePt: number
  newFile(): void
  openFile(): void
  openRecent(path: string): void
  save(saveAs: boolean): void
  /** a new untitled document from a template the user picks */
  newFromTemplate(): void
  /** a .dotx copy of the document */
  saveAsTemplate(): void
  nameWithAi(): void
  exportPdf(): void
  exportHtml(): void
  exportMarkdown(): void
  exportImages(): void
  printDoc(): void
  stepZoom(dir: 1 | -1): void
  /** clamps to the supported zoom range */
  setZoom(pct: number): void
  zoomFit(mode: 'width' | 'page' | 'multi'): void
  toggleAi(): void
  toggleDark(): void
  toggleMarks(): void
  toggleStylesPane(): void
  toggleSelectionPane(): void
  toggleCitationsPane(): void
  /** the pen on, or back to the text caret */
  toggleDrawing(): void
  /** new blank documents start with this document's theme fonts, colours and paragraph spacing */
  setDesignAsDefault(): void
  /** says the selected picture has no border to crop */
  nothingToCrop(): void
  /** selects a chart so its data table shows, or says the document has none */
  selectChartData(): void
  /** A+ / A-: bursts of steps are coalesced into one layout */
  stepFontSize(dir: 1 | -1): void
  clearFormatting(): void
  /** the numbering definition a new list takes: one the body already uses, or a new one */
  listNumId(kind: 'bullet' | 'ordered', reuse: boolean): string | null
  /** writes a numbering definition for custom levels; null when it cannot */
  createListDef(levels: CustomNumberingLevel[]): string | null
  /** the list that follows a named list style; null when the style is unknown or has no levels */
  listStyleNum(styleId: string): { numId: string; kind: 'bullet' | 'ordered' } | null
  /** `card` is the library card the define dialog starts from and changes */
  openListDialog(kind: ListDialogKind, card?: number): void
  /** a new named list style of the document, then shown picked on the List Styles tab */
  addListStyle(name: string, levels: CustomNumberingLevel[]): void
  /** takes a user-made list style out of the document; paragraphs following it become plain */
  deleteListStyle(styleId: string): void
  /** a paragraph style goes on the paragraphs, a character style toggles on the selection */
  applyStyle(styleId: string): void
  /** the document's own heading style of that level on the body paragraph; 0 is its default paragraph style */
  applyHeadingLevel(level: 0 | 1 | 2 | 3): void
  /** opens the AI panel and runs the instruction at once */
  aiPreset(instruction: string): void
  /** walks the misspelled words of the document, Word's Spelling and Grammar dialog */
  checkSpelling(): void
  /** asks the AI for synonyms of the selected word or short phrase */
  thesaurus(): void
  /** has the AI translate the selection, or the whole document, into the translation language */
  translate(scope: 'selection' | 'document'): void
  /** opens the file dialog for the mail merge recipient list */
  pickMergeDataSource(): void
  /** the merge fields show the current recipient's values, or their names again */
  toggleMergePreview(): void
  /** the document has a file of its own, which Rename and Move act on */
  onDisk: boolean
  /** asks for a folder and moves the document's file there */
  moveFile(): void
  /** asks for the condition of a Skip Record If or If…Then…Else rule */
  openRuleDialog(rule: MergeRule): void
  /** a recipient list has been picked */
  hasMergeSource: boolean
  /** flips whether the document asks to be opened read-only */
  toggleReadOnlyRecommended(): void
  openDialog(
    dialog:
      | 'table'
      | 'link'
      | 'equation'
      | 'bookmark'
      | 'crossReference'
      | 'caption'
      | 'document'
      | 'changeCase'
      | 'dateTime'
      | 'dropCap'
      | 'watermark'
      | 'columns'
      | 'symbol'
      | 'pageNumbers'
      | 'bulletsNumbering'
      | 'translationLanguage'
      | 'indexTables'
      | 'hyphenation'
      | 'equationOptions'
      | 'footnote'
      | 'revealFormatting'
      | 'fontSubstitution'
      | 'bordersShading'
      | 'rename'
      | 'number'
      | 'sort'
      | 'textDirection'
      | 'altText'
      | 'lineNumbers'
      | 'icons'
      | 'paragraphSpacing'
      | 'reduceFileSize'
      | 'sizePosition'
      | 'recipients'
      | 'filterRecipients'
      | 'findRecipient'
      | 'properties'
      | 'field'
      | 'font'
      | 'paragraph'
      | 'shortcuts'
      | 'customizeKeyboard'
      | 'autocorrect'
      | 'preferences',
  ): void
  /** the Insert Chart dialog, on the given type when one was picked; an open dialog switches to it */
  openChartDialog(type?: ChartType): void
  openTableDialog(kind: TableDialogKind): void
  /** Paste Special for the editor that takes the paste */
  openPasteSpecial(target: Editor): void
  /** the Tabs dialog, on the paragraph the caret is in */
  openTabsDialog(target: Editor): void
  /** the Text Effects dialog, on the text selected in that editor */
  openTextEffects(target: Editor): void
  /** the Phonetic Guide dialog, on the text or the guide selected in that editor */
  openPhoneticGuide(target: Editor): void
  /** the Enclose Characters or Combine Characters dialog, on what is selected in that editor */
  openEqLayout(target: Editor, kind: EqLayoutKind): void
  /** the Mark Citation dialog, on the text selected in the body */
  openMarkCitation(): void
  /** the Mark Index Entry dialog, on the text selected in the body */
  openMarkIndexEntry(): void
  /** asks for a concordance file and marks its entries as index entries */
  autoMarkIndex(): void
  /**
   * Changes the section the document's `nth` index sits in alone. With `wrap`, an index that shares
   * its section gets one first: a continuous break before and after it, as Word keeps
   * a multi-column index.
   */
  setIndexSection(
    nth: number,
    change: (section: SectionSettings) => SectionSettings,
    wrap: boolean,
  ): void
  /** AutoText: `new` keeps the selection under a name, `insert` lists what is kept */
  openAutoText(mode: 'new' | 'insert'): void
  /** asks for another document and puts its body in at the selection */
  insertFile(): void
  /** the heading a table of authorities gives a category of citations */
  citationHeading(category: number): string
  /** the Language dialog, on the text selected in that editor */
  openLanguage(target: Editor): void
  /** the Fit Text dialog, on the text selected in that editor */
  openFitText(target: Editor): void
  /** two lines in one, or horizontal in vertical, on or off for the text selected in that editor */
  toggleAsianLayout(target: Editor, kind: 'combine' | 'vert'): void
  /** crop or background removal for the picture with this data URL */
  openPictureDialog(kind: PictureDialogKind, dataUrl: string): void
  pickImage(): Promise<{ mime: string; base64: string } | null>
  openFind(tab: 'find' | 'replace' | 'goto'): void
  /** moves to the next or previous match of the search in the Find panel, opening it if closed */
  findStep(dir: 1 | -1): void
  showLayoutTab(): void
  /** the mail merge panel, on the kind of output asked for */
  openMailMerge(mode: MergeOutputMode): void
  /** a new table brings up its Table Design tab, as in Word */
  showTableDesignTab(): void
  /** adds a table style to the document's styles, written out by the next save */
  defineTableStyle(style: StyleUpsert): void
  /** the format tab of the selected picture, text box or shape */
  showObjectFormatTab(): void
  /** `lang` is the language the field's text is marked as */
  insertField(instr: string, lang?: string): void
  /** asks for the note text, then inserts the reference */
  insertNote(kind: 'footnote' | 'endnote'): void
  /** opens the text of an existing note */
  editNote(kind: 'footnote' | 'endnote', id: string): void
  /** tells the user the document has no footnotes or endnotes */
  noNotes(): void
  /** the document as parsed, for commands that read its headings and lists */
  blocks: Block[]
  /** the sources citations and the bibliography draw on */
  sources: SourceInfo[]
  /** the style citations and the bibliography are written in */
  citationStyle: CitationStyleId
  /** keeps the style for the document: its sources part names it on save */
  setCitationStyle(style: CitationStyleId): void
  /** page number of each heading in document order; null while the layout cannot say */
  headingPages(): number[] | null
  /** the page each top-level node at these positions starts on; null while there is no layout */
  pagesOf(positions: number[]): number[] | null
  review: ReviewHost
  design: DesignHost
  view: ViewHost
  /** the header and footer the Insert tab's galleries rewrite as a whole */
  headerFooter: HeaderFooterHost
  /** what the Header & Footer tab does to the strip being edited */
  headerFooterAction(action: HfAction): void
  setInkTool(tool: InkTool): void
  clearInk(): void
  /** page setup of the section the caret is in; null with no document */
  section: SectionSettings | null
  setSection(next: SectionSettings): void
  /**
   * A page setup change beyond the caret's section: on every section, or from
   * the caret on (a section break of `breakType` goes in after its paragraph
   * and the change lands on what follows).
   */
  setSectionsIn(
    scope: Exclude<SectionScope, 'section'>,
    change: (section: SectionSettings) => SectionSettings,
    breakType: SectionInfo['startType'],
  ): void
  setMirrorMargins(on: boolean): void
  /**
   * The document's gutter position: at the top of the page, or at the binding
   * side. Moves the gutters of the other sections; the caret's section comes
   * through setSection.
   */
  setGutterAtTop(atTop: boolean): void
  /** how the caret's section starts after the one before it */
  setSectionStart(type: SectionInfo['startType']): void
  insertSectionBreak(type: SectionBreakType): void
  insertImage(): void
  insertPageBreak(): void
  insertBlankPage(): void
  newComment(): void
  wordCount(): void
  aiProofread(): void
  toggleTableGridlines(): void
  tableSectionWidthPx: number
  /** tells the user the caret is not in a table */
  notInTable(): void
}

/** the comment and tracked-change actions App owns */
export interface ReviewHost {
  toggleSpellcheck(): void
  /** forgets the words Ignore All set aside, so they are checked again */
  resetIgnoredWords(): void
  /** the comment thread with this id; without one, the thread at the caret */
  deleteComment(id?: string): void
  resolveComment(id?: string): void
  replyToComment(id: string): void
  deleteAllComments(resolvedOnly: boolean): void
  gotoComment(dir: 1 | -1): void
  toggleComments(): void
  toggleTrackChanges(): void
  revisionDisplay: RevisionDisplayMode
  setRevisionDisplay(mode: RevisionDisplayMode): void
  /** shows or hides one kind of markup, or one author's */
  toggleMarkup(kind: 'insDel' | 'format' | { author: string }): void
  setBalloonMode(mode: BalloonMode): void
  openTrackChangesOptions(): void
  openHighlightChanges(): void
  resolveRevisions(action: 'accept' | 'reject', scope: RevisionScope): void
  gotoRevision(dir: 1 | -1): void
  toggleReviewPane(): void
  /** the Compare Documents dialog */
  compare(): void
  /** the same dialog as Combine Documents */
  combine(): void
  protect(): void
}

export interface HeaderFooterHost {
  header: HeaderFooter | null
  footer: HeaderFooter | null
  setHeader(next: HeaderFooter): void
  setFooter(next: HeaderFooter): void
  /**
   * The page number alone on facing pages: at the binding side or the outer
   * edge of each. Odd and even pages get strips of their own for it.
   */
  setFacingPageNumber(kind: 'header' | 'footer', side: 'inside' | 'outside'): void
  /** opens the strip's editor */
  edit(kind: 'header' | 'footer'): void
  openPageNumberFormat(): void
}

/** how the document is shown; none of it changes the document */
export interface ViewHost {
  mode: ViewMode
  setMode(mode: ViewMode): void
  readMode: boolean
  setReadMode(on: boolean): void
  openPagePreview(): void
  openZoomDialog(): void
  toggleRuler(): void
  toggleGrid(): void
  toggleNavigation(): void
  /** the navigation pane on its page thumbnails; closes the pane when it already shows them */
  toggleThumbnails(): void
  toggleSplit(): void
  /** hides everything but the page, or brings the window's panes back */
  toggleFocus(): void
  /** collapses the ribbon to its tab row, or brings the band back */
  toggleRibbon(): void
  /** opens the current document in another window */
  newWindow(): void
  arrangeAll(): void
  /** the detached windows return to the tab strip */
  mergeWindows(): void
}

/** the document-wide looks App owns */
export interface DesignHost {
  setThemeFonts(fonts: ThemeFonts): void
  setThemeColors(colors: ThemeColors): void
  setPageColor(hex: string | null): void
  /** a text watermark; null removes whatever watermark the document has */
  setWatermark(text: string | null): void
  /** the spacing of the document's default paragraph style; space before is none unless given */
  setParagraphSpacing(spacing: { beforeTwips?: number; afterTwips: number; line: number }): void
}

/** Design > Custom Paragraph Spacing: twips before and after, line spacing as a multiple */
export interface CustomParagraphSpacing {
  beforeTwips: number
  afterTwips: number
  line: number
}

/** a Themes gallery pick: either half may be left out */
export interface ThemePayload {
  fonts?: ThemeFonts
  colors?: ThemeColors
}

type Handler = (host: CommandHost, payload?: string) => void

/** the numbering definition of the first list item of that kind in the body */
export function firstListNumId(blocks: Block[], kind: 'bullet' | 'ordered'): string | null {
  for (const block of blocks) {
    if (block.type === 'listItem' && block.list?.kind === kind) return block.list.numId
  }
  return null
}

/**
 * The editor a formatting command lands on: the strip / textbox editor that has
 * the focus, else the document. Null while the document is read-only; a greyed
 * control is only the visible half of that fence.
 */
function formatTarget(host: CommandHost): Editor | null {
  if (!host.hasDoc || !host.editor?.isEditable) return null
  return getActiveSubEditor() ?? host.editor
}

const toggleMark =
  (mark: 'bold' | 'italic' | 'underline' | 'strike'): Handler =>
  (host) => {
    formatTarget(host)?.chain().focus().toggleMark(mark).run()
  }

const toggleVertAlign =
  (kind: 'subscript' | 'superscript'): Handler =>
  (host) => {
    const target = formatTarget(host)
    if (!target) return
    const on = target.getAttributes('docTextStyle').vertAlign === kind
    target
      .chain()
      .focus()
      .setMark('docTextStyle', { vertAlign: on ? null : kind })
      .run()
  }

const align =
  (value: 'left' | 'center' | 'right' | 'justify' | 'distribute'): Handler =>
  (host) => {
    const target = formatTarget(host)
    if (target) setSelectionAlign(target, value)
  }

const TEXTBOX_INDENT_STEP_TWIPS = 720

const indent =
  (delta: 1 | -1): Handler =>
  (host) => {
    const target = formatTarget(host)
    if (!target || !host.editor) return
    if (target === host.editor) {
      stepParagraphIndent(host.editor, delta)
      return
    }
    // a textbox paragraph carries its own indent; the document goes through styles and the grid
    const current = Number(target.getAttributes('docParagraph').indentLeft ?? 0)
    const indentLeft = Math.max(0, current + delta * TEXTBOX_INDENT_STEP_TWIPS)
    target.chain().focus().updateAttributes('docParagraph', { indentLeft }).run()
  }

export { NONE }

const textStyle =
  (patch: (payload: string) => Record<string, unknown> | null): Handler =>
  (host, payload) => {
    const target = formatTarget(host)
    const attrs = payload === undefined ? null : patch(payload)
    if (target && attrs) target.chain().focus().setMark('docTextStyle', attrs).run()
  }

const paragraphAttrs =
  (
    patch: (payload: string | undefined, host: CommandHost) => Record<string, unknown> | null,
  ): Handler =>
  (host, payload) => {
    const target = formatTarget(host)
    const attrs = target && patch(payload, host)
    if (!target || !attrs) return
    runUiOps(target, [{ op: 'setParagraphAttrs', target: { scope: 'selection' }, attrs }], {
      focus: true,
    })
  }

const SPACE_TOGGLE_TWIPS = 240

/** Word's menu item: 12 pt when the paragraph has none, none when it has some */
const toggleSpace = (side: 'spaceBefore' | 'spaceAfter'): Handler =>
  paragraphAttrs((_, host) => {
    const current = host.editor ? Number(paraAttrsOf(host.editor)[side]) || 0 : 0
    return { [side]: current > 0 ? 0 : SPACE_TOGGLE_TWIPS, [`${side}Auto`]: null }
  })

const CASE_MODES: readonly CaseMode[] = [
  'sentence',
  'lower',
  'upper',
  'title',
  'toggle',
  'halfWidth',
  'fullWidth',
]
const PASTE_MODES: readonly PasteMode[] = ['source', 'merge', 'text']

const toggleList =
  (kind: 'bullet' | 'ordered'): Handler =>
  (host) => {
    const target = formatTarget(host)
    if (!target || !host.editor) return
    const chain = () => target.chain().focus()
    if (target !== host.editor) {
      // a text box paragraph is a list item through its listMarker attr; Enter carries it on
      const marker = target.getAttributes('docParagraph').listMarker as
        { kind?: string } | undefined
      if (marker?.kind === kind) {
        chain().updateAttributes('docParagraph', { listMarker: null }).run()
        return
      }
      const numId = host.listNumId(kind, kind === 'bullet')
      if (!numId) return
      const text = kind === 'bullet' ? '\u2022' : '1.'
      chain()
        .updateAttributes('docParagraph', { listMarker: { text, kind, numId, ilvl: 0 } })
        .run()
      return
    }
    if (target.isActive('docListItem', { kind })) {
      chain().setNode('docParagraph').run()
      return
    }
    chain()
      .setNode('docListItem', { kind, numId: host.listNumId(kind, true), ilvl: 0 })
      .run()
  }

/** list galleries act on body paragraphs only; a text box keeps its own marker */
function bodyTarget(host: CommandHost): Editor | null {
  const target = formatTarget(host)
  return target && target === host.editor ? target : null
}

const LIST_DIALOGS: readonly ListDialogKind[] = [
  'bullet',
  'number',
  'multi',
  'value',
  'indents',
  'style',
]

/** what a gallery card sends: the levels to define, and the recents list that remembers the pick */
export interface ListPresetPayload {
  levels: CustomNumberingLevel[]
  recent?: ListPresetKind
}

const applyListPreset: Handler = (host, payload) => {
  const target = bodyTarget(host)
  if (!target || !payload) return
  const { levels, recent } = JSON.parse(payload) as ListPresetPayload
  const numId = host.createListDef(levels)
  if (!numId) return
  if (recent) rememberListPreset(recent, levels)
  const kind = levels[0]?.numFmt === 'bullet' ? 'bullet' : 'ordered'
  const ilvl = target.isActive('docListItem')
    ? Number(target.getAttributes('docListItem').ilvl) || 0
    : 0
  target.chain().focus().setNode('docListItem', { kind, numId, ilvl }).run()
}

const direction =
  (dir: 'ltr' | 'rtl'): Handler =>
  (host) => {
    if (formatTarget(host) === host.editor && host.editor) setParagraphDirection(host.editor, dir)
  }

/** an Insert command: the body editor of an open, editable document */
const inserting =
  (run: (editor: Editor, payload: string | undefined, host: CommandHost) => void): Handler =>
  (host, payload) => {
    const { editor } = host
    if (editor && host.hasDoc && editor.isEditable) run(editor, payload, host)
  }

const DROP_CAP_LINES = 3

export type PictureDialogKind = 'crop' | 'cutout'
const PICTURE_DIALOGS: readonly PictureDialogKind[] = ['crop', 'cutout']

const Z_ORDER_STEPS: Record<string, (editor: Editor) => void> = {
  front: bringToFront,
  forward: bringForward,
  backward: sendBackward,
  back: sendToBack,
}

export type SectionBreakType = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage'
const SECTION_BREAKS: readonly SectionBreakType[] = [
  'nextPage',
  'continuous',
  'evenPage',
  'oddPage',
]

/** a Margins gallery pick; `mirror` is set only by a preset that states it */
export type PageMarginsPayload = PageMargins & { mirror?: boolean }
/** what a dialog's Apply to picks: the caret's section, every section, or the caret onward */
export type SectionScope = 'section' | 'whole' | 'forward'

/** what the Document dialog changed; one payload, so the section is written once */
export interface DocumentSettingsPayload {
  margins?: PageMargins
  mirror?: boolean
  headerDist?: number
  footerDist?: number
  /** the binding gutter, twips; the margins that come with it already hold it */
  gutter?: number
  /** the document's gutter position; comes with the margins, which hold the gutter on that side */
  gutterAtTop?: boolean
  startType?: SectionInfo['startType']
  vAlign?: NonNullable<SectionSettings['vAlign']>
  columns?: number
  /** the document grid: none, lines only, or lines and characters */
  gridType?: DocGrid['type']
  /** the grid's line pitch, twips */
  linePitch?: number
  /** the grid's character pitch as w:charSpace: its distance from the default font size, 1/4096 pt */
  charSpace?: number
  /** the section's text flow: lrTb for horizontal, or a vertical w:textDirection value */
  textDirection?: string
  /** w:noEndnote: the section hands its endnotes to the next section's end */
  suppressEndnotes?: boolean
  scope?: SectionScope
}

/** Format > Columns: the caret's section as the dialog leaves it */
export interface ColumnsPayload {
  columns: number
  /** unequal columns, twips; absent for equal ones */
  colWidths?: number[]
  /** the gap after each unequal column but the last, twips; absent for the section's own gap */
  colSpaces?: number[]
  /** the gap between equal columns, twips */
  colSpace?: number
  lineBetween?: boolean
  /** with This point forward: what follows the caret starts in the next column */
  startNewColumn?: boolean
  scope?: SectionScope
}

const validLengths = (list: number[] | undefined, n: number, min: number): list is number[] =>
  list?.length === n && list.every((w) => Number.isFinite(w) && w >= min)

/** the columns of a Columns payload on a section; unequal widths keep their proportions on another page width */
function withColumns(
  section: SectionSettings,
  { columns, colWidths, colSpaces, colSpace, lineBetween }: ColumnsPayload,
  from: SectionSettings,
): SectionSettings {
  const unequal = columns > 1 && validLengths(colWidths, columns, 1)
  const gap = colSpace !== undefined && Number.isFinite(colSpace) && colSpace >= 0 ? colSpace : null
  const spaces = unequal && validLengths(colSpaces, columns - 1, 0) ? colSpaces : null
  const gaps = (s: SectionSettings): number =>
    spaces ? spaces.reduce((sum, w) => sum + w, 0) : (gap ?? s.colSpace ?? 720) * (columns - 1)
  const textWidthOf = (s: SectionSettings): number =>
    s.pageWidth - s.marginLeft - s.marginRight - gaps(s)
  const scale = textWidthOf(section) / textWidthOf(from)
  const {
    colSep: _sep,
    colSpaces: _spaces,
    ...next
  } = unequal
    ? {
        ...section,
        columns,
        colWidths: colWidths.map((w) => Math.round(scale > 0 ? w * scale : w)),
        colEqualWidth: false,
      }
    : inEqualColumns(section, columns)
  return {
    ...next,
    ...(spaces && { colSpaces: spaces }),
    ...(gap !== null && { colSpace: gap }),
    ...(lineBetween && columns > 1 && { colSep: true }),
  }
}

/** a section as the Document dialog leaves it; null when nothing of the section changes or the margins do not fit it */
function documentSection(
  section: SectionSettings,
  {
    margins,
    headerDist,
    footerDist,
    gutter,
    gutterAtTop,
    vAlign,
    columns,
    gridType,
    linePitch,
    charSpace,
    textDirection,
    suppressEndnotes,
  }: DocumentSettingsPayload,
  basePt: number,
): SectionSettings | null {
  if (margins && !marginsFitPage(margins, section.pageWidth, section.pageHeight)) return null
  const cols = columns !== undefined && Number.isInteger(columns) && columns > 0 ? columns : null
  const moved = gutterAtTop === undefined ? section : gutterMoved(section, gutterAtTop)
  const grid = gridType !== undefined && GRID_TYPES.includes(gridType) ? gridType : null
  const pitch =
    linePitch !== undefined && Number.isInteger(linePitch) && linePitch >= MIN_LINE_PITCH_TWIPS
      ? linePitch
      : null
  const textWidthPt =
    (section.pageWidth -
      (margins?.left ?? section.marginLeft) -
      (margins?.right ?? section.marginRight)) /
    20
  const space =
    charSpace !== undefined &&
    CHAR_GRIDS.includes(grid ?? section.docGrid?.type ?? 'default') &&
    charSpaceInRange(charSpace, textWidthPt, basePt)
      ? charSpace
      : null
  // only the flows the page lays out sideways are set; anything else leaves the section's own
  const flow =
    textDirection === 'lrTb'
      ? undefined
      : sectionVertical({ ...section, textDirection })
        ? textDirection
        : section.textDirection
  const noEndnote = suppressEndnotes ?? !!section.noEndnote
  if (
    !margins &&
    headerDist === undefined &&
    footerDist === undefined &&
    vAlign === undefined &&
    cols === null &&
    grid === null &&
    pitch === null &&
    space === null &&
    flow === section.textDirection &&
    noEndnote === !!section.noEndnote &&
    moved === section
  )
    return null
  return {
    // another number of columns cannot keep a width list made for the old number
    ...(cols !== null && cols !== moved.columns ? inEqualColumns(moved, cols) : moved),
    // a margin the user sets is an ordinary one, not the file's header-proof fixed margin
    ...(margins && {
      marginTop: margins.top,
      marginRight: margins.right,
      marginBottom: margins.bottom,
      marginLeft: margins.left,
      marginTopFixed: undefined,
      marginBottomFixed: undefined,
    }),
    ...(headerDist !== undefined && { headerDist }),
    ...(footerDist !== undefined && { footerDist }),
    ...(margins &&
      gutter !== undefined && {
        gutter,
        gutterAtTop:
          gutterAtTop === undefined
            ? section.gutterAtTop
            : (gutter > 0 && gutterAtTop) || undefined,
      }),
    ...(vAlign !== undefined && { vAlign: vAlign === 'top' ? undefined : vAlign }),
    // a grid needs a pitch to snap lines to: the section's own, or the one Word starts with
    ...((grid !== null || pitch !== null || space !== null) && {
      docGrid: {
        ...section.docGrid,
        type: grid ?? section.docGrid?.type ?? 'default',
        linePitch: pitch ?? section.docGrid?.linePitch ?? DEFAULT_LINE_PITCH_TWIPS,
        ...(space !== null && { charSpace: space }),
      },
    }),
    textDirection: flow,
    noEndnote: noEndnote || undefined,
  }
}

const SECTION_STARTS: ReadonlyArray<SectionInfo['startType']> = [
  'nextPage',
  'continuous',
  'evenPage',
  'oddPage',
  'nextColumn',
]

/** a section in equal columns: one that had its own widths gives them up */
function inEqualColumns(section: SectionSettings, columns: number): SectionSettings {
  const { colWidths: _widths, colSpaces: _spaces, colEqualWidth: _equal, ...rest } = section
  return {
    ...rest,
    columns,
    ...(section.colEqualWidth === false ? { colEqualWidth: true } : {}),
  }
}

/** a page setup command: changes the caret's section of an open, editable document */
const pageSetup =
  (
    change: (
      section: SectionSettings,
      payload: string,
      host: CommandHost,
    ) => SectionSettings | null,
  ): Handler =>
  (host, payload) => {
    const { section } = host
    if (!section || payload === undefined || !formatTarget(host)) return
    const next = change(section, payload, host)
    if (next) host.setSection(next)
  }

/** a Review command: needs an open document, not an editable one */
const reviewing =
  (run: (review: ReviewHost, payload: string | undefined) => void): Handler =>
  (host, payload) => {
    if (host.hasDoc) run(host.review, payload)
  }

const REVISION_DISPLAYS: readonly RevisionDisplayMode[] = ['simple', 'all', 'none', 'original']
const BALLOON_MODES: readonly BalloonMode[] = ['revisions', 'inline', 'comments']
const AUTHOR_PREFIX = 'author:'

/** payload: a scope name, or the JSON of { author } */
const resolveRevisions = (action: 'accept' | 'reject'): Handler =>
  reviewing((review, payload) => {
    if (!payload) return
    const scope: RevisionScope = payload.startsWith('{')
      ? (JSON.parse(payload) as { author: string })
      : (payload as RevisionScope)
    review.resolveRevisions(action, scope)
  })

/** a Design command: changes how the whole document looks */
const designing =
  (run: (design: DesignHost, payload: string) => void): Handler =>
  (host, payload) => {
    if (payload !== undefined && formatTarget(host)) run(host.design, payload)
  }

const applyTheme: Handler = designing((design, payload) => {
  const { fonts, colors } = JSON.parse(payload) as ThemePayload
  if (fonts) design.setThemeFonts(fonts)
  if (colors) design.setThemeColors(colors)
})

/** a View command: needs an open document */
const viewing =
  (run: (view: ViewHost, payload: string | undefined) => void): Handler =>
  (host, payload) => {
    if (host.hasDoc) run(host.view, payload)
  }

const chartOf =
  (type: ChartType): Handler =>
  (host) => {
    if (host.editor && host.hasDoc) host.openChartDialog(type)
  }

/** the recipient dialogs work on a list: without one, the list is asked for first */
const recipients =
  (dialog: 'recipients' | 'filterRecipients' | 'findRecipient'): Handler =>
  (host) => {
    if (!host.hasDoc) return
    if (host.hasMergeSource) host.openDialog(dialog)
    else host.pickMergeDataSource()
  }

const VIEW_MODES: readonly ViewMode[] = ['print', 'web', 'outline', 'draft', 'story']

/** Print Layout also leaves Read Mode; the other layouts toggle against Print Layout */
function showLayout(view: ViewHost, mode: ViewMode): void {
  if (mode === 'print') {
    view.setMode('print')
    view.setReadMode(false)
  } else view.setMode(view.mode === mode ? 'print' : mode)
}

const INK_TOOLS: readonly InkTool[] = ['select', 'pen', 'highlighter', 'pencil', 'eraser']

const HF_KINDS = ['header', 'footer'] as const
const PAGE_NUMBER_ALIGNS = ['left', 'center', 'right', 'inside', 'outside'] as const

/** an Insert > Header / Footer / Page Number pick on an open, editable document */
const headerFooter =
  (run: (hf: HeaderFooterHost, payload: string | undefined) => void): Handler =>
  (host, payload) => {
    if (formatTarget(host)) run(host.headerFooter, payload)
  }

/** the display shows tracked changes as markup rather than as the finished or the original text */
export const markupShown = (mode: RevisionDisplayMode): boolean =>
  mode === 'all' || mode === 'simple'

/** the box Word's Character Border button writes: a half-point single line */
export const CHARACTER_BORDER = JSON.stringify({ val: 'single', sz: 4 })

/** the fill Word's Character Shading button writes (a 15% grey) */
export const CHARACTER_SHADING = 'D9D9D9'

const LINE_NUMBER_RESTARTS = ['continuous', 'newPage', 'newSection'] as const

const tableDialog =
  (kind: TableDialogKind): Handler =>
  (host) => {
    const { editor } = host
    if (!editor || !host.hasDoc || !editableOrToast(editor)) return
    if (!enterSelectedTable(editor.state, editor.view.dispatch)) {
      host.notInTable()
      return
    }
    host.openTableDialog(kind)
  }

/**
 * A table command that is a plain editor command. On a table selected as a whole it acts on
 * every cell, as Word does. Selecting is fine on a read-only document.
 */
const tableCommand =
  (
    run: (editor: Editor, widthPx: number) => Command | void,
    { readOnlyOk = false } = {},
  ): Handler =>
  (host) => {
    const { editor } = host
    if (!editor || !host.hasDoc) return
    if (!readOnlyOk && !editableOrToast(editor)) return
    if (!enterSelectedTable(editor.state, editor.view.dispatch)) {
      host.notInTable()
      return
    }
    editor.view.focus()
    const command = run(editor, host.tableSectionWidthPx)
    if (command) command(editor.state, editor.view.dispatch)
  }

/** a table command whose pick rides as the payload */
const tablePick =
  (make: (payload: string, editor: Editor) => Command | void | null): Handler =>
  (host, payload) => {
    if (payload === undefined) return
    tableCommand((editor) => make(payload, editor) ?? undefined)(host)
  }

const CELL_V_ALIGNS = ['top', 'center', 'bottom'] as const
const CELL_H_ALIGNS = ['left', 'center', 'right'] as const
const CELL_TEXT_DIRECTIONS = ['lrTb', 'tbRl', 'btLr'] as const
const TABLE_LOOK_OPTIONS = [
  'firstRow',
  'lastRow',
  'bandedRows',
  'firstColumn',
  'lastColumn',
  'bandedColumns',
] as const

const select = (part: 'table' | 'column' | 'row' | 'cell'): Handler =>
  tableCommand(() => selectTablePart(part), { readOnlyOk: true })

export const MENU_COMMAND_HANDLERS: Record<ActionCommandId, Handler> = {
  new: (host) => host.newFile(),
  open: (host) => host.openFile(),
  'open-path': (host, payload) => {
    if (payload) host.openRecent(payload)
  },
  save: (host) => host.save(false),
  'save-as': (host) => host.save(true),
  'file-new-from-template': (host) => host.newFromTemplate(),
  'file-save-as-template': (host) => {
    if (host.hasDoc) host.saveAsTemplate()
  },
  'name-with-ai': (host) => host.nameWithAi(),
  undo: (host) => {
    host.editor?.chain().focus().undo().run()
  },
  redo: (host) => {
    host.editor?.chain().focus().redo().run()
  },
  // Word's one item: Redo while there is something to redo, else the last command again
  repeat: (host) => {
    if (host.editor?.can().redo()) {
      host.editor.chain().focus().redo().run()
      return
    }
    const again = repeatable(repeatEditors(host))
    if (again) runMenuCommand(host, again.command, again.payload)
  },
  'zoom-in': (host) => host.stepZoom(1),
  'zoom-out': (host) => host.stepZoom(-1),
  'zoom-100': (host) => host.setZoom(100),
  'zoom-set': (host, payload) => {
    if (payload) host.setZoom(Number(payload))
  },
  'zoom-page-width': (host) => host.zoomFit('width'),
  'zoom-whole-page': (host) => host.zoomFit('page'),
  'zoom-multi-page': (host) => host.zoomFit('multi'),
  'view-mode': viewing((view, payload) => {
    const mode = VIEW_MODES.find((m) => m === payload)
    if (mode) showLayout(view, mode)
  }),
  'view-print-layout': viewing((view) => showLayout(view, 'print')),
  'view-web-layout': viewing((view) => showLayout(view, 'web')),
  'view-outline': viewing((view) => showLayout(view, 'outline')),
  'view-draft': viewing((view) => showLayout(view, 'draft')),
  'view-header-footer': headerFooter((hf) => hf.edit('header')),
  'read-mode': viewing((view) => view.setReadMode(!view.readMode)),
  'page-preview': viewing((view) => view.openPagePreview()),
  'zoom-dialog': viewing((view) => view.openZoomDialog()),
  'toggle-ruler': viewing((view) => view.toggleRuler()),
  'toggle-grid': viewing((view) => view.toggleGrid()),
  'toggle-navigation': viewing((view) => view.toggleNavigation()),
  'view-thumbnails': viewing((view) => view.toggleThumbnails()),
  'view-reveal-formatting': (host) => {
    if (host.hasDoc && host.editor) host.openDialog('revealFormatting')
  },
  'split-view': viewing((view) => view.toggleSplit()),
  'view-ribbon': (host) => host.view.toggleRibbon(),
  'view-focus': viewing((view) => view.toggleFocus()),
  // Word moves to the notes; here a note's text opens in its own box
  'view-footnotes': (host) => {
    const { editor } = host
    if (!editor || !host.hasDoc) return
    const note = noteRefNear(editor.state.doc, editor.state.selection.from)
    if (note) host.editNote(note.kind, note.id)
    else host.noNotes()
  },
  'new-window': (host) => host.view.newWindow(),
  'view-arrange-all': viewing((view) => view.arrangeAll()),
  'view-merge-windows': viewing((view) => view.mergeWindows()),
  'edit-header-footer': headerFooter((hf, payload) => {
    const kind = HF_KINDS.find((k) => k === payload)
    if (kind) hf.edit(kind)
  }),
  'remove-header-footer': headerFooter((hf, payload) => {
    if (payload === 'header') hf.setHeader({ text: '' })
    else if (payload === 'footer') hf.setFooter({ text: '' })
  }),
  // payload: header|footer,left|center|right|inside|outside; the strip becomes the page number alone
  'page-number': headerFooter((hf, payload) => {
    const [where, side] = (payload ?? '').split(',')
    const kind = HF_KINDS.find((k) => k === where)
    const align = PAGE_NUMBER_ALIGNS.find((k) => k === side)
    if (!kind || !align) return
    if (align === 'inside' || align === 'outside') hf.setFacingPageNumber(kind, align)
    else if (kind === 'header') hf.setHeader(hfPageNumberOnly(align))
    else hf.setFooter(hfPageNumberOnly(align))
  }),
  'insert-page-numbers': (host) => {
    if (host.hasDoc && formatTarget(host)) host.openDialog('pageNumbers')
  },
  'page-number-format': headerFooter((hf) => hf.openPageNumberFormat()),
  'page-number-remove': headerFooter((hf) => {
    if (hf.footer && hfHasPageField(hf.footer)) hf.setFooter(hfWithoutPageMarks(hf.footer))
    if (hf.header && hfHasPageField(hf.header)) hf.setHeader(hfWithoutPageMarks(hf.header))
  }),
  // payload: the action as JSON
  'header-footer': (host, payload) => {
    if (payload && host.hasDoc) host.headerFooterAction(JSON.parse(payload) as HfAction)
  },
  // leaving the Draw tab drops back to the text caret even with no document open
  'ink-tool': (host, payload) => {
    const tool = INK_TOOLS.find((k) => k === payload)
    if (tool) host.setInkTool(tool)
  },
  'ink-clear': (host) => {
    if (host.hasDoc) host.clearInk()
  },
  'edit-toggle-drawing': (host) => {
    if (host.hasDoc) host.toggleDrawing()
  },
  'design-set-default': (host) => {
    if (host.hasDoc) host.setDesignAsDefault()
  },
  'toggle-ai': (host) => host.toggleAi(),
  'toggle-dark': (host) => host.toggleDark(),
  // Word semantics: the menu opens the Insert Table dialog (custom rows/cols)
  'insert-table': (host) => {
    if (host.editor && host.hasDoc) host.openDialog('table')
  },
  'insert-image': (host) => {
    if (host.editor && host.hasDoc) host.insertImage()
  },
  'insert-page-break': (host) => {
    if (host.editor && host.hasDoc) host.insertPageBreak()
  },
  'insert-blank-page': (host) => {
    if (host.editor && host.hasDoc) host.insertBlankPage()
  },
  // payload: the chart type the dialog opens on; the ribbon button leaves it to the dialog
  'insert-chart': (host, payload) => {
    if (host.editor && host.hasDoc) host.openChartDialog(CHART_TYPES.find((k) => k === payload))
  },
  'insert-chart-bar': chartOf('barH'),
  'insert-chart-area': chartOf('area'),
  'insert-chart-radar': chartOf('radar'),
  'insert-chart-scatter': chartOf('scatter'),
  'insert-chart-stock': chartOf('stock'),
  'insert-chart-treemap': chartOf('treemap'),
  'insert-chart-sunburst': chartOf('sunburst'),
  'insert-chart-histogram': chartOf('histogram'),
  'insert-chart-pareto': chartOf('pareto'),
  'insert-chart-box-whisker': chartOf('boxWhisker'),
  'insert-chart-waterfall': chartOf('waterfall'),
  'insert-chart-funnel': chartOf('funnel'),
  'insert-chart-surface': chartOf('surface'),
  'insert-bookmark': (host) => {
    if (host.editor && host.hasDoc) host.openDialog('bookmark')
  },
  'insert-cross-reference': (host) => {
    if (host.editor && host.hasDoc) host.openDialog('crossReference')
  },
  'insert-caption': (host) => {
    if (host.editor && host.hasDoc) host.openDialog('caption')
  },
  'insert-cover-page': inserting((editor, id) => {
    const preset = COVER_PRESETS.find((p) => p.id === id)
    if (preset) insertCoverPage(editor, preset)
  }),
  'remove-cover-page': inserting((editor) => removeCoverPage(editor)),
  // payload: rows,columns picked on the grid
  'insert-table-grid': inserting((editor, size, host) => {
    const [rows, cols] = (size ?? '').split(',').map(Number)
    if (!(rows > 0) || !(cols > 0)) return
    insertTableAt(editor, rows, cols)
    host.showTableDesignTab()
  }),
  // arms the crosshair: a click drops the default size, a drag sets it, Esc cancels
  'insert-shape': inserting((editor, prst) => {
    if (prst) startShapeDrawMode(editor, prst, (opts) => insertShapeAt(editor, prst, opts))
  }),
  'insert-text-box': inserting((editor, how) => {
    if (how === 'draw') {
      startShapeDrawMode(editor, 'rect', (opts) => insertTextboxAt(editor, opts), {
        clickSize: { widthEmu: TEXTBOX_WIDTH_EMU, heightEmu: TEXTBOX_HEIGHT_EMU },
        ghost: { fill: 'rgba(255,255,255,0.6)', border: '#000000' },
        onInserted: (pos) => focusTextboxEditorAt(editor, pos),
      })
      return
    }
    const pos = insertTextboxAt(editor)
    if (pos != null) requestAnimationFrame(() => focusTextboxEditorAt(editor, pos))
  }),
  // Word's Draw Vertical Text Box: a click drops a tall box, a drag sets its size
  'insert-vertical-text-box': inserting((editor) => {
    startShapeDrawMode(
      editor,
      'rect',
      (opts) => insertTextboxAt(editor, { ...opts, vert: 'eaVert' }),
      {
        clickSize: { widthEmu: TEXTBOX_HEIGHT_EMU, heightEmu: TEXTBOX_WIDTH_EMU },
        ghost: { fill: 'rgba(255,255,255,0.6)', border: '#000000' },
        onInserted: (pos) => focusTextboxEditorAt(editor, pos),
      },
    )
  }),
  'insert-wordart': inserting((editor, id) => {
    const preset = WORDART_PRESETS.find((p) => p.id === id)
    if (preset) insertWordArtAt(editor, preset)
  }),
  // payload: none, drop or margin, then optionally (comma separated) the lines to drop, the
  // distance from the text in twips and the letter's font
  'drop-cap': inserting((editor, payload) => {
    // a drop cap belongs to a body paragraph: not to the header, footer or text box with the caret
    if (getActiveSubEditor()) return
    const [type, lines, gap, ...font] = (payload ?? '').split(',')
    if (type !== NONE && type !== 'drop' && type !== 'margin') return
    const hSpace = Math.round(Number(gap))
    const dropCap =
      type === NONE
        ? null
        : JSON.stringify({
            type,
            lines: lines === undefined ? DROP_CAP_LINES : clampDropCapLines(lines),
            ...(hSpace > 0 ? { hSpace } : {}),
          })
    runUiOps(
      editor,
      [{ op: 'setParagraphAttrs', target: { scope: 'selection' }, attrs: { dropCap } }],
      { focus: true },
    )
    if (type !== NONE && font.length > 0) setDropCapFont(editor, font.join(','))
  }),
  'insert-field': inserting((_, instr, host) => {
    if (instr) host.insertField(instr)
  }),
  // payload: the language of the date, a comma, the field instruction
  'insert-date-field': inserting((_, payload, host) => {
    const at = (payload ?? '').indexOf(',')
    if (at > 0) host.insertField(payload!.slice(at + 1), payload!.slice(0, at))
  }),
  // text goes where the caret is: a header, footer or text box as well as the body
  'insert-symbol': (host, symbol) => {
    const target = formatTarget(host)
    if (target && symbol) target.chain().focus().insertContent(symbol).run()
  },
  // payload: a symbol font and its character (see fontSymbolPayload); saves as w:sym
  'insert-font-symbol': (host, payload) => {
    const target = formatTarget(host)
    const picked = payload ? parseFontSymbolPayload(payload) : null
    const content = picked && fontSymbolContent(picked.font, picked.char)
    if (target && content) target.chain().focus().insertContent(content).run()
  },
  'insert-advanced-symbol': (host) => {
    if (host.hasDoc && formatTarget(host)) host.openDialog('symbol')
  },
  'footnote-next': (host) => {
    const { editor } = host
    if (!editor || !host.hasDoc) return
    const pos = nextNoteRefPos(editor.state.doc, editor.state.selection.to, 'footnote')
    if (pos !== null) editor.chain().focus().setNodeSelection(pos).scrollIntoView().run()
  },
  'insert-footnote': inserting((_, __, host) => host.insertNote('footnote')),
  'insert-endnote': inserting((_, __, host) => host.insertNote('endnote')),
  'insert-footnote-dialog': inserting((_, __, host) => host.openDialog('footnote')),
  // payload: a table spec from Index and Tables; without one, every heading level
  'insert-toc': inserting((editor, payload, host) =>
    insertToc(editor, host.blocks, host.headingPages, tableSpecOfPayload(payload), host.pagesOf),
  ),
  'update-toc': inserting((editor, _, host) =>
    updateTocField(editor, host.blocks, host.headingPages, undefined, undefined, host.pagesOf),
  ),
  // payload: the source's tag
  'insert-citation': inserting((editor, tag, host) => {
    const source = host.sources.find((s) => s.tag === tag)
    if (!source) return
    insertCitationField(editor, host.blocks, source, host.sources, host.citationStyle)
  }),
  'insert-bibliography': inserting((editor, _, host) =>
    insertBibliography(editor, host.blocks, host.sources, host.citationStyle),
  ),
  // payload: the style's id
  'citation-style': inserting((editor, payload, host) => {
    const style = CITATION_STYLES.find((s) => s.id === payload)?.id
    if (!style) return
    host.setCitationStyle(style)
    restyleCitations(editor, host.blocks, host.sources, style)
  }),
  'insert-index': inserting((editor, payload, host) => {
    const spec = tableSpecOfPayload(payload)
    const before = spec ? tableAtSelection(editor, host.blocks, 'index')?.spec : undefined
    const columns = indexColumnsOf(spec)
    const inColumns = (nth: number) =>
      host.setIndexSection(nth, (section) => inEqualColumns(section, columns), columns > 1)
    // the index the selection is in is the one rebuilt: making its section moves the caret out of it
    const existing = before ? indexAtSelection(editor, host.blocks) : -1
    if (before && existing >= 0) {
      // no entry left to list: the rebuild only reports it, and the section must not change without it
      if (hasIndexEntries(editor) && indexOptionsOf(before).columns !== columns) inColumns(existing)
      insertIndex(editor, host.blocks, host.pagesOf, spec, existing)
      return
    }
    // a new index may not be the document's first, and the caret does not end up in it
    const nth = indexesBefore(editor, host.blocks, editor.state.selection.from)
    if (!insertIndex(editor, host.blocks, host.pagesOf, spec)) return
    // the ribbon's index (no spec) stays where it is put, as before
    if (spec && columns > 1) inColumns(nth)
  }),
  'mark-index-entry': inserting((_, __, host) => host.openMarkIndexEntry()),
  'index-automark': inserting((_, __, host) => host.autoMarkIndex()),
  'update-index': inserting((editor, _, host) => updateIndex(editor, host.blocks, host.pagesOf)),
  // payload: the caption label to list, or a table spec; without one, the first label the document uses
  'insert-table-of-figures': inserting((editor, payload, host) => {
    const spec = tableSpecOfPayload(payload)
    insertTableOfFigures(editor, host.blocks, host.pagesOf, spec ? undefined : payload, spec)
  }),
  'update-table-of-figures': inserting((editor, _, host) =>
    updateTableOfFigures(editor, host.blocks, host.pagesOf),
  ),
  'insert-file': inserting((_, __, host) => host.insertFile()),
  'insert-icons': inserting((_, __, host) => host.openDialog('icons')),
  // the dialog prepares what it picked, so its icons go in at once
  'insert-icon': inserting((editor, payload) => {
    const names = payload?.split(',') ?? []
    if (!insertPreparedIcons(editor, names)) void insertLibraryIcons(editor, names)
  }),
  'insert-horizontal-line': inserting((editor) => {
    insertHorizontalLine(editor)
  }),
  'insert-autotext': inserting((_, __, host) => host.openAutoText('insert')),
  'insert-autotext-new': inserting((_, __, host) => host.openAutoText('new')),
  'toa-mark-citation': inserting((_, __, host) => host.openMarkCitation()),
  // payload: the category number to list, or a table spec; without one, the first category with a citation
  'toa-insert': inserting((editor, payload, host) =>
    insertTableOfAuthorities(
      editor,
      host.blocks,
      host.pagesOf,
      host.citationHeading,
      Number(payload) > 0 ? Number(payload) : undefined,
      tableSpecOfPayload(payload),
    ),
  ),
  'toa-update': inserting((editor, _, host) =>
    updateTableOfAuthorities(editor, host.blocks, host.pagesOf, host.citationHeading),
  ),
  'insert-link': (host) => {
    if (host.editor && host.hasDoc) host.openDialog('link')
  },
  'insert-equation': (host) => {
    if (host.editor && host.hasDoc) host.openDialog('equation')
  },
  'insert-comment': (host) => {
    if (host.editor && host.hasDoc) host.newComment()
  },
  bold: toggleMark('bold'),
  italic: toggleMark('italic'),
  underline: toggleMark('underline'),
  strikethrough: toggleMark('strike'),
  subscript: toggleVertAlign('subscript'),
  superscript: toggleVertAlign('superscript'),
  'indent-decrease': indent(-1),
  'indent-increase': indent(1),
  'show-marks': (host) => host.toggleMarks(),
  'font-family': textStyle(fontFamilyMarkAttrs),
  'font-size': textStyle((pt) =>
    Number(pt) > 0 ? { sizeHalfPoints: Math.round(Number(pt) * 2) } : null,
  ),
  'font-color': textStyle((hex) => ({ color: hex === NONE ? null : hex })),
  highlight: textStyle((name) => ({ highlight: name === NONE ? null : name })),
  // Word's Character Shading: its one grey on, any shading off
  // payload: the fill Borders and Shading picked for the text, or none
  'character-shading': (host, payload) => {
    const target = formatTarget(host)
    if (!target) return
    if (payload !== undefined && payload !== NONE && !/^[0-9A-Fa-f]{6}$/.test(payload)) return
    const { shading, shadingDisplay } = target.getAttributes('docTextStyle')
    const fill =
      payload === undefined
        ? (shadingDisplay ?? shading)
          ? null
          : CHARACTER_SHADING
        : payload === NONE
          ? null
          : payload
    target.chain().focus().setMark('docTextStyle', { shading: fill, shadingDisplay: null }).run()
  },
  // payload: one of Word's built-in paragraph spacings, by id
  // payload: one of the built-ins by id, or a spacing typed by hand as JSON
  'design-paragraph-spacing': designing((design, payload) => {
    if (payload.startsWith('{')) {
      const { beforeTwips, afterTwips, line } = JSON.parse(payload) as CustomParagraphSpacing
      const ok = (n: unknown): n is number => typeof n === 'number' && Number.isFinite(n) && n >= 0
      if (ok(beforeTwips) && ok(afterTwips) && ok(line) && line > 0)
        design.setParagraphSpacing({ beforeTwips, afterTwips, line })
      return
    }
    const spacing = PARAGRAPH_SPACINGS.find((s) => s.id === payload)
    if (spacing) design.setParagraphSpacing(spacing)
  }),
  'design-custom-paragraph-spacing': (host) => {
    if (host.hasDoc && formatTarget(host)) host.openDialog('paragraphSpacing')
  },
  'tools-hyphenation': (host) => {
    if (host.hasDoc && host.editor) host.openDialog('hyphenation')
  },
  // Word's Character Border: its thin box on, any border off (a style's included)
  // payload: the box Borders and Shading picked for the text as the mark's JSON, or none
  'character-border': (host, payload) => {
    const target = formatTarget(host)
    if (!target) return
    if (payload !== undefined && payload !== NONE && !parseRunBorderAttr(payload)) return
    const bdr =
      payload === undefined
        ? target.getAttributes('docTextStyle').bdr
          ? null
          : CHARACTER_BORDER
        : payload === NONE
          ? null
          : payload
    target
      .chain()
      .focus()
      .setMark('docTextStyle', { bdr, bdrOwn: bdr ?? 'none' })
      .run()
  },
  'insert-index-tables': (host) => {
    if (bodyTarget(host)) host.openDialog('indexTables')
  },
  // a preset is a multiple: it clears any at-least / exact rule so it takes effect
  'line-spacing': paragraphAttrs((multiple) =>
    Number(multiple) > 0
      ? { lineSpacing: Number(multiple), lineRule: null, lineRawTwips: null }
      : null,
  ),
  'space-before': toggleSpace('spaceBefore'),
  'space-after': toggleSpace('spaceAfter'),
  // payload: a fill, which takes a pattern off as Word's Shading button does; none; or a
  // fill with its pattern as JSON
  'paragraph-shading': paragraphAttrs((payload) => {
    if (payload === undefined) return null
    const spec = shadingSpecOfPayload(payload) ?? {
      fill: payload === NONE ? null : payload,
      pattern: null,
      color: null,
    }
    return shadingAttrsOf(spec)
  }),
  // payload: the sides that carry a line, any of t, b, l and r, each keeping the look it has;
  // none takes them all off; or the sides with their line style, colour, width and gap as JSON
  'paragraph-borders': paragraphAttrs((payload) => {
    if (payload === undefined) return null
    const spec = borderSpecOfPayload(payload)
    return spec ? borderAttrsOf(spec) : { borders: payload === NONE ? null : payload }
  }),
  'format-borders-shading': (host) => {
    if (formatTarget(host)) host.openDialog('bordersShading')
  },
  'format-font-substitution': (host) => {
    if (host.hasDoc) host.openDialog('fontSubstitution')
  },
  // payload: the rule; a rule with no argument goes in at once, a condition needs a list's fields
  'mailings-update-labels': (host) => {
    if (host.editor && host.hasDoc && host.editor.isEditable) updateLabels(host.editor)
  },
  'mailings-rules': inserting((_, rule, host) => {
    const plain = Object.entries(PLAIN_RULES).find(([name]) => name === rule)?.[1]
    if (plain) return host.insertField(plain)
    const asked = MERGE_RULES.find((r) => r === rule)
    if (!asked) return
    if (host.hasMergeSource || !ruleNeedsList(asked)) host.openRuleDialog(asked)
    else host.pickMergeDataSource()
  }),
  'font-grow': (host) => {
    if (formatTarget(host)) host.stepFontSize(1)
  },
  'font-shrink': (host) => {
    if (formatTarget(host)) host.stepFontSize(-1)
  },
  'format-change-case': (host) => {
    if (formatTarget(host)) host.openDialog('changeCase')
  },
  'format-drop-cap': (host) => {
    if (host.editor && host.hasDoc && host.editor.isEditable && !getActiveSubEditor())
      host.openDialog('dropCap')
  },
  'insert-watermark': (host) => {
    if (host.editor && host.hasDoc && host.editor.isEditable) host.openDialog('watermark')
  },
  'insert-date-time': (host) => {
    if (formatTarget(host)) host.openDialog('dateTime')
  },
  'insert-field-dialog': (host) => {
    if (formatTarget(host)) host.openDialog('field')
  },
  'edit-paste-special': (host) => {
    const target = formatTarget(host)
    if (target) host.openPasteSpecial(target)
  },
  'change-case': (host, payload) => {
    const target = formatTarget(host)
    const mode = CASE_MODES.find((m) => m === payload)
    if (target && mode) applyCase(target, mode)
  },
  'edit-clear-contents': (host) => {
    formatTarget(host)?.chain().focus().deleteSelection().run()
  },
  'clear-formatting': (host) => {
    if (formatTarget(host)) host.clearFormatting()
  },
  bullets: toggleList('bullet'),
  numbering: toggleList('ordered'),
  'list-none': (host) => {
    const target = bodyTarget(host)
    if (target?.isActive('docListItem')) target.chain().focus().setNode('docParagraph').run()
  },
  'list-level': (host, payload) => {
    const target = bodyTarget(host)
    const ilvl = Number(payload)
    if (!target?.isActive('docListItem') || !Number.isInteger(ilvl) || ilvl < 0) return
    target.chain().focus().updateAttributes('docListItem', { ilvl }).run()
  },
  'list-preset': applyListPreset,
  // payload: the list style's id
  'list-style': (host, payload) => {
    const target = bodyTarget(host)
    const list = target && payload ? host.listStyleNum(payload) : null
    if (!target || !list) return
    // setNode gives every paragraph the same attributes; items already in a list keep their own level
    const levels: Array<[pos: number, ilvl: number]> = []
    const { from, to } = target.state.selection
    target.state.doc.nodesBetween(from, to, (node, pos) => {
      const ilvl = node.type.name === 'docListItem' ? Number(node.attrs.ilvl) || 0 : 0
      if (ilvl > 0) levels.push([pos, ilvl])
      return !node.isTextblock
    })
    target
      .chain()
      .focus()
      .setNode('docListItem', { ...list, ilvl: 0 })
      .command(({ tr }) => {
        for (const [pos, ilvl] of levels) tr.setNodeAttribute(pos, 'ilvl', ilvl)
        return true
      })
      .run()
  },
  // lists are the body's: a header, footer or text box has none
  'format-bullets-numbering': (host) => {
    if (bodyTarget(host)) host.openDialog('bulletsNumbering')
  },
  // payload: the dialog, and after "#" the library card it customizes
  'list-dialog': (host, payload) => {
    const [name, card] = (payload ?? '').split('#')
    const kind = LIST_DIALOGS.find((k) => k === name)
    if (kind && host.hasDoc) host.openListDialog(kind, card ? Number(card) : undefined)
  },
  // payload: JSON { name, levels }
  'list-style-add': (host, payload) => {
    if (!payload || !bodyTarget(host)) return
    const { name, levels } = JSON.parse(payload) as { name: string; levels: CustomNumberingLevel[] }
    host.addListStyle(name, levels)
  },
  'list-style-delete': (host, payload) => {
    if (payload && bodyTarget(host)) host.deleteListStyle(payload)
  },
  'apply-style': (host, payload) => {
    if (payload && formatTarget(host)) host.applyStyle(payload)
  },
  'ai-preset': (host, payload) => {
    if (payload && host.hasDoc) host.aiPreset(payload)
  },
  'direction-ltr': direction('ltr'),
  'direction-rtl': direction('rtl'),
  'paste-as': (host, payload) => {
    const target = formatTarget(host)
    if (target)
      void pasteFromClipboard(
        target,
        payload === 'picture' ? 'picture' : PASTE_MODES.find((m) => m === payload),
      )
  },
  'styles-pane': (host) => host.toggleStylesPane(),
  'format-object': (host) => {
    if (host.hasDoc && host.editor) host.showObjectFormatTab()
  },
  'edit-select-data': (host) => {
    if (host.hasDoc && host.editor) host.selectChartData()
  },
  'citations-pane': (host) => {
    if (host.hasDoc && host.editor) host.toggleCitationsPane()
  },
  'format-selection-pane': (host) => {
    if (host.hasDoc && host.editor) host.toggleSelectionPane()
  },
  'align-left': align('left'),
  'align-center': align('center'),
  'align-right': align('right'),
  'align-justify': align('justify'),
  'align-distributed': align('distribute'),
  'toggle-spellcheck': reviewing((review) => review.toggleSpellcheck()),
  'tools-hide-spelling': reviewing((review) => review.toggleSpellcheck()),
  'tools-reset-ignored': reviewing((review) => review.resetIgnoredWords()),
  'delete-comment': reviewing((review, id) => review.deleteComment(id)),
  'comment-resolve': reviewing((review, id) => review.resolveComment(id)),
  'comment-reply': reviewing((review, id) => {
    if (id) review.replyToComment(id)
  }),
  'delete-comments': reviewing((review, which) => review.deleteAllComments(which === 'resolved')),
  'comment-previous': reviewing((review) => review.gotoComment(-1)),
  'comment-next': reviewing((review) => review.gotoComment(1)),
  'show-comments': reviewing((review) => review.toggleComments()),
  'track-changes': reviewing((review) => review.toggleTrackChanges()),
  'revision-display': reviewing((review, payload) => {
    const mode = REVISION_DISPLAYS.find((m) => m === payload)
    if (mode) review.setRevisionDisplay(mode)
  }),
  // Word's View > Markup: every tracked change and comment shown, or none
  'view-markup': reviewing((review) =>
    review.setRevisionDisplay(markupShown(review.revisionDisplay) ? 'none' : 'all'),
  ),
  // payload: insDel, format, or author:<name>
  'markup-filter': reviewing((review, payload) => {
    if (payload === 'insDel' || payload === 'format') review.toggleMarkup(payload)
    else if (payload?.startsWith(AUTHOR_PREFIX))
      review.toggleMarkup({ author: payload.slice(AUTHOR_PREFIX.length) })
  }),
  'balloon-mode': reviewing((review, payload) => {
    const mode = BALLOON_MODES.find((m) => m === payload)
    if (mode) review.setBalloonMode(mode)
  }),
  'track-changes-options': reviewing((review) => review.openTrackChangesOptions()),
  'tools-highlight-changes': reviewing((review) => review.openHighlightChanges()),
  'accept-revision': resolveRevisions('accept'),
  'reject-revision': resolveRevisions('reject'),
  'revision-previous': reviewing((review) => review.gotoRevision(-1)),
  'revision-next': reviewing((review) => review.gotoRevision(1)),
  'review-pane': reviewing((review) => review.toggleReviewPane()),
  'compare-documents': reviewing((review) => review.compare()),
  'tools-combine-documents': reviewing((review) => review.combine()),
  'protect-document': reviewing((review) => review.protect()),
  'font-dialog': (host) => {
    if (host.hasDoc) host.openDialog('font')
  },
  'paragraph-dialog': (host) => {
    if (host.hasDoc) host.openDialog('paragraph')
  },
  'page-setup': (host) => host.showLayoutTab(),
  'mailings-select-recipients': (host) => {
    if (host.hasDoc) host.pickMergeDataSource()
  },
  'mailings-edit-recipients': recipients('recipients'),
  // with no list yet there is nothing to preview: ask for one, as the other list commands do
  'mailings-preview-results': (host) => {
    if (!host.hasDoc) return
    if (host.hasMergeSource) host.toggleMergePreview()
    else host.pickMergeDataSource()
  },
  'mailings-filter-recipients': recipients('filterRecipients'),
  'mailings-find-recipient': recipients('findRecipient'),
  'file-rename': (host) => {
    if (host.hasDoc && host.onDisk) host.openDialog('rename')
  },
  'file-move': (host) => {
    if (host.hasDoc && host.onDisk) host.moveFile()
  },
  'insert-number': inserting((_, __, host) => host.openDialog('number')),
  // the body's paragraphs: a text box or a header sorts nothing
  sort: inserting((_, __, host) => {
    if (bodyTarget(host)) host.openDialog('sort')
  }),
  // no payload opens the dialog; a payload is a direction of the target at the caret
  'format-text-direction': (host, payload) => {
    const { editor, section } = host
    if (!editor || !host.hasDoc || !editor.isEditable) return
    if (payload === undefined) return host.openDialog('textDirection')
    const target = textDirectionTarget(editor)
    // run as the target's own command, so Repeat turns the same kind of target again
    if (target === 'textBox') return runMenuCommand(host, 'shape-text-direction', payload)
    if (target === 'cells') return runMenuCommand(host, 'table-text-direction', payload)
    if (!section || (payload !== 'lrTb' && payload !== SECTION_VERTICAL)) return
    // a section already vertical keeps the flow its file gave it
    const textDirection = payload === 'lrTb' ? payload : (sectionVertical(section) ?? payload)
    const next = documentSection(section, { textDirection }, host.gridBasePt)
    if (next) host.setSection(next)
  },
  'file-properties': (host) => {
    if (host.hasDoc) host.openDialog('properties')
  },
  'file-reduce-size': (host) => {
    if (host.hasDoc && host.editor?.isEditable) host.openDialog('reduceFileSize')
  },
  'file-always-read-only': (host) => {
    if (host.hasDoc) host.toggleReadOnlyRecommended()
  },
  // the merge panel on one document of every record, Word's Edit Individual Documents
  'mailings-finish-merge': (host) => {
    if (host.hasDoc) host.openMailMerge('single')
  },
  'tools-spelling': (host) => {
    if (host.hasDoc) host.checkSpelling()
  },
  'tools-thesaurus': (host) => {
    if (host.hasDoc) host.thesaurus()
  },
  'tools-translate-selection': (host) => {
    if (host.hasDoc) host.translate('selection')
  },
  'tools-translate-document': (host) => {
    if (host.hasDoc) host.translate('document')
  },
  'tools-translation-language': (host) => {
    if (host.hasDoc) host.openDialog('translationLanguage')
  },
  // the shortcuts are listed there; they cannot be reassigned yet
  'tools-customize-keyboard': (host) => host.openDialog('customizeKeyboard'),
  // payload: 0 takes the paragraph out of the table of contents, 1-3 put it in at that level
  'toc-add-text': (host, payload) => {
    const level = Number(payload)
    if (bodyTarget(host) && (level === 0 || level === 1 || level === 2 || level === 3))
      host.applyHeadingLevel(level)
  },
  'mailings-start': (host) => {
    if (host.hasDoc) host.openMailMerge('perRecord')
  },
  'tools-envelopes': (host) => {
    if (host.hasDoc) host.openMailMerge('envelopes')
  },
  'tools-labels': (host) => {
    if (host.hasDoc) host.openMailMerge('labels')
  },
  'tools-language': (host) => {
    const target = formatTarget(host)
    if (host.hasDoc && target) host.openLanguage(target)
  },
  'format-fit-text': (host) => {
    const target = formatTarget(host)
    if (host.hasDoc && target) host.openFitText(target)
  },
  'format-two-lines-in-one': (host) => {
    const target = formatTarget(host)
    if (host.hasDoc && target) host.toggleAsianLayout(target, 'combine')
  },
  'format-horizontal-in-vertical': (host) => {
    const target = formatTarget(host)
    if (host.hasDoc && target) host.toggleAsianLayout(target, 'vert')
  },
  'format-enclose-characters': (host) => {
    const target = formatTarget(host)
    if (host.hasDoc && target) host.openEqLayout(target, 'enclose')
  },
  'format-combine-characters': (host) => {
    const target = formatTarget(host)
    if (host.hasDoc && target) host.openEqLayout(target, 'combine')
  },
  'format-text-effects': (host) => {
    const target = formatTarget(host)
    if (host.hasDoc && target) host.openTextEffects(target)
  },
  // payload: the id of a ready-made look
  'text-effects-preset': (host, payload) => {
    const target = formatTarget(host)
    if (host.hasDoc && target) applyTextFxPreset(target, payload)
  },
  'format-phonetic-guide': (host) => {
    const target = formatTarget(host)
    if (host.hasDoc && target) host.openPhoneticGuide(target)
  },
  'format-tabs': (host) => {
    const target = formatTarget(host)
    if (host.hasDoc && target) host.openTabsDialog(target)
  },
  'format-equation-options': (host) => {
    if (host.editor && host.hasDoc) host.openDialog('equationOptions')
  },
  'format-columns': (host) => {
    if (host.section && formatTarget(host)) host.openDialog('columns')
  },
  'format-document': (host) => {
    if (host.editor && host.hasDoc) host.openDialog('document')
  },
  'document-theme': applyTheme,
  'theme-fonts': applyTheme,
  'theme-colors': applyTheme,
  'page-color': designing((design, hex) => design.setPageColor(hex === NONE ? null : hex)),
  watermark: designing((design, text) => design.setWatermark(text === NONE ? null : text)),
  // payload: box for the plain box border, none to remove it
  'page-border': pageSetup((section, kind) => {
    if (kind === NONE) return { ...section, pageBorder: false }
    if (kind !== 'box') return null
    const { pageBorderProps: _custom, ...rest } = section
    return { ...rest, pageBorder: true }
  }),
  'document-settings': pageSetup((section, payload, host) => {
    const changes = JSON.parse(payload) as DocumentSettingsPayload
    const { margins, mirror, gutterAtTop, startType, scope } = changes
    if (margins && !marginsFitPage(margins, section.pageWidth, section.pageHeight)) return null
    if (mirror !== undefined) host.setMirrorMargins(mirror)
    if (gutterAtTop !== undefined) host.setGutterAtTop(gutterAtTop)
    const start = startType !== undefined && SECTION_STARTS.includes(startType) ? startType : null
    if (scope !== 'whole' && scope !== 'forward') {
      if (start) host.setSectionStart(start)
      return documentSection(section, changes, host.gridBasePt)
    }
    // the section the break makes starts the way the dialog says, on a new page unless it says otherwise
    if (start && scope === 'whole') host.setSectionStart(start)
    host.setSectionsIn(
      scope,
      (s) => documentSection(s, changes, host.gridBasePt) ?? s,
      start ?? 'nextPage',
    )
    return null
  }),
  'page-margins': pageSetup((section, payload, host) => {
    const { mirror, ...m } = JSON.parse(payload) as PageMarginsPayload
    if (!marginsFitPage(m, section.pageWidth, section.pageHeight)) return null
    if (mirror !== undefined) host.setMirrorMargins(mirror)
    // a margin the user sets is an ordinary one, not the file's header-proof fixed margin
    return {
      ...section,
      marginTop: m.top,
      marginRight: m.right,
      marginBottom: m.bottom,
      marginLeft: m.left,
      marginTopFixed: undefined,
      marginBottomFixed: undefined,
    }
  }),
  'page-orientation': pageSetup((section, orientation) =>
    (orientation === 'portrait' || orientation === 'landscape') &&
    section.orientation !== orientation
      ? { ...section, orientation, pageWidth: section.pageHeight, pageHeight: section.pageWidth }
      : null,
  ),
  // payload: the paper's portrait width,height in twips
  'page-size': pageSetup((section, size) => {
    const [w, h] = size.split(',').map(Number)
    if (!(w > 0) || !(h > 0)) return null
    const landscape = section.orientation === 'landscape'
    return { ...section, pageWidth: landscape ? h : w, pageHeight: landscape ? w : h }
  }),
  // payload: a number of equal columns
  'page-columns': pageSetup((section, count) =>
    Number.isInteger(Number(count)) && Number(count) > 0
      ? inEqualColumns(section, Number(count))
      : null,
  ),
  'columns-settings': pageSetup((section, payload, host) => {
    const picked = JSON.parse(payload) as ColumnsPayload
    if (!Number.isInteger(picked.columns) || picked.columns < 1) return null
    if (picked.scope !== 'whole' && picked.scope !== 'forward')
      return withColumns(section, picked, section)
    // Word starts the new columns on the same page, or in the next column when asked to
    const start = picked.scope === 'forward' && picked.startNewColumn ? 'nextColumn' : 'continuous'
    host.setSectionsIn(picked.scope, (s) => withColumns(s, picked, section), start)
    return null
  }),
  // Word's Suppress for Current Paragraph: the selected paragraphs drop out of the count, or come back
  'layout-suppress-line-numbers': paragraphAttrs((_, host) => ({
    suppressLineNumbers:
      host.editor && paraAttrsOf(host.editor).suppressLineNumbers === true ? null : true,
  })),
  'layout-line-numbering-options': (host) => {
    if (host.section && formatTarget(host)) host.openDialog('lineNumbers')
  },
  // payload: none, when the count starts over (w:restart), or the whole numbering as JSON
  'layout-line-numbers': pageSetup((section, payload) => {
    if (payload.startsWith('{')) {
      const { countBy, start, distance, restart } = JSON.parse(payload) as LineNumbering
      if (!LINE_NUMBER_RESTARTS.some((k) => k === restart)) return null
      if (!Number.isInteger(countBy) || countBy < 1 || !Number.isInteger(start) || start < 1)
        return null
      return {
        ...section,
        lineNumbers: {
          countBy,
          start,
          restart,
          ...(distance !== undefined && distance >= 0 ? { distance } : {}),
        },
      }
    }
    if (payload === 'none') {
      const { lineNumbers: _off, ...rest } = section
      return rest
    }
    const restart = LINE_NUMBER_RESTARTS.find((k) => k === payload)
    if (!restart) return null
    return { ...section, lineNumbers: { countBy: 1, start: 1, ...section.lineNumbers, restart } }
  }),
  'insert-column-break': inserting((editor) => insertColumnBreak(editor)),
  'insert-text-wrapping-break': inserting((editor) => insertTextWrappingBreak(editor)),
  'insert-section-break': inserting((_, payload, host) => {
    const type = SECTION_BREAKS.find((k) => k === payload)
    if (type) host.insertSectionBreak(type)
  }),
  // the picture may sit in a header, footer or text box: the editor that has the focus holds it
  'picture-wrap': inserting((editor, wrap) => {
    if (wrap) setFloatingWrap(getActiveSubEditor() ?? editor, wrap === NONE ? null : wrap)
  }),
  // payload: the colour the selected graphic is filled with, or none
  'picture-graphics-fill': inserting((editor, payload) => {
    if (payload === NONE) void setSelectedGraphicFill(editor, null)
    else if (payload && /^[0-9A-Fa-f]{6}$/.test(payload))
      void setSelectedGraphicFill(editor, payload.toUpperCase())
  }),
  // payload: horz, or the wps:bodyPr vert value the selected text box takes
  'shape-text-direction': inserting((editor, payload) => {
    const vert = textboxDirectionOf(payload)
    if (vert !== undefined) setSelectedBoxDirection(editor, vert)
  }),
  // payload: front, forward, backward or back
  'arrange-order': inserting((editor, step) => {
    if (step && Object.hasOwn(Z_ORDER_STEPS, step))
      Z_ORDER_STEPS[step](getActiveSubEditor() ?? editor)
  }),
  // payload: an align mode, then what to line up against (page, margin or objects)
  'arrange-align': inserting((editor, payload) => {
    const [mode, frame] = (payload ?? '').split(':')
    const align = ALIGN_MODES.find((k) => k === mode)
    if (align)
      alignSelectedObjects(
        getActiveSubEditor() ?? editor,
        align,
        ALIGN_FRAMES.find((k) => k === frame) ?? 'objects',
      )
  }),
  // payload: group or ungroup
  'arrange-group': inserting((editor, payload, host) => {
    const target = getActiveSubEditor() ?? editor
    if (payload === 'ungroup') ungroupSelectedObjects(target, host.blocks)
    else groupSelectedObjects(target)
  }),
  'picture-replace': inserting((editor, _payload, host) => {
    if (editor.getAttributes('docProtected')?.blockType !== 'image') return
    void host.pickImage().then((picked) => {
      if (picked)
        return replaceSelectedPictureBytes(editor, `data:${picked.mime};base64,${picked.base64}`)
    })
  }),
  // payload: crop or cutout
  'picture-edit': inserting((editor, payload, host) => {
    const kind = PICTURE_DIALOGS.find((k) => k === payload)
    const attrs = editor.getAttributes('docProtected')
    if (kind && attrs?.blockType === 'image' && attrs.imageDataUrl)
      host.openPictureDialog(kind, attrs.imageDataUrl as string)
  }),
  'picture-alt-text': (host) => {
    if (host.editor && host.hasDoc && selectedPictureAlt(host.editor) !== null)
      host.openDialog('altText')
  },
  'picture-size-position': inserting((editor, _payload, host) => {
    if (selectedPictureGeometry(editor)) host.openDialog('sizePosition')
  }),
  'picture-auto-crop': inserting((editor, _payload, host) => {
    void autoCropSelectedPicture(editor).then((cropped) => {
      if (!cropped) host.nothingToCrop()
    })
  }),
  // payload: degrees to turn the selected picture by, clockwise
  'picture-rotate': inserting((editor, payload) => {
    const delta = Number(payload)
    const attrs = editor.getAttributes('docProtected')
    if (!Number.isFinite(delta) || attrs?.blockType !== 'image') return
    const turned = ((((Number(attrs.imageRotDeg) || 0) + delta) % 360) + 360) % 360
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', { imageRotDeg: turned || null })
      .run()
  }),
  // payload: h or v
  'picture-flip': inserting((editor, axis) => {
    const attrs = editor.getAttributes('docProtected')
    if ((axis !== 'h' && axis !== 'v') || attrs?.blockType !== 'image') return
    const key = axis === 'h' ? 'imageFlipH' : 'imageFlipV'
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', { [key]: !attrs[key] })
      .run()
  }),
  // payload: horizontal,vertical alignment on the page; the picture wraps away from its side
  'picture-position': inserting((editor, position) => {
    const [h, v] = (position ?? '').split(',')
    if (!['left', 'center', 'right'].includes(h) || !['top', 'center', 'bottom'].includes(v)) return
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', {
        imageWrap: h === 'right' ? 'square-right' : 'square-left',
        imagePosH: h,
        imagePosV: v,
        // the frame the save writes the pair against, so the page shows the saved place
        imageRelH: 'margin',
        imageRelV: 'margin',
        imageOffsetXEmu: null,
        imageOffsetYEmu: null,
      })
      .run()
  }),
  find: (host) => {
    if (host.hasDoc) host.openFind('find')
  },
  'edit-find-next': (host) => {
    if (host.hasDoc) host.findStep(1)
  },
  'edit-find-previous': (host) => {
    if (host.hasDoc) host.findStep(-1)
  },
  // our Find panel is Word's Find and Replace dialog, options included
  'edit-advanced-find': (host) => {
    if (host.hasDoc) host.openFind('find')
  },
  replace: (host) => {
    if (host.hasDoc) host.openFind('replace')
  },
  goto: (host) => {
    if (host.hasDoc) host.openFind('goto')
  },
  'export-pdf': (host) => host.exportPdf(),
  'export-html': (host) => host.exportHtml(),
  'export-markdown': (host) => host.exportMarkdown(),
  'export-images': (host) => host.exportImages(),
  print: (host) => {
    if (host.hasDoc) host.printDoc()
  },
  'word-count': (host) => {
    if (host.hasDoc) host.wordCount()
  },
  'autocorrect-options': (host) => host.openDialog('autocorrect'),
  preferences: (host) => host.openDialog('preferences'),
  'table-insert-cols-left': tableCommand((_, w) => insertRowsOrColumns(1, 'left', w)),
  'table-insert-cols-right': tableCommand((_, w) => insertRowsOrColumns(1, 'right', w)),
  'table-insert-rows-above': tableCommand((_, w) => insertRowsOrColumns(1, 'above', w)),
  'table-insert-rows-below': tableCommand((_, w) => insertRowsOrColumns(1, 'below', w)),
  'table-insert-cells': tableDialog('insertCells'),
  'table-delete-table': tableCommand(() => deleteTable),
  'table-delete-columns': tableCommand(() => deleteTableColumn),
  'table-delete-rows': tableCommand(() => deleteRow),
  'table-delete-cells': tableDialog('deleteCells'),
  'table-select-table': select('table'),
  'table-select-column': select('column'),
  'table-select-row': select('row'),
  'table-select-cell': select('cell'),
  'table-merge-cells': tableCommand(() => mergeCells),
  'table-split-cells': tableDialog('splitCells'),
  'table-split-table': tableCommand(() => splitTableAtSelection()),
  'table-autofit-contents': tableCommand((_, w) => setTableAutoFit('contents', w)),
  'table-autofit-window': tableCommand((_, w) => setTableAutoFit('window', w)),
  'table-autofit-fixed': tableCommand((_, w) => setTableAutoFit('fixed', w)),
  'table-distribute-rows': tableCommand((editor) => {
    distributeRowsEvenly(editor.view)
  }),
  'table-distribute-columns': tableCommand((_, w) => distributeSelectedColumns(w)),
  'table-repeat-header': tableCommand(() => toggleRepeatHeaderRows()),
  // payload: vertical,horizontal
  'table-cell-align': tablePick((position) => {
    const [vertical, horizontal] = position.split(',')
    const v = CELL_V_ALIGNS.find((k) => k === vertical)
    const h = CELL_H_ALIGNS.find((k) => k === horizontal)
    return v && h ? setCellAlignment(v, h) : null
  }),
  'table-text-direction': tablePick((payload) => {
    const direction = CELL_TEXT_DIRECTIONS.find((k) => k === payload)
    return direction ? setCellTextDirection(direction) : null
  }),
  // payload: the table style's id, or none to drop the style
  'table-style': tablePick((styleId, editor) => {
    editor
      .chain()
      .updateAttributes('docTable', { tblStyleId: styleId === NONE ? null : styleId })
      .run()
  }),
  // payload: option=true|false
  'table-look': tablePick((payload) => {
    const [name, value] = payload.split('=')
    const option = TABLE_LOOK_OPTIONS.find((k) => k === name)
    return option && (value === 'true' || value === 'false')
      ? setTableLookOption(option, value === 'true')
      : null
  }),
  // payload: the preset as JSON; its look is saved as a table style of the document
  'table-preset': (host, payload) => {
    if (payload === undefined) return
    const preset = JSON.parse(payload) as TablePreset
    tableCommand(() => {
      host.defineTableStyle({ styleId: preset.styleId, sourceXml: tablePresetStyleXml(preset) })
      return applyTablePreset(preset)
    })(host)
  },
  // body paragraphs: a caret in a table or a text box has no text to turn into one
  'convert-text-to-table': inserting((editor, _, host) => {
    if (bodyTarget(host) && textToTableScope(editor.state)) host.openTableDialog('textToTable')
  }),
  'table-convert-to-text': tableDialog('tableToText'),
  // in a table the field goes in the cell; in body text, as in Word, at the caret
  'table-formula': (host) => {
    const { editor } = host
    if (!editor || !host.hasDoc || !editableOrToast(editor)) return
    const inTable = enterSelectedTable(editor.state, editor.view.dispatch)
    if (inTable || (bodyTarget(host) && formulaPlace(editor.state))) host.openTableDialog('formula')
    else host.notInTable()
  },
  'table-gridlines': (host) => host.toggleTableGridlines(),
  'table-properties': tableDialog('properties'),
  'ai-proofread': (host) => {
    if (host.hasDoc) host.aiProofread()
  },
  shortcuts: (host) => host.openDialog('shortcuts'),
}

const repeatEditors = (host: Pick<CommandHost, 'editor'>): (Editor | null)[] => [
  host.editor,
  getActiveSubEditor(),
]

/** the last command can run again: the document is as that command left it */
export const canRepeatCommand = (host: Pick<CommandHost, 'editor'>): boolean =>
  repeatable(repeatEditors(host)) !== null

export function runMenuCommand(
  host: CommandHost,
  command: ActionCommandId,
  payload?: string,
): void {
  runRecorded(
    () => repeatEditors(host),
    command,
    payload,
    () => MENU_COMMAND_HANDLERS[command](host, payload),
  )
}
