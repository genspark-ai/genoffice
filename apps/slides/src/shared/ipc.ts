/**
 * slides main-process <-> renderer IPC contract (Phase 3: open/save/edit, AI not included yet).
 *
 * Architecture: pptx parsing/saving needs node:crypto/Buffer and can only run in the main
 * process (Node). The main process holds the parsed deck (with originalXml/archive) and sends
 * the renderer only plain-data RenderSlide (built by pptx-render, no Node dependency); the
 * renderer sends edit intents (text/geometry changes) back to the main process, which applies
 * them to the model and rebuilds the RenderSlide.
 */
import type { RenderSlide } from '@genoffice/pptx-render'
import type { SlideComment, SectionInfo } from '@genoffice/pptx-engine'
import type {
  AiSettings,
  AiStreamChunk,
  AiStreamRequest,
  GenSparkAccountStatus,
} from '@genoffice/ai-provider'

export type { SlideComment, SectionInfo } from '@genoffice/pptx-engine'

// Canonical definitions of AI-related types live in @genoffice/ai-provider / @genoffice/agent-core (shared with docs)
export type {
  AiProviderConfig,
  AiProviderId,
  AiProviderMeta,
  AiSettings,
  AiStreamChunk,
  AiStreamRequest,
  GenSparkAccountStatus,
} from '@genoffice/ai-provider'
export { AI_PROVIDERS } from '@genoffice/ai-provider'
export type { AgentToolCall, AgentToolDef } from '@genoffice/agent-core'

export interface OpenResult {
  path: string
  slides: RenderSlide[]
  /** Slide page size in EMU */
  size: { cx: number; cy: number }
  /** Theme body default font (fallback shown in the font box when the selection has no text element) */
  defaultFont?: string
}

// ---- Chat attachments (local files fed to the agent via tools; structure copied from apps/docs) ----

/** Image attachment extensions: no text extraction; read as base64 on send and passed to the model as multimodal images with the user message */
export const ATTACHMENT_IMAGE_EXTS = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp'])

export interface AttachmentMeta {
  /** Absolute local path; files never leave this machine */
  path: string
  name: string
  /** Lowercase extension, without the dot */
  ext: string
  sizeBytes: number
}

export interface AttachmentAddResult {
  accepted: AttachmentMeta[]
  /** Per-file rejection reason (too large / unsupported type / unreadable) */
  rejected: string[]
}

export interface AttachmentReadResult {
  ok: boolean
  error?: string
  name?: string
  /** Total character count of the extracted text */
  totalChars?: number
  /** The requested slice */
  text?: string
  offset?: number
}

/** Image attachment raw-byte read (multimodal input, slides:files-read-image) */
export interface AttachmentImageResult {
  ok: boolean
  /** raw base64 (without the data: prefix) */
  base64?: string
  mime?: string
  error?: string
}

/** Attachment bridge (window.desktop): same names/signatures as docs' DesktopApi attachment subset, so files-skill can be copied wholesale */
export interface DesktopFilesApi {
  /** Multi-select attachment file dialog */
  pickAttachments(): Promise<AttachmentAddResult | null>
  /** Validate dragged-in paths and return attachment metadata */
  addAttachmentPaths(paths: string[]): Promise<AttachmentAddResult>
  /** Save a clipboard-pasted image (no local path) to a temp file and add it as an attachment */
  addPastedImage(data: ArrayBuffer, ext: string): Promise<AttachmentAddResult>
  /** Read one slice of an attachment's extracted text */
  readAttachment(path: string, offset: number, maxChars: number): Promise<AttachmentReadResult>
  /** Read an image attachment as base64 for multimodal (≤5MB) */
  readAttachmentImage(path: string): Promise<AttachmentImageResult>
  /** Absolute path of a File dropped on the window (Electron webUtils) */
  getPathForFile(file: File): string
}

/** One rich-text run (sent by the editor, with independent formatting). */
export interface EditRun {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontSize?: number
  fontFamily?: string
  color?: string
  /** Strikethrough (DOM-authoritative boolean, like bold/italic/underline) */
  strike?: boolean
  /** Super/subscript baseline % (positive = superscript; 0 = none, used to disable explicitly) */
  baseline?: number
  /** Text outline (for WordArt), width in EMU */
  outline?: { color: string; widthEmu: number }
  /** Dynamic field (slidenum / datetime1…); text is the cached value */
  field?: string
  /** Source model run index (index into the original paragraph's runs); the main process uses it to backtrack unedited format fields */
  srcRun?: number
  /** Run hyperlink. undefined = keep the original run's link (programmatic paths that can't
   * express links); null = explicitly none (the editor DOM is authoritative, link removed) */
  link?: LinkTargetOp | null
}

/** One paragraph (with alignment). */
export interface EditParagraph {
  runs: EditRun[]
  align?: 'left' | 'center' | 'right' | 'justify'
  /** Indent level 0..8 (returned after editor Tab/⇧Tab adjustment; defaults to the original paragraph's) */
  level?: number
  /** Source model paragraph index; the main process uses it to inherit bullet/line spacing etc. (both halves of a split share a source) */
  srcPara?: number
  /** Per-paragraph format explicitly changed during this edit session (absent = keep the original) */
  bullet?: 'char' | 'number' | 'none'
  bulletChar?: string
  lineSpacingPct?: number
  spaceBeforePt?: number
  spaceAfterPt?: number
}

/**
 * Text edit intent (run-level rich text): replace the element's text by paragraph/run structure.
 * The editor preserves each run's independent formatting (no longer flattens the whole box to one format).
 */
export interface EditTextOp {
  slideIndex: number
  sourceId: string
  paragraphs: EditParagraph[]
  /** In-group editing: sourceId is a direct child of that group */
  groupId?: string
}

/**
 * Change font/size directly on a selected element (not in text-editing mode): applies to all of
 * the element's text runs (text/shape/table; like changing font on a selected shape in PowerPoint).
 */
export interface SetElementFontOp {
  slideIndex: number
  sourceIds: string[]
  fontFamily?: string
  fontSizePt?: number
  /** Bold/italic/underline/strike toggles (apply to all runs; selected shapes change directly without entering edit mode) */
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  /** Font color #RRGGBB */
  color?: string
  /** In-group editing: all sourceIds are direct children of that group */
  groupId?: string
}

/**
 * Change paragraph format on a selected element (bullet/line spacing/paragraph spacing/align):
 * applies to all of the element's paragraphs (text/shape/table; like clicking bullets on a
 * selected shape in PowerPoint).
 */
export interface SetElementParagraphFormatOp {
  slideIndex: number
  sourceIds: string[]
  /** 'char' bullet dot / 'number' numbered / 'none' explicitly none */
  bullet?: 'char' | 'number' | 'none'
  /** Custom bullet character (with bullet: 'char'; defaults to '•') */
  bulletChar?: string
  /** Bullet hanging indent (EMU); alone it adjusts existing bullets' indent */
  bulletHangEmu?: number
  /** Bullet size (% of text size, 100 = same); alone it only touches bulleted paragraphs */
  bulletSizePct?: number
  /** Bullet color (#RRGGBB); alone it only touches bulleted paragraphs */
  bulletColor?: string
  /** Line spacing (%, 100 = single) */
  lineSpacingPct?: number
  /** Space before / after (pt) */
  spaceBeforePt?: number
  spaceAfterPt?: number
  align?: 'left' | 'center' | 'right' | 'justify'
  /** Indent level increment/decrement (multi-level lists; applies to all paragraphs) */
  indentDelta?: 1 | -1
  /** In-group editing: all sourceIds are direct children of that group */
  groupId?: string
}

/** Whole-picture opacity (0..1; 1 = opaque, clears the marker). */
export interface EditPictureOpacityOp {
  slideIndex: number
  sourceId: string
  opacity: number
}

/** Slide size (EMU; 16:9=12192000×6858000, 4:3=9144000×6858000). */
export interface SetSlideSizeOp {
  cx: number
  cy: number
}

/** Switch the layout of an existing page; omitted layoutPath = reset layout (placeholder geometry restored). */
export interface SetSlideLayoutOp {
  slideIndex: number
  layoutPath?: string
}

// ── Master edit view (exception to the fidelity rule: only user-modified layout/master parts are written back) ──

export interface MasterPartItem {
  /** Path inside the zip, e.g. ppt/slideMasters/slideMaster1.xml */
  partPath: string
  kind: 'master' | 'layout'
  /** <p:cSld name> (layouts commonly "Title Slide" etc.) */
  name: string
  slide: RenderSlide
}

export interface MasterEnterResult {
  /** master first, its layouts after; on entry the main process has already set items[0] as the edit target */
  items: MasterPartItem[]
}

export interface MasterEditTextOp {
  sourceId: string
  paragraphs: EditParagraph[]
}

export interface MasterEditTransformOp {
  sourceId: string
  xPx: number
  yPx: number
  wPx: number
  hPx: number
  rotationDeg: number
  fitWidthPx: number
  preview?: boolean
}

export interface MasterEditFillOp {
  sourceId: string
  fill: string | GradientFillSpec
}

export interface MasterEditStrokeOp {
  sourceId: string
  stroke: { color: string; widthPt: number } | null
}

export interface MasterDeleteElementOp {
  sourceId: string
}

/** Find/replace: matches within runs across pages; firstOnly+elementId used by "Replace" for the current hit. */
export interface FindReplaceOp {
  find: string
  replace: string
  matchCase?: boolean
  firstOnly?: boolean
  slideIndex?: number
  elementId?: string
}

/**
 * Special value 'textbox' = plain text box; anything else is any OOXML preset geometry name
 * (rect/roundRect/ellipse/triangle/star5/rightArrow/chevron…).
 */
export type InsertKind = 'textbox' | (string & {})

/** Add-element intent: pixel coordinates (relative to the fitWidth viewport); the main process converts back to EMU. */
export interface AddElementOp {
  slideIndex: number
  kind: InsertKind
  xPx: number
  yPx: number
  wPx: number
  hPx: number
  fitWidthPx: number
  /** Initial text (split into paragraphs by \n; mutually exclusive with paragraphs) */
  text?: string
  /** Rich-text paragraphs (takes precedence over text) */
  paragraphs?: EditParagraph[]
  /** Shape solid fill #RRGGBB */
  fillColor?: string
  /** Shape stroke (solid color + point width) */
  stroke?: { color: string; widthPt: number }
}

export interface DeleteElementOp {
  slideIndex: number
  sourceId: string
}

/** Mirror an element across its own axis (a:xfrm flipH/flipV); toggles the current value. */
export interface FlipElementOp {
  slideIndex: number
  sourceIds: string[]
  axis: 'h' | 'v'
  /** In-group editing: sourceIds are direct children of that group */
  groupId?: string
}

/** Fill edit: solid color value #RRGGBB or 'none'. */
/** Gradient fill (UI two colors + direction; radial=radial) */
export interface GradientFillSpec {
  gradient: { from: string; to: string; angleDeg?: number; radial?: boolean }
}

export interface EditFillOp {
  slideIndex: number
  sourceId: string
  /** 'none' | #RRGGBB(AA, optional alpha) | gradient */
  fill: string | GradientFillSpec
  /** In-group editing: sourceId is a direct child of that group */
  groupId?: string
}

/** Stroke edit: null = no stroke; widthPt is the line width (points); dash is an OOXML prstDash preset ('solid' clears it, undefined keeps the file's value). */
export interface EditStrokeOp {
  slideIndex: number
  sourceId: string
  stroke: { color: string; widthPt: number; dash?: string } | null
  /** In-group editing: sourceId is a direct child of that group */
  groupId?: string
}

/** Solid page background; slideIndex=-1 applies to all pages. */
export interface EditBackgroundOp {
  slideIndex: number
  color: string
  fitWidthPx: number
}

/** Copy the selected elements to the in-app clipboard (any type, including tables/charts/groups). */
export interface CopyElementsOp {
  slideIndex: number
  sourceIds: string[]
}

/** Paste clipboard elements onto the given page (repeated pastes auto-cascade the offset). */
export interface PasteElementsOp {
  slideIndex: number
  fitWidthPx: number
}

/** Duplicate elements in place (⌘D / Option+drag copy): bypasses the app clipboard; the caller supplies the offset. */
export interface DuplicateElementsOp {
  slideIndex: number
  sourceIds: string[]
  dxPx: number
  dyPx: number
  fitWidthPx: number
}

/**
 * Freehand ink stroke commit: one transparent PNG picture element per stroke. The pixel box is
 * the stroke's bounding box relative to the fitWidth viewport; payload is the editor's vector
 * points as JSON (written into cNvPr descr, erasable after reopening).
 */
export interface AddInkOp {
  slideIndex: number
  /** base64 of the transparent PNG (without the data: prefix) */
  base64: string
  xPx: number
  yPx: number
  wPx: number
  hPx: number
  fitWidthPx: number
  payload: string
}

/** Insert a table: pixel box + row/column counts; the main process converts back to EMU. */
export interface AddTableOp {
  slideIndex: number
  rows: number
  cols: number
  xPx: number
  yPx: number
  wPx: number
  hPx: number
  fitWidthPx: number
}

/**
 * Apply a theme (Design tab theme gallery): rewrite the color/font scheme in the package's
 * theme*.xml (scheme-referenced colors follow), and remap the deck's explicit colors wholesale
 * to the new theme palette; the main process reparses and sends back the full RenderSlide set.
 */
export interface ApplyThemeOp {
  /** Theme name (written into clrScheme name) */
  name: string
  /** dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink -> #RRGGBB */
  colors: Record<string, string>
  majorFont?: string
  minorFont?: string
  fitWidthPx: number
}

export type TransitionKind =
  | 'none'
  | 'morph'
  | 'fade'
  | 'push'
  | 'wipe'
  | 'split'
  | 'circle'
  | 'cover'
  | 'pull'
  | 'dissolve'
  | 'zoom'
  | 'random'

// ── Shape animations (the "Animations" tab) ──────────────────────────

export type AnimEffectKind =
  | 'appear'
  | 'fade'
  | 'flyIn'
  | 'wipe'
  | 'wipeDown'
  | 'splitIn'
  | 'bounce'
  | 'flipIn'
  | 'zoom' // entrance
  | 'pulse'
  | 'spin'
  | 'grow'
  | 'teeter' // emphasis
  | 'disappear'
  | 'fadeOut'
  | 'flyOut'
  | 'wipeOut'
  | 'shrink'
  | 'zoomOut' // exit
  | 'motionPath' // motion path (move along a path)

export type AnimTrigger = 'onClick' | 'withPrev' | 'afterPrev'

/** One animation (list order = play order); sourceId locates the target element. */
export interface AnimationItem {
  sourceId: string
  /** Target element display name (shown in the animation pane, filled by the main process) */
  targetName: string
  effect: AnimEffectKind
  trigger: AnimTrigger
  durationMs: number
  delayMs: number
  /** Path when effect='motionPath' (SVG subset M/L/C/Z, coordinates 0..1 relative to slide width/height) */
  motionPath?: string
  /** Per-paragraph animation: 0-based paragraph number; default = the whole shape */
  paragraph?: number
}

/** Element pairing key for Morph transitions: cNvPr id / name are stable across pages; sourceId is per-page temporary. */
export interface ShapeKey {
  sourceId: string
  spid: number | null
  name: string
}

/** Overwrite-set the whole page's animation list (add/remove/reorder/param changes all use this). */
export interface SetAnimationsOp {
  slideIndex: number
  items: Array<Omit<AnimationItem, 'targetName'>>
}

/** Set the transition effect; slideIndex=-1 applies to all pages. */
export interface SetTransitionOp {
  slideIndex: number
  kind: TransitionKind
}

/** Batch-write each page's auto-advance time (<p:transition advTm>, ms; ms=null clears). Used by rehearsal timing save. */
export interface SetAdvanceTimesOp {
  times: Array<{ slideIndex: number; ms: number | null }>
}

// ── Section management ────────────────────────────────────────────────

/** Add a section before slide atSlideIndex (the section covers this page to the end of its containing section). */
export interface AddSectionOp {
  atSlideIndex: number
  name: string
}

export interface RenameSectionOp {
  id: string
  name: string
}

/** Delete a section but keep its slides (pages merge into the adjacent section). */
export interface RemoveSectionOp {
  id: string
}

/** Move a whole section up/down (swapping together with its slides). */
export interface MoveSectionOp {
  id: string
  dir: 'up' | 'down'
}

/** Drag to reorder slides: move slide fromIndex to toIndex (the landing index after remove-then-insert). */
export interface MoveSlideOp {
  fromIndex: number
  toIndex: number
}

/** Hide/unhide a slide (writes <p:sld show="0">, skipped during the show). */
export interface SetSlideHiddenOp {
  slideIndex: number
  hidden: boolean
}

/** Duplicate a page: copy slide sourceIndex and insert after it; clearText empties text yielding a layout-preserving blank page. */
export interface AddSlideOp {
  sourceIndex: number
  clearText?: boolean
  fitWidthPx: number
}

/**
 * Slide paste modes, mirroring PowerPoint's paste options:
 * 'theme' (default) re-binds the slide to a destination layout, 'source' imports
 * the source layout→master→theme chain, 'picture' drops a rendering of the page
 * onto the anchor slide as a picture element.
 */
export type PasteSlideMode = 'theme' | 'source' | 'picture'

/** Paste the copied slide after slideIndex (-1 = at the front). */
export interface PasteSlideOp {
  afterIndex: number
  fitWidthPx: number
  mode?: PasteSlideMode
}

/** Redo the immediately preceding slide paste with another mode (paste-options floater). */
export interface RepasteSlideOp {
  mode: PasteSlideMode
  fitWidthPx: number
}

/** New blank page: inserted after slide sourceIndex, reusing its layout (background/decoration), content empty. */
export interface AddBlankSlideOp {
  sourceIndex: number
  fitWidthPx: number
}

/** New blank page (with a specific layout): inserted after slide sourceIndex, rels pointing at layoutPath. */
export interface AddSlideWithLayoutOp {
  sourceIndex: number
  /** Layout path inside the zip, e.g. 'ppt/slideLayouts/slideLayout3.xml' */
  layoutPath: string
  fitWidthPx: number
}

/** Query the pptx's slideLayout list (for the new-slide dropdown panel). */
export interface GetLayoutsResult {
  layouts: Array<{
    /** Zip path; 'builtin:<key>' = built-in standard layout, injected into the package on first use */
    path: string
    name: string
    layoutType: string
    /** Placeholder summary (type/idx/geometry) */
    placeholders: Array<{
      type: string
      idx: string
      x: number
      y: number
      cx: number
      cy: number
      hint: string
    }>
  }>
  /** Slide size (EMU), for normalizing the placeholder previews */
  size: { cx: number; cy: number }
}

/** Element z-order adjustment (elements order = spTree order). */
export type ReorderDirection = 'front' | 'back' | 'forward' | 'backward'
export interface ReorderElementOp {
  slideIndex: number
  sourceId: string
  dir: ReorderDirection
}

/** Table cell text edit: row/col are the model coordinates carried by the render node's cells. */
export interface EditTableCellOp {
  slideIndex: number
  sourceId: string
  row: number
  col: number
  paragraphs: EditParagraph[]
}

/** Table row/column insert/delete; index is the row/column number (tc index), before=true inserts before it. */
export interface TableStructureIpcOp {
  slideIndex: number
  sourceId: string
  kind: 'insert-row' | 'delete-row' | 'insert-col' | 'delete-col'
  index: number
  before?: boolean
}

/** Merge/split cells (row/col are model coordinates; col is the tc index). */
export interface TableMergeIpcOp {
  slideIndex: number
  sourceId: string
  kind: 'merge-right' | 'merge-down' | 'split'
  row: number
  col: number
}

/** Drag-resize column width (pixel value; the main process converts back to EMU). */
export interface SetTableColWidthOp {
  slideIndex: number
  sourceId: string
  col: number
  wPx: number
  fitWidthPx: number
}

/** Drag-resize row height (pixel value; the main process converts back to EMU). */
export interface SetTableRowHeightOp {
  slideIndex: number
  sourceId: string
  row: number
  hPx: number
  fitWidthPx: number
}

/** Vertical alignment of cell text. */
export interface SetTableCellAnchorOp {
  slideIndex: number
  sourceId: string
  row: number
  col: number
  anchor: 'top' | 'middle' | 'bottom'
}

/** Picture crop edit: a null srcRect resets to the full image. */
export interface EditPictureSrcRectOp {
  slideIndex: number
  sourceId: string
  /** Crop ratio per edge 0..1; null = remove the crop (full image) */
  srcRect: { l: number; t: number; r: number; b: number } | null
}

/** Group elements: merge ≥2 editable elements into one group. */
export interface GroupElementsOp {
  slideIndex: number
  /** Ids of the elements to group (≥2, all must be text/shape/picture) */
  sourceIds: string[]
}

/** Ungroup: promote the group's children to top-level slide elements. */
export interface UngroupElementOp {
  slideIndex: number
  /** Group element id */
  sourceId: string
}

/**
 * Batch geometry transform (multi-element position ops like align/distribute).
 * Each item is equivalent to an independent editTransform; only positions update, size/rotation unchanged.
 */
export interface BatchEditTransformOp {
  slideIndex: number
  fitWidthPx: number
  items: Array<{
    sourceId: string
    xPx: number
    yPx: number
    wPx: number
    hPx: number
    rotationDeg: number
  }>
}

/** Geometry transform intent: move/resize/rotate. */
export interface EditTransformOp {
  slideIndex: number
  sourceId: string
  /** In-group editing: sourceId is a direct child of that group; the pixel box is in group-local coordinates */
  groupId?: string
  /** Target pixel box (relative to the current fitWidth viewport) + viewport width; the main process converts back to EMU. */
  xPx: number
  yPx: number
  wPx: number
  hPx: number
  rotationDeg: number
  fitWidthPx: number
  /**
   * Live preview during drag (text reflows to the new box width in real
   * time): the first preview of a gesture pushes one undo snapshot; later previews
   * and the final commit (preview omitted/false) push nothing — a whole drag takes one undo step.
   */
  preview?: boolean
}

/**
 * Connector endpoint edit: new endpoint positions in viewport px,
 * plus optional attachment changes for either end (undefined = keep the current
 * attachment, null = detach, object = attach to targetId's connection point idx
 * — 0 top, 1 left, 2 bottom, 3 right).
 */
export interface EditConnectorEndpointsOp {
  slideIndex: number
  sourceId: string
  x1Px: number
  y1Px: number
  x2Px: number
  y2Px: number
  fitWidthPx: number
  start?: { targetId: string; idx: number } | null
  end?: { targetId: string; idx: number } | null
}

/** Overwrite-write speaker notes (\n splits paragraphs). */
export interface SetNotesOp {
  slideIndex: number
  text: string
}

/** Add a comment (the author is the system username fetched by the main process). */
export interface AddCommentOp {
  slideIndex: number
  text: string
}

/** Delete a comment: uniquely located by (authorId, idx). */
export interface DeleteCommentOp {
  slideIndex: number
  authorId: number
  idx: number
}

// ── New insert capabilities (charts / SmartArt / icons / audio-video / 3D / links / header-footer) ──

/** Insert a chart: built-in sample or custom data; the main process writes the chart part. */
export interface AddChartOp {
  slideIndex: number
  /** 'barH' = horizontal bar (mapped to kind 'bar' + barDir 'bar' in the main process) */
  kind:
    | 'bar'
    | 'barStacked'
    | 'barPercentStacked'
    | 'barH'
    | 'line'
    | 'area'
    | 'pie'
    | 'doughnut'
    | 'scatter'
    | 'radar'
    | 'comboBarLine'
  title?: string
  categories: string[]
  series: Array<{ name: string; values: number[] }>
  xPx: number
  yPx: number
  wPx: number
  hPx: number
  fitWidthPx: number
}

/** Insert SmartArt (simplified shape-group version). */
export interface AddSmartArtOp {
  slideIndex: number
  layout: 'list' | 'process' | 'cycle' | 'hierarchy' | 'pyramid' | 'matrix' | 'venn'
  items: string[]
  xPx: number
  yPx: number
  wPx: number
  hPx: number
  fitWidthPx: number
}

/** Insert a renderer-generated bitmap (rasterized icon library / screenshots etc.). */
export interface AddImageBytesOp {
  slideIndex: number
  /** base64 without the data: prefix */
  base64: string
  ext: string
  xPx: number
  yPx: number
  wPx: number
  hPx: number
  fitWidthPx: number
  name?: string
}

/** Insert renderer-recorded media bytes (screen-recording webm etc.). */
export interface AddMediaBytesOp {
  slideIndex: number
  kind: 'video' | 'audio'
  base64: string
  ext: string
  fitWidthPx: number
  name?: string
}

/** Element hyperlink target. */
export type LinkTargetOp = { kind: 'url'; url: string } | { kind: 'slide'; slideIndex: number }

export interface SetLinkOp {
  slideIndex: number
  sourceId: string
  /** null = clear the link */
  target: LinkTargetOp | null
}

/** Header/footer (applied to all pages). */
export interface HeaderFooterOp {
  footer?: string | null
  slideNum?: boolean
  date?: string | null
  /** Write the date as a dynamic datetime field (auto-updates when opened in PowerPoint) */
  dateAuto?: boolean
  fitWidthPx: number
}

/** Table style edit: apply a preset style (whole table) or change header/banding toggles/shading/borders. */
export interface EditTableStyleOp {
  slideIndex: number
  sourceId: string
  /** Preset style name (see TABLE_STYLE_PRESETS); takes precedence over the other fields */
  styleName?: string
  /** Header row toggle (first-row emphasis) */
  firstRow?: boolean
  /** Banded rows toggle */
  bandRow?: boolean
  /** Shading color #RRGGBB or 'none' (null = unchanged) */
  shadingColor?: string | null
  /** Border color #RRGGBB (null = unchanged) */
  borderColor?: string | null
  /** Border width (pt, null = unchanged) */
  borderWidthPt?: number | null
  /** 'all' = all border lines; 'none' = clear border lines; null = unchanged */
  borderPreset?: 'all' | 'none' | null
  /** Cells the shading/border edit applies to, as (row, tc index); undefined = whole table */
  cells?: Array<{ row: number; col: number }>
}

/** Chart edit (charts created by this app): change type / data / colors. */
export interface EditChartOp {
  slideIndex: number
  sourceId: string
  /** Change chart type (rebuilds the chart part); undefined = unchanged; 'barH' = horizontal bar */
  kind?:
    | 'bar'
    | 'barStacked'
    | 'barPercentStacked'
    | 'barH'
    | 'line'
    | 'area'
    | 'pie'
    | 'doughnut'
    | 'scatter'
    | 'radar'
    | 'comboBarLine'
  /** Replace data (rebuilds the chart part); undefined = unchanged */
  categories?: string[]
  series?: Array<{ name: string; values: number[] }>
  /** Color scheme name (see CHART_COLOR_SCHEMES); undefined = unchanged */
  colorScheme?: string
  title?: string
  /** Chart element toggles (undefined = unchanged; styles not specified during a rebuild keep the current state) */
  legendPos?: 'b' | 't' | 'r' | 'l' | 'none'
  dataLabels?: boolean
  gridlines?: boolean
  /** Axis title: '' = clear */
  catAxisTitle?: string
  valAxisTitle?: string
  /** Bar gap % (c:gapWidth) */
  gapWidthPct?: number
  /** Switch rows/columns: categories <-> series */
  switchRowCol?: boolean
  /** Per-point fill overrides, seriesIdx → pointIdx → color (null clears back to the series color) */
  pointColors?: Record<number, Record<number, string | null>>
}

// ── Export (PDF / images) ─────────────────────────────────────────────

/** Export as images: the renderer has already rendered hi-res PNGs; the main process only writes them to disk. */
export interface ExportImagesOp {
  /** Target directory (absolute path chosen via pickExportDir) */
  dir: string
  /** File base name (without extension), written as <baseName>-01.png / -02.png … */
  baseName: string
  /** base64 per page PNG (without the data: prefix), in page order */
  pngsBase64: string[]
}

export interface ExportImagesResult {
  ok: boolean
  /** Absolute paths of the written files (in page order) */
  paths?: string[]
  error?: string
}

/** Export as PDF: the main process loads each page PNG in a hidden window then printToPDF. */
export interface ExportPdfOp {
  /** Target pdf absolute path (chosen via pickExportPdfPath) */
  filePath: string
  /** base64 per page PNG (without the data: prefix), in page order */
  pngsBase64: string[]
  /** Rendered pixel width/height of the slide page (used to compute the PDF page aspect ratio) */
  widthPx: number
  heightPx: number
}

export interface ExportPdfResult {
  ok: boolean
  path?: string
  error?: string
}

/** Print: same page assembly as ExportPdfOp, using the system print dialog. */
export interface PrintSlidesOp {
  pngsBase64: string[]
  widthPx: number
  heightPx: number
  /** Layout: full slides / handouts (N per page) / notes pages (slide + notes text) */
  layout?: 'full' | 'handout2' | 'handout3' | 'handout6' | 'notes'
  /** Per-page notes text for the notes layout (same order as pngsBase64) */
  notes?: string[]
}

/** Show sync state from presenter view -> audience window (absolute state mirror; audience seek is idempotent) */
export interface ShowSyncState {
  /** Original index of the current page */
  idx: number
  /** Number of in-page animation steps already played */
  played: number
  /** Whether the current step is playing (audience plays that step from the start) */
  playing: boolean
  /** Whether this page change is forward (audience plays the transition effect) */
  fresh: boolean
  /** Reached the "end of show" black screen */
  ended: boolean
  /** Presenter toggled black screen (B key/toolbar button) */
  black: boolean
}

/** Presenter ink/laser event; coordinates are 0..1 normalized relative to the slide area (laser x<0 = off) */
export type ShowInkEvent =
  | { type: 'laser'; x: number; y: number }
  | { type: 'stroke-start'; x: number; y: number; color: string }
  | { type: 'stroke-move'; x: number; y: number }
  | { type: 'clear' }

/** Navigation actions sent back from the audience window (click/keypress) */
export type AudienceNavAction = 'next' | 'prev' | 'exit'

export type MenuCommand =
  | 'open'
  | 'save'
  | 'save-as'
  | 'export-pdf'
  | 'export-images'
  | 'print'
  | 'zoom-in'
  | 'zoom-out'
  | 'zoom-reset'
  | 'undo'
  | 'redo'
  | 'cut'
  | 'copy'
  | 'paste'

export interface SlidesApi {
  /** current UI language (persisted by the shell in app-settings.json) */
  getLanguage: () => Promise<
    'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar'
  >
  /** language switched from the shell home page */
  onLanguageChanged: (
    handler: (
      lang: 'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar',
    ) => void,
  ) => () => void
  openPptx: (fitWidthPx: number) => Promise<OpenResult | null>
  openPptxPath: (path: string, fitWidthPx: number) => Promise<OpenResult | null>
  consumePendingOpen: (fitWidthPx: number) => Promise<OpenResult | null>
  /** New blank presentation (single blank 16:9 page, untitled) */
  newBlank: (fitWidthPx: number) => Promise<OpenResult>
  /** HTML pipeline generation: mode="append" merges with the previously generated pages and rebuilds wholesale (appendedFrom = existing page count);
   *  mode="replace_at" redoes page atIndex in place from single-page HTML (other pages untouched, undoable, replacedIndex = that page's index);
   *  mode="insert_at" inserts a new page at atIndex from single-page HTML (later pages shift back, undoable, insertedIndex = that page's index);
   *  when the pipeline fails it falls back to element-level mode, fallbackReason explains why;
   *  deckName = the deck name AI derived from user input, used as the filename when saving a new draft (falls back to timestamp naming) */
  htmlToPptx: (
    pagesHtml: string[],
    fitWidthPx: number,
    mode?: 'replace' | 'append' | 'replace_at' | 'insert_at',
    atIndex?: number,
    deckName?: string,
  ) => Promise<
    | (OpenResult & {
        appendedFrom?: number
        replacedIndex?: number
        insertedIndex?: number
        fallbackReason?: string
        imageFailures?: { page: number; url: string }[]
      })
    | { error: string }
  >
  /** Whether cloud single-page generation (gsk slide_generate) is available (GENOFFICE_CLOUD_SLIDE=1 + gsk login) */
  cloudGenStatus: () => Promise<{ enabled: boolean }>
  /** Cloud single-page generation: brief → one-slide pptx temp file; the marker goes into an htmlToPptx pagesHtml slot in place of HTML */
  cloudGeneratePage: (op: {
    brief: string
    title?: string
    styleSkill?: string
    deckContext?: Record<string, unknown>
    images?: { url: string; caption?: string }[]
    width?: number
    height?: number
  }) => Promise<{ ok: boolean; marker?: string; error?: string }>
  editText: (op: EditTextOp) => Promise<RenderSlide | null>
  /** Change font/size on selected elements wholesale (elements without text ignored; returns null if all ignored) */
  setElementFont: (op: SetElementFontOp) => Promise<RenderSlide | null>
  /** Change paragraph format on selected elements (bullet/line spacing/paragraph spacing/align; elements without text ignored) */
  setElementParagraphFormat: (op: SetElementParagraphFormatOp) => Promise<RenderSlide | null>
  /** Replace (when count=0, slides is null and no history step is created) */
  findReplace: (
    op: FindReplaceOp,
  ) => Promise<{ count: number; slides: RenderSlide[] | null } | null>
  /** Switch/reset the layout of an existing page */
  setSlideLayout: (op: SetSlideLayoutOp) => Promise<RenderSlide | null>
  /** Slide size (returns the fully rebuilt RenderSlide set + current size marker) */
  setSlideSize: (op: SetSlideSizeOp) => Promise<RenderSlide[] | null>
  /** Current slide size (EMU) */
  getSlideSize: () => Promise<{ cx: number; cy: number } | null>
  editTransform: (op: EditTransformOp) => Promise<RenderSlide | null>
  /** Connector endpoint drag: reposition ends and attach/detach shape anchors (stCxn/endCxn) */
  editConnectorEndpoints: (op: EditConnectorEndpointsOp) => Promise<RenderSlide | null>
  /** Batch position update (align/distribute); all items share one undo step */
  batchEditTransform: (op: BatchEditTransformOp) => Promise<RenderSlide | null>
  /** Read-only: RenderSlide for every page of the current session (E2E driver/debug use) */
  getRenderSlides: () => Promise<RenderSlide[] | null>
  /** Update the picture crop srcRect (0..1 ratios; null = full image); returns the updated page */
  editPictureSrcRect: (op: EditPictureSrcRectOp) => Promise<RenderSlide | null>
  /** Whole-picture opacity */
  editPictureOpacity: (op: EditPictureOpacityOp) => Promise<RenderSlide | null>
  /** Shape picture fill (the main process shows the image picker dialog; cancel returns null) */
  editImageFill: (op: { slideIndex: number; sourceId: string }) => Promise<RenderSlide | null>
  /** Text box vertical alignment */
  setTextAnchor: (op: {
    slideIndex: number
    sourceId: string
    anchor: 'top' | 'middle' | 'bottom'
  }) => Promise<RenderSlide | null>
  /** External clipboard content probe (internal/slide = last copy came from this app) */
  clipboardExternal: () => Promise<
    | { kind: 'internal' }
    | { kind: 'slide' }
    | { kind: 'image'; base64: string; ext: string }
    | { kind: 'text'; text: string }
    | { kind: 'none' }
  >
  /** Group ≥2 editable elements; returns the updated page + new group id */
  groupElements: (op: GroupElementsOp) => Promise<{ slide: RenderSlide; groupId: string } | null>
  /** Ungroup; children promoted to top level (coordinates converted), returns the updated page */
  ungroupElement: (op: UngroupElementOp) => Promise<RenderSlide | null>
  addElement: (op: AddElementOp) => Promise<{ slide: RenderSlide; sourceId: string } | null>
  deleteElement: (op: DeleteElementOp) => Promise<RenderSlide | null>
  addSlide: (op: AddSlideOp) => Promise<{ slides: RenderSlide[]; index: number } | null>
  /** New blank page (reuses slide sourceIndex's layout, content empty) */
  addBlankSlide: (op: AddBlankSlideOp) => Promise<{ slides: RenderSlide[]; index: number } | null>
  /** Copy a slide onto the app-wide slide clipboard, so another open deck can paste it; pngBase64 is a rendering of the page for 'picture'-mode pastes */
  copySlide: (slideIndex: number, pngBase64?: string) => Promise<boolean>
  /** Paste the clipboard slide into this deck (mode: destination theme / source formatting / picture) */
  pasteSlide: (
    op: PasteSlideOp,
  ) => Promise<{ slides: RenderSlide[]; index: number; sourceId?: string } | null>
  /** Redo the just-completed paste with another mode; null when anything else touched the deck since */
  repasteSlide: (
    op: RepasteSlideOp,
  ) => Promise<{ slides: RenderSlide[]; index: number; sourceId?: string } | null>
  /** Is there a slide on the clipboard? (drives the Paste Slide menu item) */
  hasSlideClipboard: () => Promise<boolean>
  /** Delete a slide (refused when only one page remains); returns the full RenderSlide array */
  deleteSlide: (slideIndex: number) => Promise<RenderSlide[] | null>
  /** Bring element to front/back or move one layer forward/backward */
  reorderElement: (op: ReorderElementOp) => Promise<RenderSlide | null>
  /** Table cell text edit */
  editTableCell: (op: EditTableCellOp) => Promise<RenderSlide | null>
  /** Row/column insert/delete (tables with merged cells refused, returns null); ids change after reparse, requiring a whole-page replace and reselect */
  tableStructure: (
    op: TableStructureIpcOp,
  ) => Promise<{ slide: RenderSlide; sourceId: string } | null>
  /** Merge/split cells */
  tableMerge: (op: TableMergeIpcOp) => Promise<{ slide: RenderSlide; sourceId: string } | null>
  /** Drag-resize column width (element id stable) */
  setTableColWidth: (op: SetTableColWidthOp) => Promise<RenderSlide | null>
  setTableRowHeight: (op: SetTableRowHeightOp) => Promise<RenderSlide | null>
  setTableCellAnchor: (op: SetTableCellAnchorOp) => Promise<RenderSlide | null>
  editFill: (op: EditFillOp) => Promise<RenderSlide | null>
  editStroke: (op: EditStrokeOp) => Promise<RenderSlide | null>
  /** Mirror selected elements horizontally/vertically */
  flipElements: (op: FlipElementOp) => Promise<RenderSlide | null>
  /** Returns the full affected RenderSlide array (when applied to all pages) */
  editBackground: (op: EditBackgroundOp) => Promise<RenderSlide[] | null>
  /** Show the system image picker and insert into the current page; returns the updated page + new element id, cancel returns null, undecodable format returns the error */
  insertImage: (
    slideIndex: number,
    fitWidthPx: number,
  ) => Promise<
    { slide: RenderSlide; sourceId: string } | { error: 'unsupported'; ext: string } | null
  >
  /** Copy elements to the in-app clipboard; returns the number actually copied */
  copyElements: (op: CopyElementsOp) => Promise<number>
  /** Paste; empty clipboard or failure returns null. Note the whole page's element ids update, requiring a whole-page replace */
  pasteElements: (
    op: PasteElementsOp,
  ) => Promise<{ slide: RenderSlide; sourceIds: string[] } | null>
  duplicateElements: (
    op: DuplicateElementsOp,
  ) => Promise<{ slide: RenderSlide; sourceIds: string[] } | null>
  addTable: (op: AddTableOp) => Promise<{ slide: RenderSlide; sourceId: string } | null>
  /** Freehand ink stroke commit (one picture element per stroke); returns the updated page + new element id */
  addInk: (op: AddInkOp) => Promise<{ slide: RenderSlide; sourceId: string } | null>
  /** Insert a chart (writes the chart part + graphicFrame); returns the updated page + new element id */
  addChart: (op: AddChartOp) => Promise<{ slide: RenderSlide; sourceId: string } | null>
  /** Insert SmartArt (simplified shape-group version) */
  addSmartArt: (op: AddSmartArtOp) => Promise<{ slide: RenderSlide; sourceId: string } | null>
  /** Insert a renderer-generated bitmap (rasterized icon library etc.) */
  addImageBytes: (
    op: AddImageBytesOp,
  ) => Promise<
    { slide: RenderSlide; sourceId: string } | { error: 'unsupported'; ext: string } | null
  >
  /** Show the system dialog to pick a video/audio file and embed it into the current page */
  insertMedia: (
    slideIndex: number,
    kind: 'video' | 'audio',
    fitWidthPx: number,
  ) => Promise<{ slide: RenderSlide; sourceId: string } | null>
  /** Insert renderer-recorded media (screen recording); placed centered */
  addMediaBytes: (op: AddMediaBytesOp) => Promise<{ slide: RenderSlide; sourceId: string } | null>
  /** Read an audio/video element's media data (double-click playback); embedded media converts to dataUrl, external links return as-is */
  getMediaData: (
    slideIndex: number,
    sourceId: string,
  ) => Promise<{ kind: 'video' | 'audio'; dataUrl: string } | null>
  /** Show a dialog to pick a 3D model (glb/gltf), embed + poster placeholder */
  insertModel3d: (
    slideIndex: number,
    fitWidthPx: number,
  ) => Promise<{ slide: RenderSlide; sourceId: string } | null>
  /** Set/clear an element hyperlink; the whole page's element ids update, requiring a whole-page replace and clearing the selection */
  setLink: (op: SetLinkOp) => Promise<RenderSlide | null>
  /** Read the element's current hyperlink (dialog echo) */
  getLink: (slideIndex: number, sourceId: string) => Promise<LinkTargetOp | null>
  /** All element hyperlinks on a slide (groups included) — slideshow click hit-testing */
  getSlideLinks: (slideIndex: number) => Promise<Array<{ sourceId: string; target: LinkTargetOp }>>
  /** Run-level hyperlinks on a slide (resolved live); keyed by element + paragraph + run indexes */
  getRunLinks: (
    slideIndex: number,
  ) => Promise<
    Array<{ sourceId: string; paraIndex: number; runIndex: number; target: LinkTargetOp }>
  >
  /** Apply header/footer to all pages; returns the full RenderSlide set */
  applyHeaderFooter: (op: HeaderFooterOp) => Promise<RenderSlide[] | null>
  /** Current page footer state (dialog echo) */
  getHeaderFooter: (
    slideIndex: number,
  ) => Promise<{ footer: string | null; slideNum: boolean; date: string | null }>
  /** Apply a theme (color/font scheme + per-page background); returns the reparsed full RenderSlide set, null = no-op, { error } = failed (state rolled back) */
  applyTheme: (op: ApplyThemeOp) => Promise<RenderSlide[] | { error: string } | null>
  /** Set the transition effect (takes effect in PowerPoint shows of the saved pptx); returns success */
  setTransition: (op: SetTransitionOp) => Promise<boolean>
  /** The current page's transition effect (echoed on page switch) */
  getTransition: (slideIndex: number) => Promise<TransitionKind>
  /** Batch-write each page's auto-advance time (rehearsal timing save; the saved pptx auto-advances in PowerPoint shows); returns success */
  setAdvanceTimes: (op: SetAdvanceTimesOp) => Promise<boolean>
  /** The current page's animation list (read by the Animations tab / during shows) */
  getAnimations: (slideIndex: number) => Promise<AnimationItem[]>
  /** Morph pairing keys of the current page's elements (matched against same-name/same-id elements on the previous page during morph tweening) */
  getShapeKeys: (slideIndex: number) => Promise<ShapeKey[]>
  /** Overwrite-set the whole page's animation list (takes effect in PowerPoint shows of the saved pptx); returns success */
  setAnimations: (op: SetAnimationsOp) => Promise<boolean>
  /** Hide/unhide a slide; returns the updated page's RenderSlide (hidden flag), null on failure */
  setSlideHidden: (op: SetSlideHiddenOp) => Promise<RenderSlide | null>
  /** The current document's section list ([] when there are no sections) */
  getSections: () => Promise<SectionInfo[]>
  /** Overwrite-write the section structure; returns the updated section list */
  setSections: (sections: SectionInfo[]) => Promise<SectionInfo[] | null>
  /** Add a section before the given page; returns the updated section list, null on failure */
  addSection: (op: AddSectionOp) => Promise<SectionInfo[] | null>
  renameSection: (op: RenameSectionOp) => Promise<SectionInfo[] | null>
  /** Delete a section (keeping slides; pages merge into the adjacent section) */
  removeSection: (op: RemoveSectionOp) => Promise<SectionInfo[] | null>
  /** Move a whole section up/down; slide order changes, returns the full RenderSlide set + section list */
  moveSection: (
    op: MoveSectionOp,
  ) => Promise<{ slides: RenderSlide[]; sections: SectionInfo[] } | null>
  /** Drag to reorder slides; returns the full RenderSlide set + section list, null on failure */
  moveSlide: (op: MoveSlideOp) => Promise<{ slides: RenderSlide[]; sections: SectionInfo[] } | null>
  /** Plain text of the current page's speaker notes ('' when there are none) */
  getNotes: (slideIndex: number) => Promise<string>
  /** Overwrite-write notes (into the pptx's notesSlide part); returns success */
  setNotes: (op: SetNotesOp) => Promise<boolean>
  /** All comments on a page (in add order) */
  getComments: (slideIndex: number) => Promise<SlideComment[]>
  /** Add a comment; returns the page's updated comment list, null on failure */
  addComment: (op: AddCommentOp) => Promise<SlideComment[] | null>
  /** Delete a comment; returns the page's updated comment list, null on failure */
  deleteComment: (op: DeleteCommentOp) => Promise<SlideComment[] | null>
  /** System clipboard while text-editing (webContents.cut/copy/paste, for menu command echo) */
  nativeClipboard: (op: 'cut' | 'copy' | 'paste') => Promise<void>
  /** Nestable history transaction; all edits between begin/end become one undo step.
      The outermost end registers an AI rollback point and returns its id (null when nothing changed). */
  beginHistoryBatch: () => Promise<boolean>
  endHistoryBatch: () => Promise<number | null>
  /** Roll the deck back to an AI rollback point; returns the restored full RenderSlide array, null when the id is unknown */
  aiSnapshotRestore: (id: number) => Promise<RenderSlide[] | null>
  /** Undo/redo (main-process snapshot history): returns the restored full RenderSlide array, null when nothing to undo */
  undo: () => Promise<RenderSlide[] | null>
  redo: () => Promise<RenderSlide[] | null>
  // slides: after saving, the main process reopens the file (clears dirty + refreshes byte
  // anchors) and all element ids update; the renderer must replace its RenderSlide with this,
  // or old sourceIds dangle and later edits silently fail
  /** Table style edit (works for tables created by this app or already modeled); returns the updated page */
  /** sourceId: the table's new element id after reparse (the renderer uses it to keep the selection), null on error */
  editTableStyle: (
    op: EditTableStyleOp,
  ) => Promise<{ slide: RenderSlide; sourceId: string | null } | null>
  /** Chart edit (charts created by this app, rebuilds the chart part); returns the updated page */
  /** sourceId: the chart's new element id after reparse (the renderer uses it to keep the selection), null on error */
  editChart: (op: EditChartOp) => Promise<{ slide: RenderSlide; sourceId: string | null } | null>
  /** Theme-derived chart color schemes (colorful + mono gradients; pass key to EditChartOp.colorScheme) */
  getChartColorSchemes: () => Promise<Array<{
    key: string
    label: string
    colors: string[]
  }> | null>
  /** Read the chart's current data (dialog echo) */
  getChartData: (
    slideIndex: number,
    sourceId: string,
  ) => Promise<{
    kind: string
    title: string
    categories: string[]
    series: Array<{ name: string; values: number[] }>
    seriesColors: Array<string | undefined>
    pointColors: Array<Array<string | undefined> | undefined>
  } | null>
  /** Export as images: shows the directory picker dialog, cancel returns null */
  pickExportDir: () => Promise<string | null>
  /** Write each page PNG to disk as <baseName>-01.png …; returns the written paths */
  exportImages: (op: ExportImagesOp) => Promise<ExportImagesResult>
  /** Export as PDF: shows the save dialog for the target path, cancel returns null */
  pickExportPdfPath: (defaultName: string) => Promise<string | null>
  /** Main process printToPDF via a hidden window, written to disk */
  exportPdf: (op: ExportPdfOp) => Promise<ExportPdfResult>
  /** Print (system dialog; cancel counts as ok=false without an error) */
  printSlides: (op: PrintSlidesOp) => Promise<{ ok: boolean; error?: string }>
  save: () => Promise<{ ok: boolean; path?: string; error?: string; slides?: RenderSlide[] }>
  saveAs: (
    defaultName: string,
  ) => Promise<{ ok: boolean; path?: string; error?: string; slides?: RenderSlide[] }>
  /** The close guard chose "Save": the main process asks the renderer to run the full save flow */
  onCloseSaveRequest: (handler: () => void) => () => void
  reportCloseSaveResult: (ok: boolean) => void
  /** Mirror the autosave toggle state to the main process: files with it on save silently on close, no dialog */
  setAutoSavePref: (on: boolean) => void
  isDirty: () => Promise<boolean>
  getRecentFiles: () => Promise<string[]>
  onMenuCommand: (handler: (command: MenuCommand) => void) => () => void
  onOpened: (handler: (result: OpenResult) => void) => () => void
  /** The file was renamed externally (shell Home list rename) — pushes the new path, the renderer updates the title bar */
  onRenamed: (handler: (newPath: string) => void) => () => void
  getAiSettings: () => Promise<AiSettings>
  setAiSettings: (settings: AiSettings) => Promise<void>
  aiStream: (request: AiStreamRequest) => Promise<void>
  aiStreamCancel: (requestId: string) => Promise<void>
  /** Genspark account status (gsk login state); with withEmail also fetches the email (needs a network request, slower) */
  aiGskStatus: (withEmail?: boolean) => Promise<GenSparkAccountStatus>
  /** Open the browser to log into Genspark (fire-and-forget; aiGskStatus turns logged-in once done) */
  aiGskLogin: () => Promise<void>
  webSearch: (
    query: string,
    maxResults?: number,
  ) => Promise<{
    results: Array<{ title: string; url: string; snippet: string }>
    answer?: string
    method: string
  }>
  imageSearch: (
    query: string,
    maxResults?: number,
  ) => Promise<{
    images: Array<{
      title: string
      imageUrl: string
      sourceUrl: string
      source: string
      width?: number
      height?: number
    }>
    method: string
  }>
  insertImageUrl: (op: {
    slideIndex: number
    url: string
    xPx: number
    yPx: number
    wPx: number
    hPx: number
    fitWidthPx: number
  }) => Promise<{ slide: RenderSlide; sourceId: string } | null>
  /** gsk (Genspark) AI image generation/editing, returns the image URL (error prompts login when logged out) */
  generateImage: (op: {
    prompt: string
    model?: string
    referenceImageUrls?: string[]
    aspectRatio?: string
    imageSize?: string
  }) => Promise<{ url?: string; error?: string }>
  /** gsk (Genspark) media analysis: image/audio/video content understanding, returns analysis text */
  analyzeMedia: (op: {
    mediaUrls: string[]
    requirements: string
  }) => Promise<{ text?: string; error?: string }>
  /** gsk availability: installed and logged in (for UI/tools to prompt login) */
  gskStatus: () => Promise<{ available: boolean; email?: string }>
  onAiStream: (handler: (chunk: AiStreamChunk) => void) => () => void
  /** Style Skill sidecar: write styleSkill to a same-named .styleskill.json next to the draft */
  saveStyleSidecar: (data: {
    topic: string
    styleSkill: string
    createdAt: string
  }) => Promise<{ ok: boolean }>
  /** Store styleSkill in userData/style-templates/<name>.json */
  saveStyleTemplate: (
    name: string,
    data: { topic: string; styleSkill: string; createdAt: string },
  ) => Promise<{ ok: boolean; error?: string }>
  /** List saved Style templates */
  listStyleTemplates: () => Promise<Array<{ name: string; topic: string; createdAt: string }>>
  /** Load a given Style template's content */
  loadStyleTemplate: (
    name: string,
  ) => Promise<{ ok: boolean; styleSkill?: string; topic?: string; error?: string }>
  /** New blank page (with a specific layout): inserted after slide sourceIndex, rels pointing at the chosen layout */
  addSlideWithLayout: (
    op: AddSlideWithLayoutOp,
  ) => Promise<{ slides: RenderSlide[]; index: number } | null>
  /** Query the current pptx's slideLayout list (for the new-slide dropdown panel) */
  getLayouts: () => Promise<GetLayoutsResult | null>
  // ── Master edit view ──────────────────────────────────────────────
  /** Enter master view: returns render trees of [master, ...layouts]; items[0] is the current edit target */
  masterEnter: (fitWidthPx: number) => Promise<MasterEnterResult | null>
  /** Switch the edit target part (re-parses that part) */
  masterOpen: (partPath: string) => Promise<RenderSlide | null>
  /** Exit master view: returns the full RenderSlide set rebuilt along the new inheritance chain */
  masterClose: () => Promise<RenderSlide[] | null>
  masterEditText: (op: MasterEditTextOp) => Promise<RenderSlide | null>
  masterEditTransform: (op: MasterEditTransformOp) => Promise<RenderSlide | null>
  masterEditFill: (op: MasterEditFillOp) => Promise<RenderSlide | null>
  masterEditStroke: (op: MasterEditStrokeOp) => Promise<RenderSlide | null>
  masterDeleteElement: (op: MasterDeleteElementOp) => Promise<RenderSlide | null>
  // ── Presenter-view multi-screen show ────────────────────────────────
  /** Enter presenter view: detects multiple displays and opens a fullscreen audience show window on the external screen (sharing this session's document) */
  presenterStart: () => Promise<{ audience: boolean }>
  /** Broadcast show state to the audience window (fire-and-forget) */
  presenterSync: (state: ShowSyncState) => void
  /** Broadcast ink/laser events to the audience window */
  presenterInk: (ev: ShowInkEvent) => void
  /** Swap the displays of the presenter/audience windows; returns false with no audience window or same screen */
  presenterSwap: () => Promise<boolean>
  /** Exit presenter view: close the audience window */
  presenterEnd: () => Promise<void>
  /** Audience window: fetch the presenter's most recent sync state (re-sent when mounting after the broadcast) */
  audienceReady: () => Promise<ShowSyncState | null>
  /** Audience window: send navigation actions back to the presenter */
  audienceNav: (action: AudienceNavAction) => void
  /** Audience window: subscribe to presenter sync state */
  onShowSync: (handler: (state: ShowSyncState) => void) => () => void
  /** Audience window: subscribe to ink/laser events */
  onShowInk: (handler: (ev: ShowInkEvent) => void) => () => void
  /** Presenter: subscribe to navigation actions sent back by the audience window */
  onAudienceNav: (handler: (action: AudienceNavAction) => void) => () => void
}

declare global {
  interface Window {
    slidesApi: SlidesApi
    desktop: DesktopFilesApi
  }
}
