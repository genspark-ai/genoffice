/**
 * Shared context for action modules extracted from App.tsx.
 *
 * App builds this bundle fresh on every render (`ctxRef.current = {...}`), and
 * exposes each extracted action as a stable `useCallback(() => impl(ctxRef.current, ...))`
 * wrapper, so module functions always see the latest state without stale closures.
 */
import type { PathCmd } from './edit-points'
import type React from 'react'
import type { GroupRenderNode, RenderNode, RenderSlide } from '@genoffice/pptx-render'
import type {
  AnimationItem,
  GetLayoutsResult,
  LinkTargetOp,
  MasterPartItem,
  PasteSlideMode,
  SectionInfo,
  TransitionKind,
} from '../shared/ipc'
import type { BrushFormat } from './format-brush'
import type { InkTool } from './ink'
import type { SlidesViewMode } from './components/ribbon-shared'
import type { CustomShow } from './slideshow-utils'

type Set<T> = React.Dispatch<React.SetStateAction<T>>

// ── State payload types shared between App useState declarations and the ctx ──

/** Viewport point the caret lands on when editing starts; select: 'word' picks the word there (double-click). */
export interface EditCaret {
  x: number
  y: number
  select?: 'word'
}

export interface EditingState {
  sourceId: string
  caret?: EditCaret
  groupId?: string
  replaceWith?: string
  /** Open with the whole text selected (F2, WordArt placeholder) */
  selectAll?: boolean
  /** Fresh click-to-type text box: leaving it without typing deletes it (PowerPoint) */
  discardIfEmpty?: boolean
}

export interface EditingCellState {
  sourceId: string
  row: number
  col: number
}

export interface SlideShowState {
  startAt: number
  customOrder?: number[]
  rehearse?: boolean
}

export type CtxMenuState =
  | { kind: 'element'; x: number; y: number; targetId: string; cell?: { row: number; col: number } }
  | { kind: 'canvas'; x: number; y: number }
  /** Inside the text edit overlay; collapsed = caret only (nothing to cut/copy) */
  | { kind: 'text'; x: number; y: number; collapsed: boolean }
  | { kind: 'thumb'; x: number; y: number; index: number }
  /** Blank space of the thumbnail rail / sorter: pos = insertion point, 0..slides.length */
  | { kind: 'gap'; x: number; y: number; pos: number }
  /** Section header; sectionId null = the unsectioned lead group */
  | { kind: 'section'; x: number; y: number; sectionId: string | null }
  | null

export interface CropTargetState {
  sourceId: string
  /** Current (possibly cropped) on-screen box */
  box: { x: number; y: number; w: number; h: number }
  srcRect?: { l: number; t: number; r: number; b: number }
  /** Full original-image extent derived from box + srcRect: the crop frame may
   * expand back out to this box, restoring toward the uncropped original */
  fullBox: { x: number; y: number; w: number; h: number }
  /** Source bitmap for the ghost preview of the original beyond the crop */
  dataUrl?: string
}

/** Edit Points mode on one top-level shape; vertex = index into vertices(shapeToEditablePath(node)) */
export interface EditPointsState {
  sourceId: string
  vertex: number | null
  /** A vertex/control drag is in progress: its release commit would resurrect a deleted vertex */
  dragging?: boolean
}

export interface CutoutTargetState {
  sourceId: string
  dataUrl: string
}

export interface LinkDialogState {
  /** Element to link (element mode); null in run mode */
  sourceId: string | null
  initial: LinkTargetOp | null
  /** Apply to the text-edit selection instead of an element (run-level link) */
  run?: boolean
}

/**
 * Former members of an ungrouped group; Regroup re-creates it from whichever still exist.
 * Keyed by the slide part path, not the index: element ids are only unique within a slide.
 */
export interface UngroupedSet {
  slidePath: string
  sourceIds: string[]
}

export interface HfDialogState {
  footer: string | null
  slideNum: boolean
  date: string | null
}

export interface ChartDataDialogInit {
  kind: string
  title: string
  categories: string[]
  series: Array<{ name: string; values: number[] }>
}

/** App state bundle passed to every extracted action; refreshed each render. */
export interface ActionCtx {
  // Document
  slides: RenderSlide[]
  setSlides: Set<RenderSlide[]>
  current: number
  setCurrent: Set<number>
  /** Thumbnail rail multi-selection: ascending, always contains current (the anchor shown on the canvas) */
  selectedSlides: number[]
  setSelectedSlides: Set<number[]>
  /** slides[current] */
  slide: RenderSlide | undefined
  path: string | null
  setPath: Set<string | null>
  setDirty: Set<boolean>
  setStatus: Set<string>
  images: Map<string, HTMLImageElement>

  // Selection / editing
  selectedIds: string[]
  setSelectedIds: Set<string[]>
  editing: EditingState | null
  setEditing: Set<EditingState | null>
  editingCell: EditingCellState | null
  setEditingCell: Set<EditingCellState | null>
  enteredGroupId: string | null
  setEnteredGroupId: Set<string | null>
  enteredGroupNode: GroupRenderNode | null
  selectedNode: RenderNode | null
  ungroupedSets: UngroupedSet[]
  setUngroupedSets: Set<UngroupedSet[]>

  // Clipboard / format painter
  setHasClipboard: Set<boolean>
  hasClipboard: boolean
  canPasteSlide: boolean
  setCanPasteSlide: Set<boolean>
  /** Paste-options floater on the just-pasted slide's thumbnail (PowerPoint-style) */
  pasteFloater: { index: number; mode: PasteSlideMode } | null
  setPasteFloater: Set<{ index: number; mode: PasteSlideMode } | null>
  brushFormat: BrushFormat | null
  setBrushFormat: Set<BrushFormat | null>
  brushMode: 'once' | 'continuous' | null
  setBrushMode: Set<'once' | 'continuous' | null>
  /** Freehand ink tool ('select' = not drawing); Esc drops back to select */
  inkTool: InkTool
  setInkTool: Set<InkTool>
  viewMode: SlidesViewMode

  // Animations / transition
  animations: AnimationItem[]
  setAnimations: Set<AnimationItem[]>
  selAnim: number
  setSelAnim: Set<number>
  setHoverAnim: Set<{ nonce: number; items: AnimationItem[] } | null>
  setTransition: Set<TransitionKind>
  animByParagraph: boolean
  timingIdx: number

  // Show / presenter / rehearse
  slideShow: SlideShowState | null
  setSlideShow: Set<SlideShowState | null>
  presenter: { startAt: number } | null
  setPresenter: Set<{ startAt: number } | null>
  setCustomShows: Set<CustomShow[]>
  setCustomShowDlgOpen: Set<boolean>
  pendingRehearse: number[] | null
  setPendingRehearse: Set<number[] | null>

  // Sections
  sections: SectionInfo[]
  setSections: Set<SectionInfo[]>
  renamingSec: { id: string; value: string } | null
  setRenamingSec: Set<{ id: string; value: string } | null>
  setCollapsedSecs: Set<globalThis.Set<string>>

  // Menus / dialogs / overlays
  ctxMenu: CtxMenuState
  setCtxMenu: Set<CtxMenuState>
  cropTarget: CropTargetState | null
  setCropTarget: Set<CropTargetState | null>
  /** Some overlay (menu, dialog, popover, crop, media, draw mode) owns Escape */
  overlayOpen: boolean
  cutoutTarget: CutoutTargetState | null
  setCutoutTarget: Set<CutoutTargetState | null>
  editPointsTarget: EditPointsState | null
  setEditPointsTarget: Set<EditPointsState | null>
  /** Edit Points commit (box-local px path); preview follows the adjust-handle gesture semantics */
  commitEditPoints: (
    sourceId: string,
    path: { w: number; h: number; cmds: PathCmd[] },
    preview: boolean,
  ) => void
  linkDialog: LinkDialogState | null
  setLinkDialog: Set<LinkDialogState | null>
  setHfDialog: Set<HfDialogState | null>
  setEqDialogOpen: Set<boolean>
  setChartDataDialogInit: Set<ChartDataDialogInit | null>
  setChartDataDialogOpen: Set<boolean>
  setFindOpen: Set<boolean>
  setPrintDlgOpen: Set<boolean>
  /** Open the AI annotation popover on the current selection (no-op when nothing is selected) */
  openAskPopover: () => void
  zoom: number
  setZoom: Set<number>
  masterItems: MasterPartItem[] | null
  /** slideLayout list of the open deck (null until fetched) */
  layouts: GetLayoutsResult['layouts'] | null
  showRuler: boolean
  showGrid: boolean
  showGuides: boolean
  toggleRuler: () => void
  toggleGrid: () => void
  toggleGuides: () => void

  // Screen recording
  recorderRef: { current: { rec: MediaRecorder; stream: MediaStream } | null }
  setRecording: Set<boolean>

  // Text-edit flush support
  editingActiveRef: { current: boolean }

  // App-retained helpers (stable or latest-bound in App)
  applySlide: (slideIndex: number, updated: RenderSlide) => void
  flushNotes: () => Promise<void>
  findNodeCtx: (id: string) => { node: RenderNode; groupId?: string } | null
  groupIdOf: (id: string) => string | undefined
  startEdit: (sourceId: string, caret?: { x: number; y: number }) => void
  undo: () => Promise<void>
  redo: () => Promise<void>
  onTransform: (
    sourceId: string,
    box: { x: number; y: number; w: number; h: number; rotationDeg: number },
    preview?: boolean,
    groupId?: string,
  ) => Promise<void>
  /** Open the format-background pane (canvas context menu entry) */
  openBgFormat: () => void
  /** Open the format pane for the selection; 'size' lands on its Size & Properties tab */
  openFormat: (section?: 'size') => void
  /** Open the comments pane with the composer focused */
  newComment: () => void
  /** Open the "Change Shape" gallery popover for a shape (context menu entry) */
  openChangeShape: (targetId: string, x: number, y: number) => void
}
