/**
 * Shared context for action modules extracted from App.tsx.
 *
 * App builds this bundle fresh on every render (`ctxRef.current = {...}`), and
 * exposes each extracted action as a stable `useCallback(() => impl(ctxRef.current, ...))`
 * wrapper, so module functions always see the latest state without stale closures.
 */
import type React from 'react'
import type { GroupRenderNode, RenderNode, RenderSlide } from '@genoffice/pptx-render'
import type {
  AnimationItem,
  LinkTargetOp,
  MasterPartItem,
  PasteSlideMode,
  SectionInfo,
  TransitionKind,
} from '../shared/ipc'
import type { BrushFormat } from './format-brush'
import type { CustomShow } from './slideshow-utils'

type Set<T> = React.Dispatch<React.SetStateAction<T>>

// ── State payload types shared between App useState declarations and the ctx ──

export interface EditingState {
  sourceId: string
  caret?: { x: number; y: number }
  groupId?: string
  replaceWith?: string
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
  | { kind: 'thumb'; x: number; y: number; index: number }
  | { kind: 'section'; x: number; y: number; sectionId: string }
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

  // Menus / dialogs / overlays
  ctxMenu: CtxMenuState
  setCtxMenu: Set<CtxMenuState>
  cropTarget: CropTargetState | null
  setCropTarget: Set<CropTargetState | null>
  cutoutTarget: CutoutTargetState | null
  setCutoutTarget: Set<CutoutTargetState | null>
  linkDialog: LinkDialogState | null
  setLinkDialog: Set<LinkDialogState | null>
  setHfDialog: Set<HfDialogState | null>
  setEqDialogOpen: Set<boolean>
  setChartDataDialogInit: Set<ChartDataDialogInit | null>
  setChartDataDialogOpen: Set<boolean>
  setFindOpen: Set<boolean>
  setPrintDlgOpen: Set<boolean>
  setZoom: Set<number>
  masterItems: MasterPartItem[] | null

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
}
