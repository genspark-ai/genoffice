import type { LocalMarkup } from './annotations'
import type { LocalDrawing } from './DrawLayer'
import type { LocalImageEdit } from './ImageEditLayer'
import type { SavedNoteAnnot } from './note-threads'
import type { HeaderFooterConfig, WatermarkConfig } from './stamps'
import type { LocalTextEdit, LocalTextInsert } from './text-edit-preview'
import type { FormValueInput, MarkupType, MetadataInput, PageImageRef } from '../shared/ipc'

/** pdf.js AnnotationType codes for the markup subtypes we can delete */
export const MARKUP_TYPE_BY_ANNOT: Record<number, MarkupType> = {
  9: 'highlight',
  10: 'underline',
  12: 'strikeout',
}

export const rectsNear = (a: readonly number[], b: readonly number[], tolerance = 2): boolean =>
  a.length === 4 &&
  b.length === 4 &&
  a.every((value, index) => Math.abs(value - b[index]!) <= tolerance)

/** Full snapshot of unsaved edits (for undo/redo; data is small, whole-copy replace is safest) */
/** Watermark/header-footer are kept as config and rendered in final page order only at save time, so page numbers survive reorders/deletions */
export interface StampConfig {
  wm: WatermarkConfig | null
  hf: HeaderFooterConfig | null
}

/** A markup annotation already saved in the file (read via pdf.js getAnnotations) */
export interface SavedMarkupAnnot {
  pageIndex: number
  /** PDF object number (pdf.js id "123R" → 123) */
  objNum: number
  type: MarkupType
  quads: number[][]
  rect: [number, number, number, number]
}

/** Pending deletion of a saved markup or note annotation */
export interface LocalAnnotDelete {
  id: string
  annot: SavedMarkupAnnot | SavedNoteAnnot
}

/** Pending in-place content edit of a note comment saved in the file.
    `annot.contents` stays the on-disk text — the save path matches by it. */
export interface LocalNoteEdit {
  id: string
  annot: SavedNoteAnnot
  contents: string
}

export interface EditSnapshot {
  markups: LocalMarkup[]
  annotDeletes: LocalAnnotDelete[]
  noteEdits: LocalNoteEdit[]
  drawings: LocalDrawing[]
  textEdits: LocalTextEdit[]
  textInserts: LocalTextInsert[]
  imageEdits: LocalImageEdit[]
  stampCfg: StampConfig | null
  formEdits: Map<string, FormValueInput>
  rotations: Map<number, number>
  deleted: Set<number>
  order: number[] | null
  metadata: MetadataInput | null
}

/** What a running save wrote, captured when the save starts. The post-save reload
    subtracts exactly this instead of wiping all edit state, so anything the user did
    while the write was in flight stays pending on the reloaded document. */
export interface SavedSnapshot {
  markupIds: Set<string>
  annotDeleteIds: Set<string>
  noteEditIds: Set<string>
  /** objNum → /Contents this save wrote via noteEdits. A re-edit of the same note made
      while the write was in flight still carries the pre-save on-disk text as its match
      key; the post-save subtraction retargets it to what the file now says. */
  noteEditWritten: Map<number, string>
  drawingIds: Set<string>
  textEditIds: Set<string>
  textInsertIds: Set<string>
  imageEditIds: Set<string>
  stampCfg: StampConfig | null
  formEdits: Map<string, FormValueInput>
  rotations: Map<number, number>
  metadata: MetadataInput | null
  /** Old original page index → its index in the saved file (saved deletions/reorder applied) */
  pageMap: Map<number, number>
}

/** Selected annotation with the anchor of its floating delete popup; a stamp click selects the whole watermark/header-footer set */
export type AnnotSelection =
  | {
      kind: 'markup' | 'drawing' | 'textEdit' | 'textInsert' | 'imageEdit'
      id: string
      x: number
      y: number
    }
  | { kind: 'savedMarkup'; annot: SavedMarkupAnnot; x: number; y: number }
  | { kind: 'pageImage'; ref: PageImageRef; x: number; y: number }
  | { kind: 'stamp'; x: number; y: number }
