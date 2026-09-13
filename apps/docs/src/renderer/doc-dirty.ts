/**
 * Composite "has unsaved changes" check shared by the close guard, the autosave
 * tick and the crash-recovery push. Only persisted state counts — transient UI
 * state (AI highlights, selection, view modes) must never appear here.
 */
import type { HeaderFooter, SectionInfo, StyleUpsert } from '@genoffice/docx-engine'

import type { PendingNumbering } from './doc-state'
export interface DocDirtyState {
  dirtyRef: { current: boolean }
  sectionDirty: boolean
  sectionsDirty: readonly number[]
  trailingStartType: unknown
  pageColorDirty: boolean
  headerDirty: boolean
  footerDirty: boolean
  hfVariantsDirty: readonly unknown[]
  sectionHfEdits: Record<string, unknown>
  pgNumEdit: unknown
  pgNumDirtySections: readonly number[]
  numberingDirty: boolean
  styleUpserts: Record<string, unknown>
  titlePgDirty: boolean
  evenOddHfDirty: boolean
  watermarkDirty: boolean
  inksDirty: boolean
  notesDirty: boolean
  sourcesDirty: boolean
  themeFontsDirty: boolean
  themeColorsDirty: boolean
  commentsDirty: boolean
  protectionDirty: boolean
  writeProtectionDirty: boolean
  removePersonalInfoDirty: boolean
}

export function isDocDirty(s: DocDirtyState): boolean {
  return (
    s.dirtyRef.current ||
    s.sectionDirty ||
    s.sectionsDirty.length > 0 ||
    s.trailingStartType !== null ||
    s.pageColorDirty ||
    s.headerDirty ||
    s.footerDirty ||
    s.hfVariantsDirty.length > 0 ||
    Object.keys(s.sectionHfEdits).length > 0 ||
    s.pgNumEdit !== null ||
    s.pgNumDirtySections.length > 0 ||
    s.numberingDirty ||
    Object.keys(s.styleUpserts).length > 0 ||
    s.titlePgDirty ||
    s.evenOddHfDirty ||
    s.watermarkDirty ||
    s.inksDirty ||
    s.notesDirty ||
    s.sourcesDirty ||
    s.themeFontsDirty ||
    s.themeColorsDirty ||
    s.commentsDirty ||
    s.protectionDirty ||
    s.writeProtectionDirty ||
    s.removePersonalInfoDirty
  )
}

/** The edit-tracking setters resetCrossDocEditState clears (structural subset
 *  of FileActionContext, so the helper stays unit-testable without the editor). */
export interface CrossDocEditStateSink {
  setSectionsDirty: (value: number[]) => void
  setTrailingStartType: (value: SectionInfo['startType'] | null) => void
  setSectionHfEdits: (value: Record<string, HeaderFooter>) => void
  setPgNumEdit: (value: { fmt?: string; start?: number } | null) => void
  setPgNumDirtySections: (value: number[]) => void
  setPendingNumbering: (value: PendingNumbering) => void
  setStyleUpserts: (value: Record<string, StyleUpsert>) => void
}

/**
 * Clear the section/numbering/style edit-tracking states. The save path calls
 * this once the bytes land; a document swap (open/new) must call it too, or
 * doc A's edits leak into pristine doc B — tripping the close guard and
 * mis-applying section indices, numbering restarts and style upserts on B's
 * next save. Single source of truth so the call sites cannot drift apart.
 */
export function resetCrossDocEditState(sink: CrossDocEditStateSink): void {
  sink.setSectionsDirty([])
  sink.setTrailingStartType(null)
  sink.setSectionHfEdits({})
  sink.setPgNumEdit(null)
  sink.setPgNumDirtySections([])
  sink.setPendingNumbering({ newDefs: [], restartNums: [] })
  sink.setStyleUpserts({})
}
