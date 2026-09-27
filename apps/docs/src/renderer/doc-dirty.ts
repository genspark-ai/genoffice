/**
 * Composite "has unsaved changes" check shared by the close guard, the autosave
 * tick and the crash-recovery push. Only persisted state counts — transient UI
 * state (AI highlights, selection, view modes) must never appear here.
 */
import type { DefaultFonts, HeaderFooter, SectionInfo, StyleUpsert } from '@genoffice/docx-engine'

import { EMPTY_PENDING_NUMBERING } from './doc-state'
import type { PendingNumbering } from './doc-state'

export async function runGuardedDocumentAction(
  confirm: () => Promise<boolean>,
  action: () => void | Promise<unknown>,
): Promise<boolean> {
  if (!(await confirm())) return false
  await action()
  return true
}

/**
 * Replace the open document with a candidate the user still has to pick (file
 * picker, recent entry). The guard runs only once a candidate exists: confirming
 * may save the current document or drop its recovery copy, so a cancelled picker
 * must never trigger it. `needsGuard` lets a caller defer the guard for a
 * candidate that is not a document yet (an encrypted file waiting for its
 * password); that caller guards itself once the real document is in hand.
 */
export async function runGuardedCandidate<T>(
  confirm: () => Promise<boolean>,
  choose: () => Promise<T | null | undefined>,
  commit: (candidate: T) => void | Promise<unknown>,
  needsGuard: (candidate: T) => boolean = () => true,
): Promise<boolean> {
  const candidate = await choose()
  if (candidate == null) return false
  if (needsGuard(candidate) && !(await confirm())) return false
  await commit(candidate)
  return true
}

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
  hfLinks: Record<string, unknown>
  pgNumEdit: unknown
  pgNumDirtySections: readonly number[]
  numberingDirty: boolean
  defaultFonts?: DefaultFonts
  styleUpserts: Record<string, unknown>
  titlePgDirty: boolean
  evenOddHfDirty: boolean
  mirrorMarginsDirty: boolean
  watermarkDirty: boolean
  inksDirty: boolean
  notesDirty: boolean
  sourcesDirty: boolean
  zoteroDocumentDataDirty: boolean
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
    Object.keys(s.hfLinks).length > 0 ||
    s.pgNumEdit !== null ||
    s.pgNumDirtySections.length > 0 ||
    s.numberingDirty ||
    Object.keys(s.styleUpserts).length > 0 ||
    s.defaultFonts !== undefined ||
    s.titlePgDirty ||
    s.evenOddHfDirty ||
    s.mirrorMarginsDirty ||
    s.watermarkDirty ||
    s.inksDirty ||
    s.notesDirty ||
    s.sourcesDirty ||
    s.zoteroDocumentDataDirty ||
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
  setHfLinks: (value: Record<string, true>) => void
  setPgNumEdit: (value: { fmt?: string; start?: number } | null) => void
  setPgNumDirtySections: (value: number[]) => void
  setPendingNumbering: (value: PendingNumbering) => void
  setDefaultFonts?: (fonts: DefaultFonts | undefined) => void
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
  sink.setHfLinks({})
  sink.setPgNumEdit(null)
  sink.setPgNumDirtySections([])
  sink.setPendingNumbering(EMPTY_PENDING_NUMBERING)
  sink.setStyleUpserts({})
  sink.setDefaultFonts?.(undefined)
}
