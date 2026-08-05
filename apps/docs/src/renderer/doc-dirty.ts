/**
 * Composite "has unsaved changes" check shared by the close guard, the autosave
 * tick and the crash-recovery push. Only persisted state counts — transient UI
 * state (AI highlights, selection, view modes) must never appear here.
 */
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
    s.protectionDirty
  )
}
